using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Win32;
using System.Windows.Forms;
using Mutagen.Bethesda.Plugins;
using Mutagen.Bethesda.Plugins.Records;
using Mutagen.Bethesda.Skyrim;

namespace MissingMasterCleaner
{
    internal class Program
    {
        // ---- CONFIG ----
        // This is only used for the "NEXT STEPS" message.
        // The tool does NOT install/copy the script.
        private const string PascalScriptFileNameWithoutExtension = "MissingMasterCleaner_ApplyJsonRemovals";

        // Logs live next to the executable.
        static readonly string LogPath = Path.Combine(AppContext.BaseDirectory, "MissingMasterCleaner.log");
        static readonly StringBuilder Log = new();

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                LogLine("Unexpected error:");
                LogLine(ex.ToString());
                Exit();
            }
        }

        static void Run()
        {
            LogLine("=== Missing Master Cleaner – JSON Reporting Phase ===");
            LogLine($"Started: {DateTime.Now}");
            LogLine("");

            var dataDir = DetectDataDirectory();
            if (dataDir == null)
                Exit();

            LogLine("Scanning plugins, please be patient!");
            Console.Write("Working...");

            var allPlugins = Directory.GetFiles(dataDir)
                .Where(f =>
                    f.EndsWith(".esp", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".esm", StringComparison.OrdinalIgnoreCase) ||
                    f.EndsWith(".esl", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => Path.GetFileName(f))
                .ToList();

            string[] vanilla =
            {
                "Skyrim.esm",
                "Update.esm",
                "Dawnguard.esm",
                "HearthFires.esm",
                "Dragonborn.esm"
            };

            var unsafePlugins = new List<(string Path, List<ModKey> MissingMasters)>();
            var safePlugins = new List<(string Path, ModKey MissingMaster)>();
            int iPass = 0;

            foreach (var pluginPath in allPlugins)
            {
                iPass++;
                if (iPass % 20 == 0)
                    Console.Write(".");

                string fileName = Path.GetFileName(pluginPath);
                if (vanilla.Contains(fileName, StringComparer.OrdinalIgnoreCase))
                    continue;

                SkyrimMod mod;
                try
                {
                    mod = SkyrimMod.CreateFromBinary(pluginPath, SkyrimRelease.SkyrimSE);
                }
                catch (Exception ex)
                {
                    LogLine("");
                    LogLine($"ERROR loading {fileName}: {ex.Message}");
                    LogLine("The error above is usually caused by a corrupted plugin. Contact the mod author and report it.");
                    continue;
                }

                var missingMasters = mod.ModHeader.MasterReferences
                    .Where(m => !File.Exists(Path.Combine(dataDir, m.Master.FileName)))
                    .Select(m => m.Master)
                    .ToList();

                if (missingMasters.Count > 1)
                    unsafePlugins.Add((pluginPath, missingMasters));
                else if (missingMasters.Count == 1)
                    safePlugins.Add((pluginPath, missingMasters[0]));
            }

            if (unsafePlugins.Count > 0)
            {
                LogLine("");
                LogLine("");
                LogLine("The following plugins have more than 1 missing master and can ONLY be cleaned manually due to safety concerns:");
                foreach (var u in unsafePlugins)
                {
                    string list = string.Join(", ", u.MissingMasters.Select(m => m.FileName));
                    LogLine($" - {Path.GetFileName(u.Path)} | Missing masters: {list}");
                }
                LogLine("");
            }

            if (safePlugins.Count == 0)
            {
                LogLine("No plugins with exactly 1 missing master were found.");
                Exit();
            }

            safePlugins = safePlugins
                .OrderBy(p => Path.GetFileName(p.Path))
                .ToList();

            LogLine("");
            LogLine("");
            LogLine($"{safePlugins.Count} plugin(s) found with 1 missing master:");
            LogLine("");

            for (int i = 0; i < safePlugins.Count; i++)
                LogLine($"[{i + 1}] {Path.GetFileName(safePlugins[i].Path)}");

            LogLine("");
            LogLine("Type the number corresponding to the line above that has the missing master you want to clean:");

            int selection;
            while (true)
            {
                Console.Write("> ");
                if (int.TryParse(Console.ReadLine(), out selection) &&
                    selection >= 1 &&
                    selection <= safePlugins.Count)
                    break;

                Console.WriteLine("Invalid selection.");
            }

            var chosen = safePlugins[selection - 1];
            LogLine("");
            LogLine($"Selected plugin: {Path.GetFileName(chosen.Path)}");
            LogLine($"Missing master: {chosen.MissingMaster.FileName}");
            LogLine("");

            // 1) Create dummy master in Data dir (MO2 redirects to Overwrite; Vortex deploys to Data)
            string dummyTargetPath = CreateDummyMasterInDataDir(dataDir, chosen.MissingMaster);
            LogLine($"Dummy master write target (Data dir / MO2 Overwrite): {dummyTargetPath}");

            // 2) Create ONE consolidated JSON into Data dir (MO2 redirects to Overwrite)
            Console.WriteLine("");
            Console.WriteLine("Scanning selected plugin deeply (this may take a while). Please be patient...");

            string jsonPath = GenerateJsonReport(chosen.Path, chosen.MissingMaster, dataDir);
            LogLine($"JSON report written to (Data dir / MO2 Overwrite): {jsonPath}");

            // User messaging (as requested; does not auto-start xEdit)
            LogLine("");
            LogLine("NEXT STEPS:");
            LogLine("1) Run SSEEdit, and be sure to load the plugin with the missing master.");
            LogLine($"2) In SSEEdit, right-click the plugin and Apply the script: {PascalScriptFileNameWithoutExtension} (must be installed in SSEEdit\\Edit Scripts).");
            LogLine("3) Save when prompted.");

            Exit();
        }

        // -------------------------------------------------
        // Dummy master → write into Data dir (MO2 redirects to Overwrite)
        // -------------------------------------------------
        static string CreateDummyMasterInDataDir(string dataDir, ModKey missingMaster)
        {
            string dummyPath = Path.Combine(dataDir, missingMaster.FileName);

            if (File.Exists(dummyPath))
                return dummyPath;

            var dummy = new SkyrimMod(missingMaster, SkyrimRelease.SkyrimSE);
            dummy.ModHeader.Description = "Temporary dummy master created by MissingMasterCleaner";
            dummy.ModHeader.Author = "Glanzer";
            dummy.WriteToBinary(dummyPath);

            return dummyPath;
        }

        // -------------------------------------------------
        // CONSOLIDATED JSON REPORT (single pass)
        //
        // Output file: <Dependent>_MissingMasterReport.json
        //
        // Sections:
        //  - PlacedObjects   : records owned by missing master (safe remove)
        //  - FormLists       : safe unresolved links
        //  - LeveledLists    : safe unresolved links
        //  - Containers      : safe unresolved links
        //  - OtherLinks      : unresolved links in other records (NOT safe; blocks CleanMasters)
        // -------------------------------------------------
        static string GenerateJsonReport(string pluginPath, ModKey missingMaster, string dataDir)
        {
            SkyrimMod mod = SkyrimMod.CreateFromBinary(pluginPath, SkyrimRelease.SkyrimSE);
            string pluginFileName = Path.GetFileName(pluginPath);

            var placedObjects = new List<Dictionary<string, object?>>();
            var placedRefs = new List<Dictionary<string, object?>>();
            var safeFormLists = new List<Dictionary<string, object?>>();
            var safeLeveledLists = new List<Dictionary<string, object?>>();
            var safeContainers = new List<Dictionary<string, object?>>();
            var otherLinks = new List<Dictionary<string, object?>>();

            // PASS 1: collect ALL missing-master-owned FormKeys up front
            var safeOwnedReferencedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var record in mod.EnumerateMajorRecords())
            {
                if (record.FormKey.ModKey == missingMaster)
                    safeOwnedReferencedKeys.Add(record.FormKey.ToString());
            }

            int ownedCount = 0;
            int placedRefCount = 0;

            // "safeLinkCount" should count entries we will attempt to remove from safe lists
            int safeLinkCount = 0;

            // "otherLinkCount" should count entries in OtherLinks after filtering
            int otherLinkCount = 0;

            foreach (var record in mod.EnumerateMajorRecords())
            {
                string recordType = record.Registration.Name;
                string recordFormIdHex = record.FormKey.ID.ToString("X8");
                string recordFormKeyStr = record.FormKey.ToString();

                // A) OWNED records (safe deletions)
                if (record.FormKey.ModKey == missingMaster)
                {
                    string? displayName = TryGetDisplayName(record);

                    var ownedEntry = new Dictionary<string, object?>
                    {
                        ["RecordType"] = recordType,
                        ["ReferencedFormKey"] = recordFormKeyStr,
                        ["RecordFormIDHex"] = recordFormIdHex,
                        ["EditorID"] = record.EditorID
                    };

                    if (!string.IsNullOrWhiteSpace(displayName))
                        ownedEntry["DisplayName"] = displayName;

                    if (string.Equals(recordType, "PlacedObject", StringComparison.OrdinalIgnoreCase))
                    {
                        placedRefs.Add(ownedEntry);
                        placedRefCount++;
                    }
                    else
                    {
                        placedObjects.Add(ownedEntry);
                        ownedCount++;
                    }

                    // Don’t scan owned records for links
                    continue;
                }

                // B) Unresolved links elsewhere
                var hits = new List<LinkHit>();
                MissingMasterLinkWalker.WalkObjectGraphForMissingMasterLinks(
                    root: record,
                    missingMaster: missingMaster,
                    pathPrefix: "",
                    hits: hits,
                    maxDepth: 16);

                if (hits.Count == 0)
                    continue;

                bool isFormList = string.Equals(recordType, "FormList", StringComparison.OrdinalIgnoreCase);
                bool isContainer = string.Equals(recordType, "Container", StringComparison.OrdinalIgnoreCase);
                bool isLeveled = recordType.StartsWith("Leveled", StringComparison.OrdinalIgnoreCase);

                // IMPORTANT:
                // - For SAFE sections (FormLists/LeveledLists/Containers), we KEEP references to missing-master-owned keys,
                //   because those are precisely what we need to remove from those lists.
                // - For OtherLinks, we FILTER OUT missing-master-owned keys to avoid duplication with PlacedObjects/PlacedRefs.
                List<LinkHit> hitsForThisSection;
                if (isFormList || isContainer || isLeveled)
                {
                    hitsForThisSection = hits;
                }
                else
                {
                    hitsForThisSection = hits
                        .Where(h => !safeOwnedReferencedKeys.Contains(h.ReferencedFormKey))
                        .ToList();

                    // NEW: drop hits that are inside a removable placed ref (Persistent[n]/Temporary[n])
                    hitsForThisSection = FilterHitsInsideRemovablePlacedRefs(record, missingMaster, hitsForThisSection);
                }

                if (hitsForThisSection.Count == 0)
                    continue;

                var linkEntry = new Dictionary<string, object?>
                {
                    ["RecordType"] = recordType,
                    ["RecordFormKey"] = recordFormKeyStr,   // owning record
                    ["RecordFormIDHex"] = recordFormIdHex,
                    ["EditorID"] = record.EditorID,
                    ["Links"] = hitsForThisSection.Select(h => new Dictionary<string, object?>
                    {
                        ["Path"] = h.Path,
                        ["ReferencedFormKey"] = h.ReferencedFormKey
                    }).ToList()
                };

                if (isFormList)
                {
                    safeFormLists.Add(linkEntry);
                    safeLinkCount += hitsForThisSection.Count;
                }
                else if (isLeveled)
                {
                    safeLeveledLists.Add(linkEntry);
                    safeLinkCount += hitsForThisSection.Count;
                }
                else if (isContainer)
                {
                    safeContainers.Add(linkEntry);
                    safeLinkCount += hitsForThisSection.Count;
                }
                else
                {
                    otherLinks.Add(linkEntry);
                    otherLinkCount += hitsForThisSection.Count;
                }
            }

            var report = new
            {
                Plugin = pluginFileName,
                MissingMaster = missingMaster,
                PlacedObjects = placedObjects,
                PlacedRefs = placedRefs,
                FormLists = safeFormLists,
                LeveledLists = safeLeveledLists,
                Containers = safeContainers,
                OtherLinks = otherLinks
            };

            string jsonPath = Path.Combine(
                dataDir,
                $"{Path.GetFileNameWithoutExtension(pluginPath)}_MissingMasterReport.json");

            File.WriteAllText(
                jsonPath,
                JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));

            LogLine($"Removable records found (PlacedObjects): {ownedCount}");
            LogLine($"Removable placed references found (PlacedRefs): {placedRefCount}");
            LogLine($"Safe list references found (FormLists/LeveledLists/Containers): {safeLinkCount}");
            LogLine($"Other links found (unsafe / blocks CleanMasters): {otherLinkCount}");

            return jsonPath;
        }

        static string? TryGetDisplayName(IMajorRecordGetter record)
        {
            try
            {
                // Many Skyrim records implement ITranslatedStringGetter Name; ToString() is fine for reporting.
                var nameProp = record.GetType().GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
                if (nameProp != null)
                {
                    var val = nameProp.GetValue(record);
                    if (val != null)
                        return val.ToString();
                }
            }
            catch
            {
                // ignore
            }
            return null;
        }

        // -------------------------------------------------
        // LinkHit + Walker (optimized; builds Path only on hit)
        // -------------------------------------------------
        private sealed class LinkHit
        {
            public string Path { get; init; } = "";
            public string ReferencedFormKey { get; init; } = "";
        }

        static class MissingMasterLinkWalker
        {
            // Cache filtered public instance properties per type
            private static readonly ConcurrentDictionary<Type, PropertyInfo[]> _propsCache = new();
            private static readonly ConcurrentDictionary<Type, bool> _leafCache = new();
            private static readonly ConcurrentDictionary<Type, bool> _genericLinkGetterCache = new();

            private static readonly HashSet<string> _skipProps = new(StringComparer.Ordinal)
            {
                "Parent", "Parents", "LinkCache", "Links", "Registration", "FormVersion"
            };

            private readonly struct PathFrame
            {
                public readonly string? Name;
                public readonly int? Index;
                public readonly bool IsIndex;

                public PathFrame(string name)
                {
                    Name = name;
                    Index = null;
                    IsIndex = false;
                }

                public PathFrame(int index)
                {
                    Name = null;
                    Index = index;
                    IsIndex = true;
                }
            }

            public static void WalkObjectGraphForMissingMasterLinks(
                object root,
                ModKey missingMaster,
                string pathPrefix,
                List<LinkHit> hits,
                int maxDepth)
            {
                var visited = new HashSet<object>(ReferenceEqualityComparer.Instance);
                var frames = new List<PathFrame>(capacity: 64);
                WalkInternal(root, missingMaster, pathPrefix ?? "", hits, visited, frames, depth: 0, maxDepth: maxDepth);
            }

            private static void WalkInternal(
                object? obj,
                ModKey missingMaster,
                string prefix,
                List<LinkHit> hits,
                HashSet<object> visited,
                List<PathFrame> frames,
                int depth,
                int maxDepth)
            {
                if (obj is null)
                    return;
                if (depth > maxDepth)
                    return;

                var t = obj.GetType();
                if (IsLeafType(t))
                    return;

                if (!t.IsValueType)
                {
                    if (!visited.Add(obj))
                        return;
                }

                // FormKey (direct)
                if (obj is FormKey fk)
                {
                    if (!fk.IsNull && fk.ModKey == missingMaster)
                    {
                        hits.Add(new LinkHit
                        {
                            Path = BuildPath(prefix, frames),
                            ReferencedFormKey = fk.ToString()
                        });
                    }
                    return;
                }

                // Non-generic IFormLinkGetter
                if (obj is IFormLinkGetter fl)
                {
                    var lk = fl.FormKey;
                    if (!lk.IsNull && lk.ModKey == missingMaster)
                    {
                        hits.Add(new LinkHit
                        {
                            Path = BuildPath(prefix, frames),
                            ReferencedFormKey = lk.ToString()
                        });
                    }
                    return;
                }

                // Generic IFormLinkGetter<T> (via reflection only when needed)
                if (IsGenericFormLinkGetter(t))
                {
                    var formKeyProp = t.GetProperty("FormKey", BindingFlags.Public | BindingFlags.Instance);
                    if (formKeyProp?.GetValue(obj) is FormKey gk)
                    {
                        if (!gk.IsNull && gk.ModKey == missingMaster)
                        {
                            hits.Add(new LinkHit
                            {
                                Path = BuildPath(prefix, frames),
                                ReferencedFormKey = gk.ToString()
                            });
                        }
                    }
                    return;
                }

                // Enumerables first
                if (obj is IEnumerable en && obj is not string)
                {
                    int idx = 0;
                    foreach (var item in en)
                    {
                        frames.Add(new PathFrame(idx));
                        WalkInternal(item, missingMaster, prefix, hits, visited, frames, depth + 1, maxDepth);
                        frames.RemoveAt(frames.Count - 1);
                        idx++;
                    }
                    return;
                }

                // Property reflection (cached)
                var props = GetCachedProps(t);
                for (int i = 0; i < props.Length; i++)
                {
                    var p = props[i];
                    object? value;
                    try
                    {
                        value = p.GetValue(obj);
                    }
                    catch
                    {
                        continue;
                    }

                    if (value is null)
                        continue;

                    frames.Add(new PathFrame(p.Name));
                    WalkInternal(value, missingMaster, prefix, hits, visited, frames, depth + 1, maxDepth);
                    frames.RemoveAt(frames.Count - 1);
                }
            }

            private static bool IsLeafType(Type t)
            {
                return _leafCache.GetOrAdd(t, static tt =>
                {
                    if (tt.IsPrimitive || tt.IsEnum)
                        return true;

                    return tt == typeof(string)
                           || tt == typeof(decimal)
                           || tt == typeof(DateTime)
                           || tt == typeof(Guid);
                });
            }

            private static PropertyInfo[] GetCachedProps(Type t)
            {
                return _propsCache.GetOrAdd(t, static tt =>
                {
                    return tt.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                             .Where(p =>
                                 p.CanRead &&
                                 p.GetIndexParameters().Length == 0 &&
                                 !_skipProps.Contains(p.Name))
                             .ToArray();
                });
            }

            private static bool IsGenericFormLinkGetter(Type t)
            {
                if (typeof(IFormLinkGetter).IsAssignableFrom(t))
                    return false;

                return _genericLinkGetterCache.GetOrAdd(t, static tt =>
                {
                    return tt.GetInterfaces().Any(i =>
                        i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IFormLinkGetter<>));
                });
            }

            private static string BuildPath(string prefix, List<PathFrame> frames)
            {
                if (frames.Count == 0)
                    return string.IsNullOrWhiteSpace(prefix) ? "<root>" : prefix;

                var sb = new StringBuilder(capacity: 128);

                if (!string.IsNullOrWhiteSpace(prefix))
                    sb.Append(prefix);

                bool wroteSomething = !string.IsNullOrWhiteSpace(prefix);

                for (int i = 0; i < frames.Count; i++)
                {
                    var f = frames[i];

                    if (!f.IsIndex)
                    {
                        if (wroteSomething)
                            sb.Append('.');
                        sb.Append(f.Name);
                        wroteSomething = true;
                    }
                    else
                    {
                        sb.Append('[').Append(f.Index!.Value).Append(']');
                        wroteSomething = true;
                    }
                }

                return sb.Length == 0 ? "<root>" : sb.ToString();
            }

            private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
            {
                public static readonly ReferenceEqualityComparer Instance = new();

                public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);

                public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
            }
        }

        static List<LinkHit> FilterHitsInsideRemovablePlacedRefs(
            IMajorRecordGetter record,
            ModKey missingMaster,
            List<LinkHit> hits)
        {
            if (hits.Count == 0)
                return hits;

            static bool TryExtractIndex(string path, string token, out int idx)
            {
                idx = -1;
                int p = path.IndexOf(token, StringComparison.OrdinalIgnoreCase);
                if (p < 0) return false;

                p += token.Length;
                int end = path.IndexOf(']', p);
                if (end < 0) return false;

                return int.TryParse(path.AsSpan(p, end - p), out idx) && idx >= 0;
            }

            static bool IsPlacedRefOwnedByMissingMaster(IReadOnlyList<IPlacedGetter> list, int index, ModKey mm)
            {
                if ((uint)index >= (uint)list.Count) return false;
                return list[index].FormKey.ModKey == mm;
            }

            var result = new List<LinkHit>(hits.Count);

            foreach (var h in hits)
            {
                int pIdx = -1;
                int tIdx = -1;

                bool isPersistent = TryExtractIndex(h.Path, "Persistent[", out pIdx);
                bool isTemporary = !isPersistent && TryExtractIndex(h.Path, "Temporary[", out tIdx);

                bool skip = false;

                if (record is ICellGetter cell)
                {
                    if (isPersistent && cell.Persistent != null && IsPlacedRefOwnedByMissingMaster(cell.Persistent, pIdx, missingMaster))
                        skip = true;
                    else if (isTemporary && cell.Temporary != null && IsPlacedRefOwnedByMissingMaster(cell.Temporary, tIdx, missingMaster))
                        skip = true;
                }
                else if (record is IWorldspaceGetter ws)
                {
                    // Worldspace persistent/temporary refs hang off TopCell in SSE
                    var top = ws.TopCell;
                    if (top != null)
                    {
                        if (isPersistent && top.Persistent != null && IsPlacedRefOwnedByMissingMaster(top.Persistent, pIdx, missingMaster))
                            skip = true;
                        else if (isTemporary && top.Temporary != null && IsPlacedRefOwnedByMissingMaster(top.Temporary, tIdx, missingMaster))
                            skip = true;
                    }
                }

                if (!skip)
                    result.Add(h);
            }

            return result;
        }


        // -------------------------------------------------
        // Logging / Exit
        // -------------------------------------------------
        static void LogLine(string line)
        {
            Console.WriteLine(line);
            Log.AppendLine(line);
            File.WriteAllText(LogPath, Log.ToString());
        }

        static void Exit()
        {
            LogLine("");
            LogLine($"Log written to: {LogPath}");
            LogLine("Press any key to exit...");
            Console.ReadKey();
            Environment.Exit(0);
        }

        // -------------------------------------------------
        // Detect Skyrim SE Data directory
        // -------------------------------------------------
        static string? DetectDataDirectory()
        {
            // 1) Try registry first
            using var key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\WOW6432Node\Bethesda Softworks\Skyrim Special Edition");

            var installPath = key?.GetValue("Installed Path") as string;

            if (!string.IsNullOrWhiteSpace(installPath))
            {
                var dataDir = Path.Combine(installPath, "Data");
                if (IsValidSkyrimDataDirectory(dataDir))
                {
                    // NEW: Ask user to confirm the detected path
                    while (true)
                    {
                        LogLine($"Detected Data directory from registry: {dataDir}");
                        LogLine("Is this Data directory location correct?");
                        LogLine("1) Yes, use it");
                        LogLine("2) No, let me choose a different location");
                        LogLine("3) Exit");
                        LogLine("");

                        var choice = (Console.ReadLine() ?? "").Trim();

                        if (choice == "1")
                        {
                            LogLine($"Using Data directory: {dataDir}");
                            LogLine("");
                            return dataDir;
                        }

                        if (choice == "2")
                        {
                            LogLine("OK, please choose a different Data directory.");
                            LogLine("");
                            break; // continue to manual selection loop below
                        }

                        if (choice == "3")
                        {
                            LogLine("Exiting.");
                            Environment.Exit(0);
                            return null; // unreachable, but keeps compiler happy
                        }

                        LogLine("Invalid choice");
                        LogLine("");
                    }
                }
            }

            //LogLine("Skyrim SE Data directory not found automatically.");
            //LogLine("");

            // 2) Manual selection loop
            while (true)
            {
                var selectedPath = PromptUserForDataDirectory();

                if (selectedPath == null)
                {
                    LogLine("User cancelled. Exiting.");
                    return null;
                }

                if (IsValidSkyrimDataDirectory(selectedPath))
                {
                    LogLine($"Using Data directory: {selectedPath}");
                    LogLine("");
                    return selectedPath;
                }

                MessageBox.Show(
                    text: "The selected folder is not a valid Skyrim Data directory.\n\nIt must contain Skyrim.esm.",
                    caption: "Invalid Data Directory",
                    buttons: MessageBoxButtons.OK,
                    icon: MessageBoxIcon.Error,
                    defaultButton: MessageBoxDefaultButton.Button1,
                    options: MessageBoxOptions.DefaultDesktopOnly
                );
            }
        }
        static bool IsValidSkyrimDataDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            if (!Directory.Exists(path))
                return false;

            var skyrimEsm = Path.Combine(path, "Skyrim.esm");
            return File.Exists(skyrimEsm);
        }

        static string? PromptUserForDataDirectory()
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "Select your Skyrim Special Edition Data folder",
                UseDescriptionForTitle = true,
                ShowNewFolderButton = false
            };

            var result = dialog.ShowDialog();

            if (result == DialogResult.OK)
                return dialog.SelectedPath;

            return null; // user cancelled
        }

    }
}
