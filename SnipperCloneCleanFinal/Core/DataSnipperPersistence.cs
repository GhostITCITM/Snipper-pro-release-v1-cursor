using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace SnipperCloneCleanFinal.Core
{
    internal static class DataSnipperPersistence
    {
        private const string Prop = "SnipperSnips2";
        private static Dictionary<string, SnipData> _cache = new Dictionary<string, SnipData>();

        public static void Upsert(SnipData snip) => _cache[snip.Id] = snip;
        public static bool TryGet(string id, out SnipData snip) => _cache.TryGetValue(id, out snip);

        public static void Save(Excel.Workbook wb)
        {
            var json = JsonConvert.SerializeObject(_cache.Values);
            var props = (DocumentProperties)wb.CustomDocumentProperties;
            var existing = props.Cast<DocumentProperty>().FirstOrDefault(p => p.Name == Prop);
            existing?.Delete();
            props.Add(Prop, false, MsoDocProperties.msoPropertyTypeString, json);
        }

        public static void Load(Excel.Workbook wb)
        {
            var props = (DocumentProperties)wb.CustomDocumentProperties;
            var prop = props.Cast<DocumentProperty>().FirstOrDefault(p => p.Name == Prop);
            if (prop == null)
            {
                _cache.Clear();
                return;
            }
            var json = prop.Value as string;
            var list = JsonConvert.DeserializeObject<List<SnipData>>(json) ?? new List<SnipData>();
            _cache = list.ToDictionary(x => x.Id);
            
            // Also load the new snip overlays
            LoadSnipOverlays(wb);
        }

        private const string SnipOverlaysProp = "SnipOverlays";

        public static void SaveSnipOverlays(Excel.Workbook wb)
        {
            if (ThisAddIn.Instance?.Snips?.All == null) return;
            
            var json = JsonConvert.SerializeObject(ThisAddIn.Instance.Snips.All);
            var props = (DocumentProperties)wb.CustomDocumentProperties;
            var existing = props.Cast<DocumentProperty>().FirstOrDefault(p => p.Name == SnipOverlaysProp);
            existing?.Delete();
            props.Add(SnipOverlaysProp, false, MsoDocProperties.msoPropertyTypeString, json);
        }

        public static void LoadSnipOverlays(Excel.Workbook wb)
        {
            if (ThisAddIn.Instance?.Snips == null) return;
            
            var props = (DocumentProperties)wb.CustomDocumentProperties;
            var prop = props.Cast<DocumentProperty>().FirstOrDefault(p => p.Name == SnipOverlaysProp);
            if (prop == null) return;
            
            var json = prop.Value as string;
            if (string.IsNullOrEmpty(json)) return;
            
            try
            {
                var overlays = JsonConvert.DeserializeObject<List<SnipOverlay>>(json);
                if (overlays != null)
                {
                    // Clear existing and add loaded overlays
                    foreach (var overlay in overlays)
                    {
                        ThisAddIn.Instance.Snips.Add(overlay);
                    }
                }
            }
            catch
            {
                // Ignore serialization errors
            }
        }
    }
}
