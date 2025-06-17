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
        private static Dictionary<string, SnipData> _cache = new();

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
            var list = JsonConvert.DeserializeObject<List<SnipData>>(json) ?? new();
            _cache = list.ToDictionary(x => x.Id);
        }
    }
}
