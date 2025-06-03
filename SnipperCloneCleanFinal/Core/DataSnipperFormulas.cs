using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using SnipperCloneCleanFinal.Infrastructure;

namespace SnipperCloneCleanFinal.Core
{
    public static class DataSnipperFormulas
    {
        private static Dictionary<string, SnipData> _snipDatabase = new Dictionary<string, SnipData>();
        
        public static string CreateTextFormula(string documentPath, int pageNumber, string extractedText, Rectangle bounds)
        {
            var snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Text,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = extractedText,
                Bounds = bounds,
                Created = DateTime.Now
            };
            
            _snipDatabase[snipId] = snipData;
            return $"=DS.TEXTS(\"{snipId}\")";
        }
        
        public static string CreateSumFormula(string documentPath, int pageNumber, double sumValue, Rectangle bounds, List<double> numbers)
        {
            var snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Sum,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = sumValue.ToString(),
                Bounds = bounds,
                Created = DateTime.Now,
                Numbers = numbers
            };

            _snipDatabase[snipId] = snipData;
            return $"=DS.SUMS(\"{snipId}\")";
        }

        public static string CreateTableFormula(string documentPath, int pageNumber, TableData table, Rectangle bounds)
        {
            var snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Table,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                Bounds = bounds,
                Created = DateTime.Now,
                Table = table
            };

            _snipDatabase[snipId] = snipData;
            return $"=DS.TABLE(\"{snipId}\")";
        }
        
        public static string CreateValidationFormula(string documentPath, int pageNumber, Rectangle bounds)
        {
            var snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Validation,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = "✓",
                Bounds = bounds,
                Created = DateTime.Now
            };
            
            _snipDatabase[snipId] = snipData;
            return $"=DS.VALIDATION(\"{snipId}\")";
        }
        
        public static string CreateExceptionFormula(string documentPath, int pageNumber, Rectangle bounds, string reason = "")
        {
            var snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Exception,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = "✗",
                Bounds = bounds,
                Created = DateTime.Now,
                ExceptionReason = reason
            };
            
            _snipDatabase[snipId] = snipData;
            return $"=DS.EXCEPTION(\"{snipId}\")";
        }
        
        public static SnipData GetSnipData(string snipId)
        {
            return _snipDatabase.TryGetValue(snipId, out var data) ? data : null;
        }
        
        public static bool NavigateToSnip(string snipId)
        {
            var snipData = GetSnipData(snipId);
            if (snipData == null) return false;
            
            try
            {
                // Open document viewer and navigate to the snip location
                var viewer = DocumentViewerManager.GetOrCreateViewer();
                viewer.LoadDocument(snipData.DocumentPath);
                viewer.NavigateToPage(snipData.PageNumber);
                viewer.HighlightRegion(snipData.Bounds.ToDrawingRectangle(), GetSnipColor(snipData.Type));
                viewer.Show();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to navigate to snip {snipId}: {ex.Message}", ex);
                return false;
            }
        }
        
        private static System.Drawing.Color GetSnipColor(SnipMode type)
        {
            return type switch
            {
                SnipMode.Text => System.Drawing.Color.Blue,
                SnipMode.Sum => System.Drawing.Color.Purple,
                SnipMode.Table => System.Drawing.Color.Purple,
                SnipMode.Validation => System.Drawing.Color.Green,
                SnipMode.Exception => System.Drawing.Color.Red,
                _ => System.Drawing.Color.Gray
            };
        }
    }
    
    public class SnipData
    {
        public string Id { get; set; }
        public SnipMode Type { get; set; }
        public string DocumentPath { get; set; }
        public int PageNumber { get; set; }
        public string ExtractedValue { get; set; }
        public Rectangle Bounds { get; set; }
        public DateTime Created { get; set; }
        public List<double> Numbers { get; set; } = new List<double>();
        public string ExceptionReason { get; set; }
        public TableData Table { get; set; }
    }
}
