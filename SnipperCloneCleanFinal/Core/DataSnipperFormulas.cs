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
            return $"=SnipperPro.Connect.TEXTS(\"{snipId}\")";
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
            return $"=SnipperPro.Connect.SUMS(\"{snipId}\")";
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
            return $"=SnipperPro.Connect.TABLE(\"{snipId}\")";
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
            System.Diagnostics.Debug.WriteLine($"Created validation snip: ID={snipId}, Doc={documentPath}, Page={pageNumber}");
            return $"=SnipperPro.Connect.VALIDATION(\"{snipId}\")";
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
            System.Diagnostics.Debug.WriteLine($"Created exception snip: ID={snipId}, Doc={documentPath}, Page={pageNumber}");
            return $"=SnipperPro.Connect.EXCEPTION(\"{snipId}\")";
        }
        
        public static SnipData GetSnipData(string snipId)
        {
            return _snipDatabase.TryGetValue(snipId, out var data) ? data : null;
        }
        
        public static bool NavigateToSnip(string snipId)
        {
            System.Diagnostics.Debug.WriteLine($"NavigateToSnip called with ID: {snipId}");
            
            var snipData = GetSnipData(snipId);
            if (snipData == null) 
            {
                System.Diagnostics.Debug.WriteLine($"No snip data found for ID: {snipId}");
                return false;
            }
            
            System.Diagnostics.Debug.WriteLine($"Found snip data: Type={snipData.Type}, Document={snipData.DocumentPath}, Page={snipData.PageNumber}");
            
            try
            {
                // Get the main document viewer from the add-in
                System.Diagnostics.Debug.WriteLine("Getting add-in instance...");
                var addIn = SnipperCloneCleanFinal.ThisAddIn.Instance;
                
                if (addIn?.DocumentViewer != null && !addIn.DocumentViewer.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Add-in and document viewer found");
                    
                    // Use the main document viewer for navigation
                    var viewer = addIn.DocumentViewer;
                    
                    // Load the document if not already loaded
                    System.Diagnostics.Debug.WriteLine($"Loading document: {snipData.DocumentPath}");
                    if (!viewer.LoadDocument(snipData.DocumentPath))
                    {
                        Logger.Error($"Failed to load document for snip {snipId}");
                        System.Diagnostics.Debug.WriteLine("Document loading failed");
                        return false;
                    }
                    
                    // Navigate to the page and highlight the region
                    System.Diagnostics.Debug.WriteLine($"Navigating to page {snipData.PageNumber}");
                    viewer.NavigateToPage(snipData.PageNumber);
                    
                    System.Diagnostics.Debug.WriteLine($"Highlighting region: {snipData.Bounds}");
                    viewer.HighlightRegion(snipData.Bounds.ToDrawingRectangle(), GetSnipColor(snipData.Type));
                    
                    // Show and bring to front
                    System.Diagnostics.Debug.WriteLine("Showing and bringing viewer to front");
                    viewer.Show();
                    viewer.BringToFront();
                    
                    Logger.Info($"Successfully navigated to snip {snipId} on page {snipData.PageNumber}");
                    System.Diagnostics.Debug.WriteLine("Navigation completed successfully");
                    return true;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Document viewer not available - addIn: {addIn}, viewer: {addIn?.DocumentViewer}, disposed: {addIn?.DocumentViewer?.IsDisposed}");
                    Logger.Error("Document viewer not available for navigation");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to navigate to snip {snipId}: {ex.Message}", ex);
                System.Diagnostics.Debug.WriteLine($"Exception in NavigateToSnip: {ex.Message}");
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
