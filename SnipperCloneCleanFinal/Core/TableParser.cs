using System;
using System.Collections.Generic;
using System.Linq;

namespace SnipperCloneCleanFinal.Core
{
    public class TableParser
    {
        public TableData ParseTable(string text)
        {
            var tableData = new TableData();
            
            if (string.IsNullOrWhiteSpace(text))
                return tableData;

            // Enhanced table parsing - handle tab-separated data from column snip
            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var line in lines)
            {
                // Split by tabs first (primary separator from column snip), then by pipes (fallback)
                string[] columns;
                if (line.Contains('\t'))
                {
                    columns = line.Split('\t');
                }
                else
                {
                    columns = line.Split(new[] { '|' }, StringSplitOptions.None);
                }
                
                // Clean up the columns and add to table data
                var cleanedColumns = columns.Select(col => col?.Trim() ?? string.Empty).ToArray();
                tableData.AddRow(cleanedColumns);
            }

            // Set header detection based on common patterns
            if (tableData.Rows.Count > 0)
            {
                var firstRow = tableData.Rows[0];
                // Simple heuristic: if first row contains common header words and no numbers, treat as header
                bool hasHeaderWords = firstRow.Any(cell => 
                    !string.IsNullOrEmpty(cell) && 
                    (cell.Contains("Name") || cell.Contains("Date") || cell.Contains("Amount") || 
                     cell.Contains("Description") || cell.Contains("Total") || cell.Contains("Item") ||
                     cell.Contains("Type") || cell.Contains("Status") || cell.Contains("Category")));
                     
                bool hasNumbers = firstRow.Any(cell => 
                    !string.IsNullOrEmpty(cell) && 
                    cell.Any(char.IsDigit) && 
                    decimal.TryParse(cell.Replace(",", "").Replace("$", ""), out _));
                
                if (hasHeaderWords && !hasNumbers && tableData.Rows.Count > 1)
                {
                    tableData.SetHeaders(firstRow);
                    tableData.Rows.RemoveAt(0); // Remove header from data rows
                }
            }

            return tableData;
        }
    }
} 