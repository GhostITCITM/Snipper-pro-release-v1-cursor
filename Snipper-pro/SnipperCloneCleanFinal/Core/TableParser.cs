using System;
using System.Collections.Generic;

namespace SnipperCloneCleanFinal.Core
{
    public class TableParser
    {
        public TableData ParseTable(string text)
        {
            var tableData = new TableData();
            
            if (string.IsNullOrWhiteSpace(text))
                return tableData;

            // Simple table parsing - split by lines and columns
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var line in lines)
            {
                var columns = line.Split(new[] { '\t', '|' }, StringSplitOptions.None);
                tableData.AddRow(columns);
            }

            return tableData;
        }
    }
} 