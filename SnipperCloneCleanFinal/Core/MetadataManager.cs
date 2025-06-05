using System;
using System.Collections.Generic;

namespace SnipperCloneCleanFinal.Core
{
    public class MetadataManager
    {
        private readonly Dictionary<string, SnipRecord> _records = new Dictionary<string, SnipRecord>();

        public void SaveRecord(SnipRecord record)
        {
            if (record != null && !string.IsNullOrEmpty(record.CellAddress))
            {
                _records[record.CellAddress] = record;
            }
        }

        public SnipRecord GetRecord(string cellAddress)
        {
            return _records.TryGetValue(cellAddress, out var record) ? record : null;
        }

        public IEnumerable<SnipRecord> GetAllRecords()
        {
            return _records.Values;
        }

        public void ClearRecords()
        {
            _records.Clear();
        }
    }
} 