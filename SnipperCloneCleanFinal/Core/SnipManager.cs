using System;
using System.Collections.Generic;
using System.Linq;

namespace SnipperCloneCleanFinal.Core
{
    public sealed class SnipManager
    {
        private readonly List<SnipOverlay> _all = new List<SnipOverlay>();
        public IReadOnlyList<SnipOverlay> All => _all;

        public void Add(SnipOverlay s) => _all.Add(s);
        public void Remove(Guid id)     => _all.RemoveAll(x => x.Id == id);

        public SnipOverlay ById(Guid id) => _all.FirstOrDefault(x => x.Id == id);

        public SnipOverlay ByCell(string sheet, string addr) =>
            _all.FirstOrDefault(x => x.SheetName == sheet && x.CellAddr == addr);
    }
} 