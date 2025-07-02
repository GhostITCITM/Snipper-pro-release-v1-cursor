using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace SnipperCloneCleanFinal.Core
{
    public enum SnipKind { Text, Sum, Table, Validation, Exception, Image }

    public sealed class SnipOverlay
    {
        public Guid   Id        { get; private set; } = Guid.NewGuid();
        public string DocPath   { get; set; }
        public int    PageIndex { get; set; }
        public System.Drawing.Rectangle Bounds { get; set; }
        public SnipKind Kind    { get; set; }
        public string Extracted { get; set; }

        // link back into Excel
        public string SheetName { get; set; }
        public string CellAddr  { get; set; }

        // tiny red delete square
        public System.Drawing.Rectangle DeleteBox => new System.Drawing.Rectangle(
            Bounds.Right - 12, Bounds.Top, 12, 12);

        public bool HitDelete(Point p) => DeleteBox.Contains(p);
    }
} 