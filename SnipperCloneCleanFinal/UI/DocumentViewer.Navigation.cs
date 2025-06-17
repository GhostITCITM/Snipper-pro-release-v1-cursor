using System.Drawing;
using System.IO;
using System.Linq;

namespace SnipperCloneCleanFinal.UI
{
    public partial class DocumentViewer
    {
        public void NavigateTo(string path, int pageNumber, RectangleF bounds)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return;

            var doc = _loadedDocuments.FirstOrDefault(d => d.FilePath == path);
            if (doc == null)
            {
                if (!LoadDocument(path))
                    return;
                doc = _loadedDocuments.FirstOrDefault(d => d.FilePath == path);
            }
            else if (_currentDocument != doc)
            {
                _currentDocument = doc;
                _documentSelector.SelectedIndex = _loadedDocuments.IndexOf(doc);
            }

            NavigateToPage(pageNumber);
            var rect = new Rectangle((int)bounds.X, (int)bounds.Y, (int)bounds.Width, (int)bounds.Height);
            HighlightRegion(rect, Color.Yellow);
        }
    }
}
