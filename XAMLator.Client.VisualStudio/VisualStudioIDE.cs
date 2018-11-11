using EnvDTE;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XAMLator.Client.VisualStudio
{
    public class VisualStudioIDE : IIDE
    {
        readonly DTE _dte;

        public event EventHandler<DocumentChangedEventArgs> DocumentChanged;

        public VisualStudioIDE(DTE dte)
        {
            _dte = dte;
        }

        public void MonitorEditorChanges()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (_dte == null)
                throw new InvalidOperationException("XAMLanator not initialized.");

            var textEditorEvents = _dte.Events.TextEditorEvents;
            textEditorEvents.LineChanged += OnLineChanged;
        }

        private void OnLineChanged(TextPoint startPoint, TextPoint endPoint, int hint)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var textDocument = startPoint?.Parent;
            if (textDocument == null) return;

            var document = startPoint.Parent.Parent;
            if (document == null) return;

            EmitDocumentChanged(document);
        }

        public Task RunTarget(string targetName)
        {
            // TODO:
            throw new NotImplementedException();
        }

        public void ShowError(string error, Exception ex = null)
        {
            Log.Error(error);
            MessageBox.Show(error);
        }

        void EmitDocumentChanged(Document document)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            TextDocument doc = (TextDocument)(_dte.ActiveDocument.Object("TextDocument"));
            var point = doc.StartPoint.CreateEditPoint();
            string docText = point.GetText(doc.EndPoint);

            Log.Information($"XAML document changed {document.Name}");
            DocumentChanged?.Invoke(this, new DocumentChangedEventArgs(document.FullName, docText, null, null));
        }
    }
}
