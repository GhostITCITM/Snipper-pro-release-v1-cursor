using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Extensibility;

namespace SnipperCloneCleanFinal
{
    [ComVisible(true)]
    [Guid("D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111")]
    [ProgId("SnipperClone.AddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class ThisAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private Excel.Application _application;

        public Excel.Application Application => _application;

        // IDTExtensibility2 interface - MANDATORY for Excel COM add-ins
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _application = (Excel.Application)application;
                System.Diagnostics.Debug.WriteLine("Snipper Pro: OnConnection successful via IDTExtensibility2");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro OnConnection Error: {ex.Message}");
                // Don't throw - would cause add-in to fail loading
            }
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Snipper Pro: OnDisconnection");
                _application = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro OnDisconnection Error: {ex.Message}");
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // IRibbonExtensibility interface - for ribbon UI
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                // Try to load from embedded resource first
                using (var stream = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("SnipperCloneCleanFinal.Assets.SnipperRibbon.xml"))
                {
                    if (stream != null)
                    {
                        using (var reader = new StreamReader(stream))
                        {
                            var ribbonXml = reader.ReadToEnd();
                            System.Diagnostics.Debug.WriteLine("Snipper Pro: Ribbon XML loaded from embedded resource");
                            return ribbonXml;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro Ribbon XML Error: {ex.Message}");
            }

            // Fallback to hardcoded ribbon
            return GetFallbackRibbonXml();
        }

        // This method is called by the ribbon XML onLoad="OnRibbonLoad"
        [ComVisible(true)]
        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            try
            {
                _ribbon = ribbonUI;
                System.Diagnostics.Debug.WriteLine("Snipper Pro: Ribbon loaded successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Snipper Pro Ribbon Load Error: {ex.Message}");
                // Don't throw - would prevent ribbon from loading
            }
        }

        private string GetFallbackRibbonXml()
        {
            System.Diagnostics.Debug.WriteLine("Snipper Pro: Using fallback ribbon XML");
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""SnipperProTab"" label=""SNIPPER PRO"">
        <group id=""SnipGroup"" label=""Document Analysis"">
          <button id=""TextSnipButton"" label=""Text Snip"" size=""large"" onAction=""OnTextSnip"" imageMso=""TextBox"" />
          <button id=""SumSnipButton"" label=""Sum Snip"" size=""large"" onAction=""OnSumSnip"" imageMso=""FunctionWizard"" />
          <button id=""TableSnipButton"" label=""Table Snip"" size=""large"" onAction=""OnTableSnip"" imageMso=""Table"" />
        </group>
        <group id=""ValidationGroup"" label=""Validation"">
          <button id=""ValidateButton"" label=""Validate"" size=""large"" onAction=""OnValidationSnip"" imageMso=""AcceptInvitation"" />
          <button id=""ExceptionButton"" label=""Exception"" size=""large"" onAction=""OnExceptionSnip"" imageMso=""CancelRequest"" />
        </group>
        <group id=""ViewerGroup"" label=""Viewer"">
          <button id=""OpenViewerButton"" label=""Open Viewer"" size=""large"" onAction=""OnOpenViewer"" imageMso=""PictureInsertFromFile"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // Button handlers - ALL must be COM-visible for Excel to find them
        [ComVisible(true)]
        public void OnTextSnip(IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Text Snip functionality - Coming Soon!", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComVisible(true)]
        public void OnSumSnip(IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Sum Snip functionality - Coming Soon!", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComVisible(true)]
        public void OnTableSnip(IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Table Snip functionality - Coming Soon!", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComVisible(true)]
        public void OnValidationSnip(IRibbonControl control)
        {
            try
            {
                if (_application?.ActiveCell != null)
                {
                    _application.ActiveCell.Value = "✓";
                    MessageBox.Show("Cell marked as validated!", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComVisible(true)]
        public void OnExceptionSnip(IRibbonControl control)
        {
            try
            {
                if (_application?.ActiveCell != null)
                {
                    _application.ActiveCell.Value = "✗";
                    MessageBox.Show("Cell marked as exception!", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComVisible(true)]
        public void OnOpenViewer(IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Document Viewer - Coming Soon!", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComVisible(true)]
        public void OnMarkupSnip(IRibbonControl control)
        {
            try
            {
                MessageBox.Show("Markup functionality - Coming Soon!", "Snipper Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Snipper Pro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
} 