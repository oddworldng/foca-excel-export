using System;
using System.Windows.Forms;
using FocaExcelExport.Forms;

namespace FocaExcelExport
{
    // This is a placeholder interface since we don't have access to the actual FOCA plugin interface
    // In a real implementation, this would be the actual interface from FOCA
    public interface IFocaPlugin
    {
        string Name { get; }
        string Description { get; }
        void Initialize();
        ToolStripMenuItem GetMenuItem();
    }

    public class FocaExcelExportPlugin : IFocaPlugin
    {
        public string Name => "Export to Excel";
        public string Description => "Export FOCA project metadata to Excel format";

        public void Initialize()
        {
            // Initialize the plugin - typically this is called when FOCA loads the plugin
            // No specific initialization required for this plugin
        }

        public ToolStripMenuItem GetMenuItem()
        {
            var menuItem = new ToolStripMenuItem(Name);
            menuItem.Click += (sender, e) =>
            {
                try
                {
                    // Show the export dialog
                    using (var exportDialog = new ExportDialog())
                    {
                        exportDialog.ShowDialog();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error launching export dialog: {ex.Message}", 
                        "Export Error", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Error);
                }
            };

            return menuItem;
        }
    }
}