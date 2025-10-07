using System;
using System.Windows.Forms;

namespace Foca.ExportImport
{
    public interface IFocaPlugin
    {
        string Name { get; }
        string Description { get; }
        string Author { get; }
        string Version { get; }
        void Initialize();
    }

    public sealed class FocaExcelExportPlugin : IFocaPlugin
    {
        public string Name => "Export to Excel";
        public string Description => "Export FOCA project metadata to Excel";
        public string Author => "Andrés Nacimiento";
        public string Version => "1.0.0";

        public void Initialize()
        {
            // En runtime con FOCA usar FocaExportImportPluginApi (FOCA_API) para registrar menús.
            Application.ApplicationExit += (s, e) => { };
        }

        public void OnExport()
        {
            using (var form = new FocaExcelExport.Forms.ExportDialog())
            {
                form.ShowDialog();
            }
        }
    }
}

#if FOCA_API
using System;
using System.IO;
using System.Windows.Forms;
using PluginsAPI;
using PluginsAPI.Elements;

namespace Foca
{
    internal static class PluginDiag
    {
        private static readonly string LogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FocaExcelExport.plugin.log");
        public static void Log(string message)
        {
            try { File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff ") + message + Environment.NewLine); } catch { }
        }
    }

    public class Plugin
    {
        private string _name = "Export to Excel";
        private string _description = "Exporta proyectos FOCA a Excel";
        private readonly Export export;

        public Export exportItems { get { return this.export; } }

        public string name
        {
            get { return this._name; }
            set { this._name = value; }
        }

        public string description
        {
            get { return this._description; }
            set { this._description = value; }
        }

        public Plugin()
        {
            try
            {
                PluginDiag.Log("Plugin ctor start");
                this.export = new Export();

                var hostPanel = new Panel { Dock = DockStyle.Fill, Visible = false };
                var pluginPanel = new PluginPanel(hostPanel, false);
                this.export.Add(pluginPanel);
                PluginDiag.Log("PluginPanel added");

                var root = new ToolStripMenuItem(this._name);
                var exportItem = new ToolStripMenuItem("Export to Excel");
                exportItem.Click += (sender, e) =>
                {
                    try
                    {
                        // Crear el formulario de forma segura usando reflexión
                        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                        var formTypeName = "FocaExcelExport.Forms.ExportDialog";
                        var formType = assembly.GetType(formTypeName);
                        
                        if (formType != null)
                        {
                            var dialog = System.Activator.CreateInstance(formType);
                            var form = (System.Windows.Forms.Form)dialog;
                            form.ShowDialog();
                        }
                        else
                        {
                            MessageBox.Show("Could not find export dialog form.", 
                                "Error", 
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Error);
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
                
                root.DropDownItems.Add(exportItem);

                var pluginMenu = new PluginToolStripMenuItem(root);
                this.export.Add(pluginMenu);
                PluginDiag.Log("Menu added");
            }
            catch (Exception ex)
            {
                PluginDiag.Log("Plugin ctor error: " + ex);
                throw;
            }
        }
    }
}
#endif