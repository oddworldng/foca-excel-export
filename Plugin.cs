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
            System.Windows.Forms.Application.ApplicationExit += (s, e) => { };
        }

        public void OnExport()
        {
            System.Windows.Forms.MessageBox.Show("Export functionality would go here.");
        }
    }
}

#if FOCA_API
namespace Foca
{
    using System;
    using System.IO;
    using System.Windows.Forms;
    using PluginsAPI;
    using PluginsAPI.Elements;

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
                
                // Intentar cargar el icono desde varias ubicaciones posibles
                string[] possiblePaths = {
                    System.IO.Path.Combine(System.IO.Path.GetDirectoryName(typeof(Plugin).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory, "img", "icon.png"),
                    System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", "icon.png"),
                    System.IO.Path.Combine(System.IO.Path.GetDirectoryName(typeof(Plugin).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory, "icon.png"),
                    System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "icon.png")
                };

                foreach (string path in possiblePaths)
                {
                    try
                    {
                        if (System.IO.File.Exists(path))
                        {
                            root.Image = System.Drawing.Image.FromFile(path);
                            PluginDiag.Log($"Icon loaded from: {path}");
                            break; // Si se carga correctamente, salir del bucle
                        }
                    }
                    catch (Exception ex)
                    {
                        PluginDiag.Log($"Failed to load icon from {path}: {ex.Message}");
                        // Probar con la siguiente ruta
                    }
                }
                
                var exportItem = new ToolStripMenuItem("Export to Excel");
                exportItem.Click += (s, e) =>
                {
                    try
                    {
                        var dialog = new FocaExcelExport.ExportDialog();
                        dialog.ShowDialog();
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