using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using FocaExcelExport.Classes;
using System.Threading.Tasks;
using System.Drawing;

namespace FocaExcelExport
{
    public partial class ExportDialog : Form
    {
        private readonly string _connectionString;
        private Image _exportIconBase;
        private Image _closeIconBase;
        private Image _openIconBase;
        private string _lastExportPath;
        
        public ExportDialog()
        {
            InitializeComponent();
            _connectionString = ConnectionResolver.GetFocaConnectionString();
            try
            {
                var stream = typeof(ExportDialog).Assembly.GetManifestResourceStream("FocaExcelExport.img.export.png");
                if (stream != null)
                {
                    _exportIconBase = Image.FromStream(stream);
                    ApplyExportIconSize();
                    // Cuando cambie el tamaño (AutoSize) reescalamos el icono y reposicionamos
                    this.btnExport.SizeChanged += (s, e) => { ApplyExportIconSize(); PositionActionButtons(); };
                }
                // Iconos
                LoadCloseIcon();
                LoadOpenIcon();
            }
            catch { }

            // Posicionar botones inicialmente
            PositionActionButtons();
        }

        private void ApplyExportIconSize()
        {
            if (_exportIconBase == null) return;
            // objetivo: coherente con botones FOCA. Usar 24px o ajustar al alto del botón
            int target = Math.Min(24, Math.Max(16, this.btnExport.Height - 8));
            try
            {
                var bmp = new Bitmap(_exportIconBase, new Size(target, target));
                // liberar imagen previa si existiese
                var old = this.btnExport.Image;
                this.btnExport.Image = bmp;
                if (old != null && !ReferenceEquals(old, _exportIconBase)) old.Dispose();
            }
            catch { }
        }

        private void LoadCloseIcon()
        {
            try
            {
                string asmDir = System.IO.Path.GetDirectoryName(typeof(ExportDialog).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory;
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] candidates = new[]
                {
                    System.IO.Path.Combine(asmDir, "img", "exit.png"),
                    System.IO.Path.Combine(baseDir, "img", "exit.png"),
                    System.IO.Path.Combine(asmDir, "exit.png"),
                    System.IO.Path.Combine(baseDir, "exit.png")
                };
                foreach (var p in candidates)
                {
                    if (System.IO.File.Exists(p))
                    {
                        using (var fs = System.IO.File.OpenRead(p))
                        {
                            _closeIconBase = Image.FromStream(fs);
                        }
                        break;
                    }
                }
                if (_closeIconBase == null)
                {
                    using (var s = typeof(ExportDialog).Assembly.GetManifestResourceStream("FocaExcelExport.img.exit.png"))
                    {
                        if (s != null) _closeIconBase = Image.FromStream(s);
                    }
                }
                if (_closeIconBase != null)
                {
                    ApplyCloseIconSize();
                    this.btnClose.SizeChanged += (s, e) => ApplyCloseIconSize();
                }
            }
            catch { }
        }

        private void LoadOpenIcon()
        {
            try
            {
                string asmDir = System.IO.Path.GetDirectoryName(typeof(ExportDialog).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory;
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] candidates = new[]
                {
                    System.IO.Path.Combine(asmDir, "img", "open.png"),
                    System.IO.Path.Combine(baseDir, "img", "open.png"),
                    System.IO.Path.Combine(asmDir, "open.png"),
                    System.IO.Path.Combine(baseDir, "open.png")
                };
                foreach (var p in candidates)
                {
                    if (System.IO.File.Exists(p))
                    {
                        using (var fs = System.IO.File.OpenRead(p))
                        {
                            _openIconBase = Image.FromStream(fs);
                        }
                        break;
                    }
                }
                if (_openIconBase == null)
                {
                    using (var s = typeof(ExportDialog).Assembly.GetManifestResourceStream("FocaExcelExport.img.open.png"))
                    {
                        if (s != null) _openIconBase = Image.FromStream(s);
                    }
                }
                if (_openIconBase != null)
                {
                    ApplyOpenIconSize();
                    this.btnOpen.SizeChanged += (s, e) => ApplyOpenIconSize();
                }
            }
            catch { }
        }

        private void ApplyOpenIconSize()
        {
            if (_openIconBase == null) return;
            int target = Math.Min(20, Math.Max(16, this.btnOpen.Height - 8));
            try
            {
                var bmp = new Bitmap(_openIconBase, new Size(target, target));
                var old = this.btnOpen.Image;
                this.btnOpen.Image = bmp;
                if (old != null && !ReferenceEquals(old, _openIconBase)) old.Dispose();
            }
            catch { }
        }
        private void ApplyCloseIconSize()
        {
            if (_closeIconBase == null) return;
            int target = Math.Min(20, Math.Max(16, this.btnClose.Height - 8));
            try
            {
                var bmp = new Bitmap(_closeIconBase, new Size(target, target));
                var old = this.btnClose.Image;
                this.btnClose.Image = bmp;
                if (old != null && !ReferenceEquals(old, _closeIconBase)) old.Dispose();
            }
            catch { }
        }

        private void PositionActionButtons()
        {
            try
            {
                int spacing = 10; // separación coherente
                // Si Abrir Excel está visible, va en la posición del botón Exportar
                if (this.btnOpen.Visible)
                {
                    this.btnOpen.Top = this.btnExport.Top;
                    this.btnOpen.Left = this.btnExport.Left;
                }
                // El botón de referencia para colocar Cerrar
                var baseBtn = this.btnOpen.Visible ? this.btnOpen : this.btnExport;
                this.btnClose.Top = baseBtn.Top;
                this.btnClose.Left = baseBtn.Left + baseBtn.Width + spacing;
            }
            catch { }
        }

        private async void ExportDialog_Load(object sender, EventArgs e)
        {
            await LoadProjectsAsync();
        }

        private async Task LoadProjectsAsync()
        {
            try
            {
                lblStatus.Text = "Cargando proyectos...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
                
                var schemaResolver = new SchemaResolver(_connectionString);
                var projectsTable = await schemaResolver.FindProjectsTableAsync();
                
                // If we couldn't find a projects table, show error and disable export
                if (string.IsNullOrEmpty(projectsTable))
                {
                    MessageBox.Show("No se pudo encontrar la tabla de proyectos en la base de datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnExport.Enabled = false;
                    lblStatus.Text = "Error: no se encontró la tabla de proyectos.";
                    return;
                }
                
                var projectIdColumn = await schemaResolver.FindProjectIdColumnAsync(projectsTable);
                var projectNameColumn = "ProjectName"; // Based on FOCA Project entity structure
                
                // If no name column found, try common name columns
                if (string.IsNullOrEmpty(projectNameColumn) || projectNameColumn == projectIdColumn)
                {
                    var columns = await schemaResolver.GetColumnsAsync(projectsTable);
                    foreach (var col in columns)
                    {
                        if (col.ToLower().Contains("name") || col.ToLower().Contains("title") || col.ToLower().Contains("project"))
                        {
                            if (col.ToLower() != "id" && !col.ToLower().Contains("id"))
                            {
                                projectNameColumn = col;
                                break;
                            }
                        }
                    }
                }

                // Load projects from database
                using (var connection = new SqlConnection(_connectionString))
                {
                    await connection.OpenAsync();
                    
                    string query = $"SELECT [{projectIdColumn}], [{projectNameColumn}] FROM [dbo].[{projectsTable}] ORDER BY [{projectNameColumn}]";
                    using (var command = new SqlCommand(query, connection))
                    {
                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            cmbProjects.Items.Clear();
                            
                            while (await reader.ReadAsync())
                            {
                                var projectId = reader[0];
                                var projectName = reader[1]?.ToString() ?? "Unnamed Project";
                                
                                // Add project to combo box with both ID and name
                                cmbProjects.Items.Add(new ProjectInfo 
                                { 
                                    Id = Convert.ToInt32(projectId), 
                                    Name = projectName 
                                });
                            }
                        }
                    }
                }

                if (cmbProjects.Items.Count > 0)
                {
                    cmbProjects.SelectedIndex = 0;
                    lblStatus.Text = $"Cargados {cmbProjects.Items.Count} proyectos.";
                }
                else
                {
                    lblStatus.Text = "No se encontraron proyectos en la base de datos.";
                    btnExport.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading projects: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = $"Error: {ex.Message}";
                btnExport.Enabled = false;
            }
            finally
            {
                progressBar.Visible = false;
                progressBar.Style = ProgressBarStyle.Continuous;
            }
        }

        private async void btnExport_Click(object sender, EventArgs e)
        {
            if (cmbProjects.SelectedItem == null)
            {
                MessageBox.Show("Selecciona un proyecto para exportar.", "Proyecto no seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedProject = (ProjectInfo)cmbProjects.SelectedItem;
            
            // Show save file dialog
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
                saveDialog.Title = "Guardar datos exportados";
                saveDialog.FileName = $"foca_export_{selectedProject.Name}_{DateTime.Now:yyyyMMdd}.xlsx";
                
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        await ExportProjectAsync(selectedProject, saveDialog.FileName);
                    }
                    catch (Exception ex)
                    {
                        // Mostrar información detallada de diagnóstico
                        var errorMessage = $"Error durante la exportación: {ex.Message}\n\n";
                        errorMessage += $"Excepción interna: {ex.InnerException?.Message}\n\n";
                        errorMessage += $"Traza: {ex.StackTrace}";
                        
                        MessageBox.Show(errorMessage, 
                            "Error de exportación", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Error);
                    }
                }
            }
        }

        private async Task ExportProjectAsync(ProjectInfo project, string fileName)
        {
            try
            {
                btnExport.Enabled = false;
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
                lblStatus.Text = "Iniciando exportación...";
                // Estado inicial botones
                btnOpen.Visible = false;
                btnExport.Visible = true;
                PositionActionButtons();
                
                var exporter = new Exporter(_connectionString);
                
                // Set up progress reporting
                var progress = new Progress<ExportProgress>(progressReport =>
                {
                    var value = progressReport.PercentComplete;
                    if (value < progressBar.Minimum) value = progressBar.Minimum;
                    if (value > progressBar.Maximum) value = progressBar.Maximum;
                    progressBar.Value = value;
                    var processed = progressReport.CurrentRecord;
                    var total = progressReport.TotalRecords;
                    if (total > 0 && processed > total)
                    {
                        lblStatus.Text = $"{processed} registros procesados";
                    }
                    else
                    {
                        lblStatus.Text = $"{processed} de {total} registros procesados";
                    }
                });
                
                await exporter.ExportToExcelAsync(project.Id, fileName, progress);
                
                // Mensaje de éxito destacado
                lblStatus.Text = string.Empty;
                btnExport.Enabled = false;
                btnExport.Visible = false; // ocultar Exportar
                btnOpen.Visible = true; // mostrar Abrir Excel en el lugar de Exportar
                btnClose.Visible = true;
                btnClose.Click -= BtnClose_Click; // evitar doble suscripción
                btnClose.Click += BtnClose_Click;
                lblSuccess.Text = "Exportación finalizada con éxito";
                lblSuccess.Visible = true;
                progressBar.Visible = false;
                // Guardar ruta y mostrar botón Abrir
                _lastExportPath = fileName;
                btnOpen.Click -= BtnOpen_Click;
                btnOpen.Click += BtnOpen_Click;
                PositionActionButtons();
                MessageBox.Show("¡Exportación completada con éxito!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error durante la exportación: {ex.Message}", "Error de exportación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = $"Error: {ex.Message}";
            }
            finally
            {
                if (!btnClose.Visible) btnExport.Enabled = true;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            try { this.Close(); } catch { }
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_lastExportPath) && System.IO.File.Exists(_lastExportPath))
                {
                    System.Diagnostics.Process.Start(_lastExportPath);
                }
                else
                {
                    MessageBox.Show("No se encontró el archivo exportado.", "Abrir Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se pudo abrir el archivo: {ex.Message}", "Abrir Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // Helper class to hold project information
    public class ProjectInfo
    {
        public int Id { get; set; }
        public string Name { get; set; }
        
        public override string ToString()
        {
            return Name;
        }
    }
}