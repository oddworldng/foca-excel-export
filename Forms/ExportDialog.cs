using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using FocaExcelExport.Classes;
using System.Threading.Tasks;

namespace FocaExcelExport
{
    public partial class ExportDialog : Form
    {
        private readonly string _connectionString;
        
        public ExportDialog()
        {
            InitializeComponent();
            _connectionString = ConnectionResolver.GetFocaConnectionString();
            try
            {
                var stream = typeof(ExportDialog).Assembly.GetManifestResourceStream("FocaExcelExport.img.export.png");
                if (stream != null)
                {
                    var img = System.Drawing.Image.FromStream(stream);
                    this.btnExport.Image = new System.Drawing.Bitmap(img, new System.Drawing.Size(16, 16));
                }
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
                        lblStatus.Text = $"{processed} registros procesados - {progressReport.StatusMessage}";
                    }
                    else
                    {
                        lblStatus.Text = $"{processed} de {total} registros procesados - {progressReport.StatusMessage}";
                    }
                });
                
                await exporter.ExportToExcelAsync(project.Id, fileName, progress);
                
                lblStatus.Text = "¡Exportación completada con éxito!";
                MessageBox.Show("¡Exportación completada con éxito!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error durante la exportación: {ex.Message}", "Error de exportación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = $"Error: {ex.Message}";
            }
            finally
            {
                btnExport.Enabled = true;
                progressBar.Visible = false;
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