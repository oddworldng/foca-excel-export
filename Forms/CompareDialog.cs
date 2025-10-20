using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace FocaExcelExport
{
    public partial class CompareDialog : Form
    {
        private string _baseFilePath;
        private string _newFilePath;
        private string _outputPath;
        private string _lastOutputPath;
        private Image _closeIconBase;
        private Image _openIconBase;
        private Image _compareIconBase;

        public CompareDialog()
        {
            InitializeComponent();
            btnCompare.Click += BtnCompare_Click;
            btnClose.Click += (s, e) => Close();
            btnBrowseBase.Click += (s, e) => BrowseExcel(txtBase, out _baseFilePath);
            btnBrowseNew.Click += (s, e) => BrowseExcel(txtNew, out _newFilePath);
            btnBrowseOut.Click += (s, e) => SaveExcel(txtOut, out _outputPath);

            // Reutilizar iconos de ExportDialog: Cerrar y Abrir
            LoadCloseIcon();
            LoadOpenIcon();
            LoadCompareIcon();
            if (_closeIconBase != null)
            {
                ApplyCloseIconSize();
                this.btnClose.SizeChanged += (s, e) => ApplyCloseIconSize();
            }
            if (_openIconBase != null)
            {
                ApplyOpenIconSize();
                this.btnOpen.SizeChanged += (s, e) => ApplyOpenIconSize();
            }
            if (_compareIconBase != null)
            {
                ApplyCompareIconSize();
                this.btnCompare.SizeChanged += (s, e) => ApplyCompareIconSize();
            }
            btnOpen.Click += BtnOpen_Click;

            // Alinear botones y controles con la misma lógica que ExportDialog
            this.Load += (s, e) => AdjustDialogLayout();
            this.SizeChanged += (s, e) => AdjustDialogLayout();
        }

        private void BrowseExcel(TextBox target, out string field)
        {
            field = null;
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
                ofd.Title = "Seleccionar Excel";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    field = ofd.FileName;
                    target.Text = field;
                }
            }
        }

        private void SaveExcel(TextBox target, out string field)
        {
            field = null;
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
                sfd.Title = "Guardar informe comparativo";
                sfd.FileName = DateTime.Now.ToString("yyyyMMdd") + "_Informe_Comparativo_FOCA.xlsx";
                sfd.OverwritePrompt = false; // no preguntar, se sobrescribe si existe
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    field = sfd.FileName;
                    target.Text = field;
                }
            }
        }

        private async void BtnCompare_Click(object sender, EventArgs e)
        {
            _baseFilePath = txtBase.Text?.Trim();
            _newFilePath = txtNew.Text?.Trim();
            _outputPath = txtOut.Text?.Trim();

            if (string.IsNullOrWhiteSpace(_baseFilePath) || string.IsNullOrWhiteSpace(_newFilePath))
            {
                MessageBox.Show("Selecciona los dos ficheros de Excel a comparar.", "Ficheros requeridos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(_outputPath))
            {
                MessageBox.Show("Selecciona dónde guardar el informe comparativo.", "Salida requerida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ToggleUi(false);
                lblSuccess.Visible = false;
                lblStatus.ForeColor = SystemColors.ControlText;
                lblStatus.Text = "Comparando...";
                progressBar.Style = ProgressBarStyle.Marquee;
                progressBar.Visible = true;
                PositionActionButtons();

                var comparer = new Classes.ExcelComparer();
                await comparer.CompareAsync(_baseFilePath, _newFilePath, _outputPath, p =>
                {
                    try
                    {
                        if (InvokeRequired)
                        {
                            BeginInvoke(new Action(() => UpdateProgress(p)));
                        }
                        else
                        {
                            UpdateProgress(p);
                        }
                    }
                    catch { }
                }, "URL", true);

                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = progressBar.Maximum;
                progressBar.Visible = false;
                lblStatus.Text = string.Empty;
                lblSuccess.Text = "Comparación finalizada con éxito";
                lblSuccess.Visible = true;
                // Mostrar botón Abrir y ocultar Comparar
                _lastOutputPath = _outputPath;
                btnOpen.Visible = true;
                btnOpen.Enabled = true;
                btnCompare.Visible = false;
                btnClose.Visible = true;
                PositionActionButtons();
                MessageBox.Show("Informe comparativo generado con éxito.", "Comparar", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Mensaje visible en UI en rojo (danger)
                lblSuccess.Visible = false;
                lblStatus.ForeColor = System.Drawing.Color.FromArgb(192, 0, 0);
                lblStatus.Text = "Error: no se pudo generar el informe. Revisa las rutas de los ficheros y permisos.";
                MessageBox.Show($"Error comparando Excels: {ex.Message}", "Comparar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Permitir volver a intentar
                btnCompare.Enabled = true;
                btnCompare.Visible = true;
                PositionActionButtons();
            }
            finally
            {
                // Rehabilitar navegación
                txtBase.Enabled = true;
                txtNew.Enabled = true;
                txtOut.Enabled = true;
                btnBrowseBase.Enabled = true;
                btnBrowseNew.Enabled = true;
                btnBrowseOut.Enabled = true;
                btnClose.Enabled = true;
                AdjustDialogLayout();
            }
        }

        private void UpdateProgress(Classes.CompareProgress p)
        {
            if (p.TotalSteps > 0)
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                var val = Math.Min(progressBar.Maximum, Math.Max(progressBar.Minimum, (int)Math.Round(100.0 * p.CurrentStep / p.TotalSteps)));
                progressBar.Value = val;
                lblStatus.ForeColor = SystemColors.ControlText;
                lblStatus.Text = p.Message;
            }
            else
            {
                progressBar.Style = ProgressBarStyle.Marquee;
                lblStatus.ForeColor = SystemColors.ControlText;
                lblStatus.Text = p.Message;
            }
        }

        private void ToggleUi(bool enabled)
        {
            txtBase.Enabled = enabled;
            txtNew.Enabled = enabled;
            txtOut.Enabled = enabled;
            btnBrowseBase.Enabled = enabled;
            btnBrowseNew.Enabled = enabled;
            btnBrowseOut.Enabled = enabled;
            btnCompare.Enabled = enabled;
            btnClose.Enabled = enabled;
            if (!enabled) btnOpen.Enabled = false;
        }

        // Misma lógica de diseño que ExportDialog: colocar botones y barra/labels
        private void PositionActionButtons()
        {
            try
            {
                int spacing = 10;
                // Base para el layout: debajo del último input (txtOut)
                int topAfterInputs = txtOut.Top + txtOut.Height + spacing;

                // Colocar botones principales
                if (this.btnOpen.Visible)
                {
                    this.btnOpen.Top = topAfterInputs;
                    this.btnOpen.Left = this.btnCompare.Left; // misma columna base
                }
                this.btnCompare.Top = topAfterInputs;

                // Cerrar a la derecha del botón visible de referencia
                var baseBtn = this.btnOpen.Visible ? this.btnOpen : this.btnCompare;
                this.btnClose.Top = topAfterInputs;
                this.btnClose.Left = baseBtn.Left + baseBtn.Width + spacing;

                // Barra de progreso y labels debajo de los botones
                int topAfterButtons = Math.Max(btnCompare.Top + btnCompare.Height, btnClose.Top + btnClose.Height) + spacing;
                if (btnOpen.Visible)
                    topAfterButtons = Math.Max(topAfterButtons, btnOpen.Top + btnOpen.Height + spacing);

                progressBar.Top = topAfterButtons;
                lblSuccess.Top = topAfterButtons;
                lblStatus.Top = topAfterButtons + progressBar.Height + 8;
                lblStatus.Visible = true;
            }
            catch { }
        }

        private void AdjustDialogHeight()
        {
            try
            {
                int bottomVisible = 0;
                if (lblStatus.Visible) bottomVisible = Math.Max(bottomVisible, lblStatus.Bottom);
                if (lblSuccess.Visible) bottomVisible = Math.Max(bottomVisible, lblSuccess.Bottom);
                if (progressBar.Visible) bottomVisible = Math.Max(bottomVisible, progressBar.Bottom);
                if (bottomVisible == 0)
                {
                    bottomVisible = btnCompare.Bottom;
                }
                int desired = bottomVisible + 12;
                int minHeight = 260;
                if (this.ClientSize.Height != Math.Max(desired, minHeight))
                {
                    this.ClientSize = new System.Drawing.Size(this.ClientSize.Width, Math.Max(desired, minHeight));
                }
            }
            catch { }
        }

        private void AdjustDialogLayout()
        {
            PositionActionButtons();
            AdjustDialogHeight();
        }

        private void LoadCloseIcon()
        {
            try
            {
                string asmDir = System.IO.Path.GetDirectoryName(typeof(CompareDialog).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory;
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
                    using (var s = typeof(CompareDialog).Assembly.GetManifestResourceStream("FocaExcelExport.img.exit.png"))
                    {
                        if (s != null) _closeIconBase = Image.FromStream(s);
                    }
                }
            }
            catch { }
        }

        private void LoadOpenIcon()
        {
            try
            {
                string asmDir = System.IO.Path.GetDirectoryName(typeof(CompareDialog).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory;
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
                    using (var s = typeof(CompareDialog).Assembly.GetManifestResourceStream("FocaExcelExport.img.open.png"))
                    {
                        if (s != null) _openIconBase = Image.FromStream(s);
                    }
                }
            }
            catch { }
        }

        private void LoadCompareIcon()
        {
            try
            {
                string asmDir = System.IO.Path.GetDirectoryName(typeof(CompareDialog).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory;
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] candidates = new[]
                {
                    System.IO.Path.Combine(asmDir, "img", "comparar.png"),
                    System.IO.Path.Combine(baseDir, "img", "comparar.png"),
                    System.IO.Path.Combine(asmDir, "comparar.png"),
                    System.IO.Path.Combine(baseDir, "comparar.png")
                };
                foreach (var p in candidates)
                {
                    if (System.IO.File.Exists(p))
                    {
                        using (var fs = System.IO.File.OpenRead(p))
                        {
                            _compareIconBase = Image.FromStream(fs);
                        }
                        break;
                    }
                }
                if (_compareIconBase == null)
                {
                    using (var s = typeof(CompareDialog).Assembly.GetManifestResourceStream("FocaExcelExport.img.comparar.png"))
                    {
                        if (s != null) _compareIconBase = Image.FromStream(s);
                    }
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

        private void ApplyCompareIconSize()
        {
            if (_compareIconBase == null) return;
            int target = Math.Min(24, Math.Max(16, this.btnCompare.Height - 8));
            try
            {
                var bmp = new Bitmap(_compareIconBase, new Size(target, target));
                var old = this.btnCompare.Image;
                this.btnCompare.Image = bmp;
                if (old != null && !ReferenceEquals(old, _compareIconBase)) old.Dispose();
            }
            catch { }
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_lastOutputPath) && System.IO.File.Exists(_lastOutputPath))
                {
                    System.Diagnostics.Process.Start(_lastOutputPath);
                }
                else
                {
                    MessageBox.Show("No se encontró el archivo generado.", "Abrir Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se pudo abrir el archivo: {ex.Message}", "Abrir Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}


