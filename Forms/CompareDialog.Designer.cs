using System.Windows.Forms;

namespace FocaExcelExport
{
    partial class CompareDialog
    {
        private System.ComponentModel.IContainer components = null;
        private Label lblTitle;
        private Label lblBase;
        private Label lblNew;
        private Label lblOut;
        private TextBox txtBase;
        private TextBox txtNew;
        private TextBox txtOut;
        private Button btnBrowseBase;
        private Button btnBrowseNew;
        private Button btnBrowseOut;
        private Button btnCompare;
        private Button btnClose;
        private Button btnOpen;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblSuccess;
        

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblTitle = new Label();
            this.lblBase = new Label();
            this.lblNew = new Label();
            this.lblOut = new Label();
            this.txtBase = new TextBox();
            this.txtNew = new TextBox();
            this.txtOut = new TextBox();
            this.btnBrowseBase = new Button();
            this.btnBrowseNew = new Button();
            this.btnBrowseOut = new Button();
            this.btnCompare = new Button();
            this.btnClose = new Button();
            this.btnOpen = new Button();
            this.progressBar = new ProgressBar();
            this.lblStatus = new Label();
            this.lblSuccess = new Label();
            this.SuspendLayout();

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Text = "Comparar informes (base vs nuevo)";
            this.lblTitle.Left = 12;
            this.lblTitle.Top = 9;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold);

            // lblBase
            this.lblBase.AutoSize = true;
            this.lblBase.Text = "Informe base (anterior):";
            this.lblBase.Left = 12;
            this.lblBase.Top = 36;

            // txtBase
            this.txtBase.Left = 12;
            this.txtBase.Top = 54;
            this.txtBase.Width = 420;

            // btnBrowseBase
            this.btnBrowseBase.Text = "...";
            this.btnBrowseBase.Left = 438;
            this.btnBrowseBase.Top = 52;
            this.btnBrowseBase.Width = 32;

            // lblNew
            this.lblNew.AutoSize = true;
            this.lblNew.Text = "Informe nuevo (actual):";
            this.lblNew.Left = 12;
            this.lblNew.Top = 86;

            // txtNew
            this.txtNew.Left = 12;
            this.txtNew.Top = 104;
            this.txtNew.Width = 420;

            // btnBrowseNew
            this.btnBrowseNew.Text = "...";
            this.btnBrowseNew.Left = 438;
            this.btnBrowseNew.Top = 102;
            this.btnBrowseNew.Width = 32;

            // lblOut
            this.lblOut.AutoSize = true;
            this.lblOut.Text = "Guardar informe comparativo:";
            this.lblOut.Left = 12;
            this.lblOut.Top = 136;

            // txtOut
            this.txtOut.Left = 12;
            this.txtOut.Top = 154;
            this.txtOut.Width = 420;

            // btnBrowseOut
            this.btnBrowseOut.Text = "...";
            this.btnBrowseOut.Left = 438;
            this.btnBrowseOut.Top = 152;
            this.btnBrowseOut.Width = 32;

            // btnCompare
            this.btnCompare.Text = "Comparar";
            this.btnCompare.Left = 12;
            this.btnCompare.Top = 194;
            this.btnCompare.AutoSize = true;
            this.btnCompare.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.btnCompare.MinimumSize = new System.Drawing.Size(110, 32);
            this.btnCompare.Width = 120;
            this.btnCompare.Padding = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.btnCompare.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCompare.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;

            // btnClose
            this.btnClose.Text = "Cerrar";
            this.btnClose.Left = 144;
            this.btnClose.Top = 194;
            this.btnClose.AutoSize = true;
            this.btnClose.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.btnClose.MinimumSize = new System.Drawing.Size(90, 32);
            this.btnClose.Width = 120;
            this.btnClose.Padding = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;

            // btnOpen
            this.btnOpen.Text = "Abrir Excel";
            this.btnOpen.Left = 12;
            this.btnOpen.Top = 194;
            this.btnOpen.AutoSize = true;
            this.btnOpen.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.btnOpen.MinimumSize = new System.Drawing.Size(100, 32);
            this.btnOpen.Width = 120;
            this.btnOpen.Padding = new System.Windows.Forms.Padding(8, 0, 8, 0);
            this.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOpen.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnOpen.Visible = false;

            

            // progressBar
            this.progressBar.Left = 12;
            this.progressBar.Top = 230;
            this.progressBar.Width = 458;

            // lblStatus
            this.lblStatus.AutoSize = true;
            this.lblStatus.Left = 12;
            this.lblStatus.Top = 258;
            this.lblStatus.Text = "Listo";

            // lblSuccess
            this.lblSuccess.AutoSize = true;
            this.lblSuccess.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblSuccess.ForeColor = System.Drawing.Color.FromArgb(34, 139, 34);
            this.lblSuccess.Left = 12;
            this.lblSuccess.Top = 230;
            this.lblSuccess.Text = "Comparación finalizada con éxito";
            this.lblSuccess.Visible = false;

            // Form
            this.ClientSize = new System.Drawing.Size(482, 284);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.lblBase);
            this.Controls.Add(this.txtBase);
            this.Controls.Add(this.btnBrowseBase);
            this.Controls.Add(this.lblNew);
            this.Controls.Add(this.txtNew);
            this.Controls.Add(this.btnBrowseNew);
            this.Controls.Add(this.lblOut);
            this.Controls.Add(this.txtOut);
            this.Controls.Add(this.btnBrowseOut);
            
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.btnCompare);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblSuccess);
            this.Controls.Add(this.lblStatus);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "Comparar informes";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}


