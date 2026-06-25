namespace WinForms
{
    partial class FormVsechnyKabely
    {
        /// <summary>
        /// Vyžadovaná proměnná návrháře.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Uvolněte všechny používané prostředky.
        /// </summary>
        /// <param name="disposing">true pokud by měl být spravovaný prostředek odstraněn; jinak false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Kód generovaný Návrhářem Windows Form

        /// <summary>
        /// Metoda vyžadovaná pro podporu Návrháře - neupravovat
        /// obsah této metody v editoru kódu.
        /// </summary>
        private void InitializeComponent()
        {
            groupBoxInfoZarizeni = new GroupBox();
            lblInfoRozvadec = new Label();
            txtInfoRozvadec = new TextBox();
            lblInfoDruh = new Label();
            txtInfoDruh = new TextBox();
            lblInfoMenic = new Label();
            txtInfoMenic = new TextBox();
            lblInfoNapeti = new Label();
            txtInfoNapeti = new TextBox();
            lblInfoProud = new Label();
            txtInfoProud = new TextBox();
            lblInfoPrikon = new Label();
            txtInfoPrikon = new TextBox();
            lblInfoPopis = new Label();
            txtInfoPopis = new TextBox();
            lblInfoTag = new Label();
            txtInfoTag = new TextBox();
            panelSearch = new Panel();
            lblSearch = new Label();
            textBoxSearch = new TextBox();
            dataGridViewKabely = new DataGridView();
            panelButtons = new Panel();
            lblStatistika = new Label();
            btnAddKabel = new Button();
            btnCopyKabel = new Button();
            btnDeleteKabel = new Button();
            btnClose = new Button();
            lblZarizeniBezKabelu = new Label();
            comboBoxZarizeniBezKabelu = new ComboBox();
            btnAddKabelProZarizeni = new Button();
            groupBoxInfoZarizeni.SuspendLayout();
            panelSearch.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewKabely).BeginInit();
            panelButtons.SuspendLayout();
            SuspendLayout();
            // 
            // groupBoxInfoZarizeni
            // 
            groupBoxInfoZarizeni.Controls.Add(lblInfoRozvadec);
            groupBoxInfoZarizeni.Controls.Add(txtInfoRozvadec);
            groupBoxInfoZarizeni.Controls.Add(lblInfoDruh);
            groupBoxInfoZarizeni.Controls.Add(txtInfoDruh);
            groupBoxInfoZarizeni.Controls.Add(lblInfoMenic);
            groupBoxInfoZarizeni.Controls.Add(txtInfoMenic);
            groupBoxInfoZarizeni.Controls.Add(lblInfoNapeti);
            groupBoxInfoZarizeni.Controls.Add(txtInfoNapeti);
            groupBoxInfoZarizeni.Controls.Add(lblInfoProud);
            groupBoxInfoZarizeni.Controls.Add(txtInfoProud);
            groupBoxInfoZarizeni.Controls.Add(lblInfoPrikon);
            groupBoxInfoZarizeni.Controls.Add(txtInfoPrikon);
            groupBoxInfoZarizeni.Controls.Add(lblInfoPopis);
            groupBoxInfoZarizeni.Controls.Add(txtInfoPopis);
            groupBoxInfoZarizeni.Controls.Add(lblInfoTag);
            groupBoxInfoZarizeni.Controls.Add(txtInfoTag);
            groupBoxInfoZarizeni.Dock = DockStyle.Top;
            groupBoxInfoZarizeni.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            groupBoxInfoZarizeni.Location = new Point(10, 10);
            groupBoxInfoZarizeni.Name = "groupBoxInfoZarizeni";
            groupBoxInfoZarizeni.Size = new Size(1114, 150);
            groupBoxInfoZarizeni.TabIndex = 0;
            groupBoxInfoZarizeni.TabStop = false;
            groupBoxInfoZarizeni.Text = "Informace o vybraném zařízení (pouze pro čtení)";
            // 
            // lblInfoRozvadec
            // 
            lblInfoRozvadec.AutoSize = true;
            lblInfoRozvadec.Location = new Point(930, 85);
            lblInfoRozvadec.Name = "lblInfoRozvadec";
            lblInfoRozvadec.Size = new Size(86, 23);
            lblInfoRozvadec.TabIndex = 14;
            lblInfoRozvadec.Text = "Rozvaděč:";
            // 
            // txtInfoRozvadec
            // 
            txtInfoRozvadec.BackColor = SystemColors.Control;
            txtInfoRozvadec.Location = new Point(930, 110);
            txtInfoRozvadec.Name = "txtInfoRozvadec";
            txtInfoRozvadec.ReadOnly = true;
            txtInfoRozvadec.Size = new Size(160, 30);
            txtInfoRozvadec.TabIndex = 15;
            // 
            // lblInfoDruh
            // 
            lblInfoDruh.AutoSize = true;
            lblInfoDruh.Location = new Point(750, 85);
            lblInfoDruh.Name = "lblInfoDruh";
            lblInfoDruh.Size = new Size(51, 23);
            lblInfoDruh.TabIndex = 12;
            lblInfoDruh.Text = "Druh:";
            // 
            // txtInfoDruh
            // 
            txtInfoDruh.BackColor = SystemColors.Control;
            txtInfoDruh.Location = new Point(750, 110);
            txtInfoDruh.Name = "txtInfoDruh";
            txtInfoDruh.ReadOnly = true;
            txtInfoDruh.Size = new Size(160, 30);
            txtInfoDruh.TabIndex = 13;
            // 
            // lblInfoMenic
            // 
            lblInfoMenic.AutoSize = true;
            lblInfoMenic.Location = new Point(570, 85);
            lblInfoMenic.Name = "lblInfoMenic";
            lblInfoMenic.Size = new Size(71, 23);
            lblInfoMenic.TabIndex = 10;
            lblInfoMenic.Text = "Měnič?:";
            // 
            // txtInfoMenic
            // 
            txtInfoMenic.BackColor = SystemColors.Control;
            txtInfoMenic.Location = new Point(570, 110);
            txtInfoMenic.Name = "txtInfoMenic";
            txtInfoMenic.ReadOnly = true;
            txtInfoMenic.Size = new Size(160, 30);
            txtInfoMenic.TabIndex = 11;
            // 
            // lblInfoNapeti
            // 
            lblInfoNapeti.AutoSize = true;
            lblInfoNapeti.Location = new Point(390, 85);
            lblInfoNapeti.Name = "lblInfoNapeti";
            lblInfoNapeti.Size = new Size(99, 23);
            lblInfoNapeti.TabIndex = 8;
            lblInfoNapeti.Text = "Napětí [V]:";
            // 
            // txtInfoNapeti
            // 
            txtInfoNapeti.BackColor = SystemColors.Control;
            txtInfoNapeti.Location = new Point(390, 110);
            txtInfoNapeti.Name = "txtInfoNapeti";
            txtInfoNapeti.ReadOnly = true;
            txtInfoNapeti.Size = new Size(160, 30);
            txtInfoNapeti.TabIndex = 9;
            // 
            // lblInfoProud
            // 
            lblInfoProud.AutoSize = true;
            lblInfoProud.Location = new Point(210, 85);
            lblInfoProud.Name = "lblInfoProud";
            lblInfoProud.Size = new Size(88, 23);
            lblInfoProud.TabIndex = 6;
            lblInfoProud.Text = "Proud [A]:";
            // 
            // txtInfoProud
            // 
            txtInfoProud.BackColor = SystemColors.Control;
            txtInfoProud.Location = new Point(210, 110);
            txtInfoProud.Name = "txtInfoProud";
            txtInfoProud.ReadOnly = true;
            txtInfoProud.Size = new Size(160, 30);
            txtInfoProud.TabIndex = 7;
            // 
            // lblInfoPrikon
            // 
            lblInfoPrikon.AutoSize = true;
            lblInfoPrikon.Location = new Point(20, 85);
            lblInfoPrikon.Name = "lblInfoPrikon";
            lblInfoPrikon.Size = new Size(106, 23);
            lblInfoPrikon.TabIndex = 4;
            lblInfoPrikon.Text = "Příkon [kW]:";
            // 
            // txtInfoPrikon
            // 
            txtInfoPrikon.BackColor = SystemColors.Control;
            txtInfoPrikon.Location = new Point(20, 110);
            txtInfoPrikon.Name = "txtInfoPrikon";
            txtInfoPrikon.ReadOnly = true;
            txtInfoPrikon.Size = new Size(170, 30);
            txtInfoPrikon.TabIndex = 5;
            // 
            // lblInfoPopis
            // 
            lblInfoPopis.AutoSize = true;
            lblInfoPopis.Location = new Point(210, 25);
            lblInfoPopis.Name = "lblInfoPopis";
            lblInfoPopis.Size = new Size(125, 23);
            lblInfoPopis.TabIndex = 2;
            lblInfoPopis.Text = "Popis zařízení:";
            // 
            // txtInfoPopis
            // 
            txtInfoPopis.BackColor = SystemColors.Control;
            txtInfoPopis.Location = new Point(210, 50);
            txtInfoPopis.Name = "txtInfoPopis";
            txtInfoPopis.ReadOnly = true;
            txtInfoPopis.Size = new Size(880, 30);
            txtInfoPopis.TabIndex = 3;
            // 
            // lblInfoTag
            // 
            lblInfoTag.AutoSize = true;
            lblInfoTag.Location = new Point(20, 25);
            lblInfoTag.Name = "lblInfoTag";
            lblInfoTag.Size = new Size(129, 23);
            lblInfoTag.TabIndex = 0;
            lblInfoTag.Text = "Označení (Tag):";
            // 
            // txtInfoTag
            // 
            txtInfoTag.BackColor = SystemColors.Control;
            txtInfoTag.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            txtInfoTag.Location = new Point(20, 50);
            txtInfoTag.Name = "txtInfoTag";
            txtInfoTag.ReadOnly = true;
            txtInfoTag.Size = new Size(170, 30);
            txtInfoTag.TabIndex = 1;
            // 
            // 
            // panelSearch
            // 
            panelSearch.Controls.Add(lblSearch);
            panelSearch.Controls.Add(textBoxSearch);
            panelSearch.Controls.Add(lblZarizeniBezKabelu);
            panelSearch.Controls.Add(comboBoxZarizeniBezKabelu);
            panelSearch.Controls.Add(btnAddKabelProZarizeni);
            panelSearch.Dock = DockStyle.Top;
            panelSearch.Location = new Point(10, 160);
            panelSearch.Name = "panelSearch";
            panelSearch.Size = new Size(1114, 50);
            panelSearch.TabIndex = 1;
            // 
            // lblSearch
            // 
            lblSearch.AutoSize = true;
            lblSearch.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            lblSearch.Location = new Point(20, 15);
            lblSearch.Name = "lblSearch";
            lblSearch.Size = new Size(149, 23);
            lblSearch.TabIndex = 0;
            lblSearch.Text = "Hledat (filtr):";
            // 
            // textBoxSearch
            // 
            textBoxSearch.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            textBoxSearch.Location = new Point(120, 12);
            textBoxSearch.Name = "textBoxSearch";
            textBoxSearch.Size = new Size(350, 30);
            textBoxSearch.TabIndex = 1;
            textBoxSearch.TextChanged += TextBoxSearch_TextChanged;
            // 
            // lblZarizeniBezKabelu
            // 
            lblZarizeniBezKabelu.AutoSize = true;
            lblZarizeniBezKabelu.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            lblZarizeniBezKabelu.Location = new Point(500, 15);
            lblZarizeniBezKabelu.Name = "lblZarizeniBezKabelu";
            lblZarizeniBezKabelu.Size = new Size(106, 23);
            lblZarizeniBezKabelu.TabIndex = 2;
            lblZarizeniBezKabelu.Text = "Bez kabelu:";
            // 
            // comboBoxZarizeniBezKabelu
            // 
            comboBoxZarizeniBezKabelu.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxZarizeniBezKabelu.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            comboBoxZarizeniBezKabelu.FormattingEnabled = true;
            comboBoxZarizeniBezKabelu.Location = new Point(610, 12);
            comboBoxZarizeniBezKabelu.Name = "comboBoxZarizeniBezKabelu";
            comboBoxZarizeniBezKabelu.Size = new Size(320, 31);
            comboBoxZarizeniBezKabelu.TabIndex = 3;
            // 
            // btnAddKabelProZarizeni
            // 
            btnAddKabelProZarizeni.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            btnAddKabelProZarizeni.Location = new Point(945, 8);
            btnAddKabelProZarizeni.Name = "btnAddKabelProZarizeni";
            btnAddKabelProZarizeni.Size = new Size(145, 33);
            btnAddKabelProZarizeni.TabIndex = 4;
            btnAddKabelProZarizeni.Text = "Přidat kabel";
            btnAddKabelProZarizeni.UseVisualStyleBackColor = true;
            btnAddKabelProZarizeni.Click += BtnAddKabelProZarizeni_Click;
            // 
            // dataGridViewKabely
            // 
            dataGridViewKabely.AllowUserToAddRows = false;
            dataGridViewKabely.AllowUserToDeleteRows = false;
            dataGridViewKabely.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewKabely.Dock = DockStyle.Fill;
            dataGridViewKabely.Location = new Point(10, 210);
            dataGridViewKabely.Name = "dataGridViewKabely";
            dataGridViewKabely.RowHeadersWidth = 51;
            dataGridViewKabely.Size = new Size(1114, 381);
            dataGridViewKabely.TabIndex = 2;
            dataGridViewKabely.SelectionChanged += DataGridViewKabely_SelectionChanged;
            // 
            // panelButtons
            // 
            panelButtons.Controls.Add(lblStatistika);
            panelButtons.Controls.Add(btnAddKabel);
            panelButtons.Controls.Add(btnCopyKabel);
            panelButtons.Controls.Add(btnDeleteKabel);
            panelButtons.Controls.Add(btnClose);
            panelButtons.Dock = DockStyle.Bottom;
            panelButtons.Location = new Point(10, 591);
            panelButtons.Name = "panelButtons";
            panelButtons.Size = new Size(1114, 60);
            panelButtons.TabIndex = 3;
            // 
            // lblStatistika
            // 
            lblStatistika.AutoSize = true;
            lblStatistika.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            lblStatistika.Location = new Point(20, 20);
            lblStatistika.Name = "lblStatistika";
            lblStatistika.Size = new Size(157, 23);
            lblStatistika.TabIndex = 0;
            lblStatistika.Text = "Počet kabelů celkem: 0";
            // 
            // btnAddKabel
            // 
            btnAddKabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnAddKabel.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            btnAddKabel.Location = new Point(530, 12);
            btnAddKabel.Name = "btnAddKabel";
            btnAddKabel.Size = new Size(130, 36);
            btnAddKabel.TabIndex = 1;
            btnAddKabel.Text = "Přidat kabel";
            btnAddKabel.UseVisualStyleBackColor = true;
            btnAddKabel.Click += BtnAddKabel_Click;
            // 
            // btnCopyKabel
            // 
            btnCopyKabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnCopyKabel.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            btnCopyKabel.Location = new Point(670, 12);
            btnCopyKabel.Name = "btnCopyKabel";
            btnCopyKabel.Size = new Size(150, 36);
            btnCopyKabel.TabIndex = 2;
            btnCopyKabel.Text = "Kopírovat vybraný";
            btnCopyKabel.UseVisualStyleBackColor = true;
            btnCopyKabel.Click += BtnCopyKabel_Click;
            // 
            // btnDeleteKabel
            // 
            btnDeleteKabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnDeleteKabel.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            btnDeleteKabel.Location = new Point(830, 12);
            btnDeleteKabel.Name = "btnDeleteKabel";
            btnDeleteKabel.Size = new Size(140, 36);
            btnDeleteKabel.TabIndex = 3;
            btnDeleteKabel.Text = "Smazat vybraný";
            btnDeleteKabel.UseVisualStyleBackColor = true;
            btnDeleteKabel.Click += BtnDeleteKabel_Click;
            // 
            // btnClose
            // 
            btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnClose.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            btnClose.Location = new Point(980, 12);
            btnClose.Name = "btnClose";
            btnClose.Size = new Size(110, 36);
            btnClose.TabIndex = 4;
            btnClose.Text = "Zavřít";
            btnClose.UseVisualStyleBackColor = true;
            btnClose.Click += BtnClose_Click;
            // 
            // FormVsechnyKabely
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1134, 661);
            Controls.Add(dataGridViewKabely);
            Controls.Add(panelButtons);
            Controls.Add(panelSearch);
            Controls.Add(groupBoxInfoZarizeni);
            Name = "FormVsechnyKabely";
            Padding = new Padding(10);
            StartPosition = FormStartPosition.CenterParent;
            Text = "Hromadná správa všech kabelů";
            Load += FormVsechnyKabely_Load;
            groupBoxInfoZarizeni.ResumeLayout(false);
            groupBoxInfoZarizeni.PerformLayout();
            panelSearch.ResumeLayout(false);
            panelSearch.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewKabely).EndInit();
            panelButtons.ResumeLayout(false);
            panelButtons.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxInfoZarizeni;
        private System.Windows.Forms.Label lblInfoTag;
        private System.Windows.Forms.TextBox txtInfoTag;
        private System.Windows.Forms.Label lblInfoPopis;
        private System.Windows.Forms.TextBox txtInfoPopis;
        private System.Windows.Forms.Label lblInfoPrikon;
        private System.Windows.Forms.TextBox txtInfoPrikon;
        private System.Windows.Forms.Label lblInfoProud;
        private System.Windows.Forms.TextBox txtInfoProud;
        private System.Windows.Forms.Label lblInfoNapeti;
        private System.Windows.Forms.TextBox txtInfoNapeti;
        private System.Windows.Forms.Label lblInfoMenic;
        private System.Windows.Forms.TextBox txtInfoMenic;
        private System.Windows.Forms.Label lblInfoDruh;
        private System.Windows.Forms.TextBox txtInfoDruh;
        private System.Windows.Forms.Label lblInfoRozvadec;
        private System.Windows.Forms.TextBox txtInfoRozvadec;
        private System.Windows.Forms.Panel panelSearch;
        private System.Windows.Forms.Label lblSearch;
        private System.Windows.Forms.TextBox textBoxSearch;
        private System.Windows.Forms.DataGridView dataGridViewKabely;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Label lblStatistika;
        private System.Windows.Forms.Button btnAddKabel;
        private System.Windows.Forms.Button btnCopyKabel;
        private System.Windows.Forms.Button btnDeleteKabel;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblZarizeniBezKabelu;
        private System.Windows.Forms.ComboBox comboBoxZarizeniBezKabelu;
        private System.Windows.Forms.Button btnAddKabelProZarizeni;
    }
}
