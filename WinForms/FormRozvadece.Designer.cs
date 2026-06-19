namespace WinForms
{
    partial class FormRozvadece
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent() {
            listBoxRozvadece = new ListBox();
            groupBoxNovy = new GroupBox();
            btnPridatRozvadec = new Button();
            txtNovyRozvadec = new TextBox();
            groupBoxPrirazeno = new GroupBox();
            dataGridViewPrirazeno = new DataGridView();
            groupBoxNeprirazeno = new GroupBox();
            dataGridViewNeprirazeno = new DataGridView();
            btnPriradit = new Button();
            btnOdebrat = new Button();
            lblStatistika = new Label();
            btnZavrit = new Button();
            groupBoxNovy.SuspendLayout();
            groupBoxPrirazeno.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewPrirazeno).BeginInit();
            groupBoxNeprirazeno.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewNeprirazeno).BeginInit();
            SuspendLayout();
            // 
            // listBoxRozvadece
            // 
            listBoxRozvadece.FormattingEnabled = true;
            listBoxRozvadece.Location = new Point(12, 12);
            listBoxRozvadece.Name = "listBoxRozvadece";
            listBoxRozvadece.Size = new Size(220, 368);
            listBoxRozvadece.TabIndex = 0;
            listBoxRozvadece.SelectedIndexChanged += ListBoxRozvadece_SelectedIndexChanged;
            // 
            // groupBoxNovy
            // 
            groupBoxNovy.Controls.Add(btnPridatRozvadec);
            groupBoxNovy.Controls.Add(txtNovyRozvadec);
            groupBoxNovy.Location = new Point(12, 400);
            groupBoxNovy.Name = "groupBoxNovy";
            groupBoxNovy.Size = new Size(220, 110);
            groupBoxNovy.TabIndex = 1;
            groupBoxNovy.TabStop = false;
            groupBoxNovy.Text = "Nový rozvaděč";
            // 
            // btnPridatRozvadec
            // 
            btnPridatRozvadec.Location = new Point(6, 68);
            btnPridatRozvadec.Name = "btnPridatRozvadec";
            btnPridatRozvadec.Size = new Size(208, 32);
            btnPridatRozvadec.TabIndex = 1;
            btnPridatRozvadec.Text = "Přidat rozvaděč";
            btnPridatRozvadec.UseVisualStyleBackColor = true;
            btnPridatRozvadec.Click += BtnPridatRozvadec_Click;
            // 
            // txtNovyRozvadec
            // 
            txtNovyRozvadec.Location = new Point(6, 28);
            txtNovyRozvadec.Name = "txtNovyRozvadec";
            txtNovyRozvadec.Size = new Size(208, 34);
            txtNovyRozvadec.TabIndex = 0;
            // 
            // groupBoxPrirazeno
            // 
            groupBoxPrirazeno.Controls.Add(dataGridViewPrirazeno);
            groupBoxPrirazeno.Location = new Point(248, 12);
            groupBoxPrirazeno.Name = "groupBoxPrirazeno";
            groupBoxPrirazeno.Size = new Size(824, 218);
            groupBoxPrirazeno.TabIndex = 2;
            groupBoxPrirazeno.TabStop = false;
            groupBoxPrirazeno.Text = "Zařízení přiřazená k vybranému rozvaděči";
            // 
            // dataGridViewPrirazeno
            // 
            dataGridViewPrirazeno.AllowUserToAddRows = false;
            dataGridViewPrirazeno.AllowUserToDeleteRows = false;
            dataGridViewPrirazeno.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewPrirazeno.Dock = DockStyle.Fill;
            dataGridViewPrirazeno.Location = new Point(3, 30);
            dataGridViewPrirazeno.Name = "dataGridViewPrirazeno";
            dataGridViewPrirazeno.ReadOnly = true;
            dataGridViewPrirazeno.RowHeadersWidth = 51;
            dataGridViewPrirazeno.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewPrirazeno.Size = new Size(818, 185);
            dataGridViewPrirazeno.TabIndex = 0;
            // 
            // groupBoxNeprirazeno
            // 
            groupBoxNeprirazeno.Controls.Add(dataGridViewNeprirazeno);
            groupBoxNeprirazeno.Location = new Point(248, 292);
            groupBoxNeprirazeno.Name = "groupBoxNeprirazeno";
            groupBoxNeprirazeno.Size = new Size(824, 218);
            groupBoxNeprirazeno.TabIndex = 3;
            groupBoxNeprirazeno.TabStop = false;
            groupBoxNeprirazeno.Text = "Nepřiřazená zařízení (chybí rozvaděč)";
            // 
            // dataGridViewNeprirazeno
            // 
            dataGridViewNeprirazeno.AllowUserToAddRows = false;
            dataGridViewNeprirazeno.AllowUserToDeleteRows = false;
            dataGridViewNeprirazeno.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewNeprirazeno.Dock = DockStyle.Fill;
            dataGridViewNeprirazeno.Location = new Point(3, 30);
            dataGridViewNeprirazeno.Name = "dataGridViewNeprirazeno";
            dataGridViewNeprirazeno.ReadOnly = true;
            dataGridViewNeprirazeno.RowHeadersWidth = 51;
            dataGridViewNeprirazeno.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewNeprirazeno.Size = new Size(818, 185);
            dataGridViewNeprirazeno.TabIndex = 0;
            dataGridViewNeprirazeno.CellContentClick += dataGridViewNeprirazeno_CellContentClick;
            // 
            // btnPriradit
            // 
            btnPriradit.Location = new Point(400, 246);
            btnPriradit.Name = "btnPriradit";
            btnPriradit.Size = new Size(240, 36);
            btnPriradit.TabIndex = 4;
            btnPriradit.Text = "▲ Přiřadit vybrané k rozvaděči";
            btnPriradit.UseVisualStyleBackColor = true;
            btnPriradit.Click += BtnPriradit_Click;
            // 
            // btnOdebrat
            // 
            btnOdebrat.Location = new Point(660, 246);
            btnOdebrat.Name = "btnOdebrat";
            btnOdebrat.Size = new Size(240, 36);
            btnOdebrat.TabIndex = 5;
            btnOdebrat.Text = "▼ Odebrat vybrané z rozvaděče";
            btnOdebrat.UseVisualStyleBackColor = true;
            btnOdebrat.Click += BtnOdebrat_Click;
            // 
            // lblStatistika
            // 
            lblStatistika.AutoSize = true;
            lblStatistika.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            lblStatistika.Location = new Point(12, 532);
            lblStatistika.Name = "lblStatistika";
            lblStatistika.Size = new Size(379, 25);
            lblStatistika.TabIndex = 6;
            lblStatistika.Text = "Statistika: Přiřazeno: 0  |  Chybí přiřadit: 0";
            // 
            // btnZavrit
            // 
            btnZavrit.Location = new Point(922, 525);
            btnZavrit.Name = "btnZavrit";
            btnZavrit.Size = new Size(150, 32);
            btnZavrit.TabIndex = 7;
            btnZavrit.Text = "Zavřít";
            btnZavrit.UseVisualStyleBackColor = true;
            btnZavrit.Click += BtnZavrit_Click;
            // 
            // FormRozvadece
            // 
            AutoScaleDimensions = new SizeF(11F, 28F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1084, 569);
            Controls.Add(btnZavrit);
            Controls.Add(lblStatistika);
            Controls.Add(btnOdebrat);
            Controls.Add(btnPriradit);
            Controls.Add(groupBoxNeprirazeno);
            Controls.Add(groupBoxPrirazeno);
            Controls.Add(groupBoxNovy);
            Controls.Add(listBoxRozvadece);
            Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 238);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Margin = new Padding(4);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FormRozvadece";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "Přiřazení zařízení k rozvaděčům";
            Load += FormRozvadece_Load;
            groupBoxNovy.ResumeLayout(false);
            groupBoxNovy.PerformLayout();
            groupBoxPrirazeno.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewPrirazeno).EndInit();
            groupBoxNeprirazeno.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewNeprirazeno).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.ListBox listBoxRozvadece;
        private System.Windows.Forms.GroupBox groupBoxNovy;
        private System.Windows.Forms.Button btnPridatRozvadec;
        private System.Windows.Forms.TextBox txtNovyRozvadec;
        private System.Windows.Forms.GroupBox groupBoxPrirazeno;
        private System.Windows.Forms.DataGridView dataGridViewPrirazeno;
        private System.Windows.Forms.GroupBox groupBoxNeprirazeno;
        private System.Windows.Forms.DataGridView dataGridViewNeprirazeno;
        private System.Windows.Forms.Button btnPriradit;
        private System.Windows.Forms.Button btnOdebrat;
        private System.Windows.Forms.Label lblStatistika;
        private System.Windows.Forms.Button btnZavrit;
    }
}
