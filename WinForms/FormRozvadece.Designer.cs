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

        private void InitializeComponent()
        {
            this.listBoxRozvadece = new System.Windows.Forms.ListBox();
            this.groupBoxNovy = new System.Windows.Forms.GroupBox();
            this.btnPridatRozvadec = new System.Windows.Forms.Button();
            this.txtNovyRozvadec = new System.Windows.Forms.TextBox();
            this.groupBoxPrirazeno = new System.Windows.Forms.GroupBox();
            this.dataGridViewPrirazeno = new System.Windows.Forms.DataGridView();
            this.groupBoxNeprirazeno = new System.Windows.Forms.GroupBox();
            this.dataGridViewNeprirazeno = new System.Windows.Forms.DataGridView();
            this.btnPriradit = new System.Windows.Forms.Button();
            this.btnOdebrat = new System.Windows.Forms.Button();
            this.lblStatistika = new System.Windows.Forms.Label();
            this.btnZavrit = new System.Windows.Forms.Button();
            this.groupBoxNovy.SuspendLayout();
            this.groupBoxPrirazeno.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPrirazeno)).BeginInit();
            this.groupBoxNeprirazeno.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNeprirazeno)).BeginInit();
            this.SuspendLayout();
            // 
            // listBoxRozvadece
            // 
            this.listBoxRozvadece.FormattingEnabled = true;
            this.listBoxRozvadece.ItemHeight = 21;
            this.listBoxRozvadece.Location = new System.Drawing.Point(12, 12);
            this.listBoxRozvadece.Name = "listBoxRozvadece";
            this.listBoxRozvadece.Size = new System.Drawing.Size(220, 382);
            this.listBoxRozvadece.TabIndex = 0;
            this.listBoxRozvadece.SelectedIndexChanged += new System.EventHandler(this.ListBoxRozvadece_SelectedIndexChanged);
            // 
            // groupBoxNovy
            // 
            this.groupBoxNovy.Controls.Add(this.btnPridatRozvadec);
            this.groupBoxNovy.Controls.Add(this.txtNovyRozvadec);
            this.groupBoxNovy.Location = new System.Drawing.Point(12, 400);
            this.groupBoxNovy.Name = "groupBoxNovy";
            this.groupBoxNovy.Size = new System.Drawing.Size(220, 110);
            this.groupBoxNovy.TabIndex = 1;
            this.groupBoxNovy.TabStop = false;
            this.groupBoxNovy.Text = "Nový rozvaděč";
            // 
            // btnPridatRozvadec
            // 
            this.btnPridatRozvadec.Location = new System.Drawing.Point(6, 68);
            this.btnPridatRozvadec.Name = "btnPridatRozvadec";
            this.btnPridatRozvadec.Size = new System.Drawing.Size(208, 32);
            this.btnPridatRozvadec.TabIndex = 1;
            this.btnPridatRozvadec.Text = "Přidat rozvaděč";
            this.btnPridatRozvadec.UseVisualStyleBackColor = true;
            this.btnPridatRozvadec.Click += new System.EventHandler(this.BtnPridatRozvadec_Click);
            // 
            // txtNovyRozvadec
            // 
            this.txtNovyRozvadec.Location = new System.Drawing.Point(6, 28);
            this.txtNovyRozvadec.Name = "txtNovyRozvadec";
            this.txtNovyRozvadec.Size = new System.Drawing.Size(208, 29);
            this.txtNovyRozvadec.TabIndex = 0;
            // 
            // groupBoxPrirazeno
            // 
            this.groupBoxPrirazeno.Controls.Add(this.dataGridViewPrirazeno);
            this.groupBoxPrirazeno.Location = new System.Drawing.Point(248, 12);
            this.groupBoxPrirazeno.Name = "groupBoxPrirazeno";
            this.groupBoxPrirazeno.Size = new System.Drawing.Size(824, 218);
            this.groupBoxPrirazeno.TabIndex = 2;
            this.groupBoxPrirazeno.TabStop = false;
            this.groupBoxPrirazeno.Text = "Zařízení přiřazená k vybranému rozvaděči";
            // 
            // dataGridViewPrirazeno
            // 
            this.dataGridViewPrirazeno.AllowUserToAddRows = false;
            this.dataGridViewPrirazeno.AllowUserToDeleteRows = false;
            this.dataGridViewPrirazeno.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewPrirazeno.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewPrirazeno.Location = new System.Drawing.Point(3, 25);
            this.dataGridViewPrirazeno.Name = "dataGridViewPrirazeno";
            this.dataGridViewPrirazeno.ReadOnly = true;
            this.dataGridViewPrirazeno.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewPrirazeno.Size = new System.Drawing.Size(818, 190);
            this.dataGridViewPrirazeno.TabIndex = 0;
            // 
            // groupBoxNeprirazeno
            // 
            this.groupBoxNeprirazeno.Controls.Add(this.dataGridViewNeprirazeno);
            this.groupBoxNeprirazeno.Location = new System.Drawing.Point(248, 292);
            this.groupBoxNeprirazeno.Name = "groupBoxNeprirazeno";
            this.groupBoxNeprirazeno.Size = new System.Drawing.Size(824, 218);
            this.groupBoxNeprirazeno.TabIndex = 3;
            this.groupBoxNeprirazeno.TabStop = false;
            this.groupBoxNeprirazeno.Text = "Nepřiřazená zařízení (chybí rozvaděč)";
            // 
            // dataGridViewNeprirazeno
            // 
            this.dataGridViewNeprirazeno.AllowUserToAddRows = false;
            this.dataGridViewNeprirazeno.AllowUserToDeleteRows = false;
            this.dataGridViewNeprirazeno.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewNeprirazeno.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewNeprirazeno.Location = new System.Drawing.Point(3, 25);
            this.dataGridViewNeprirazeno.Name = "dataGridViewNeprirazeno";
            this.dataGridViewNeprirazeno.ReadOnly = true;
            this.dataGridViewNeprirazeno.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewNeprirazeno.Size = new System.Drawing.Size(818, 190);
            this.dataGridViewNeprirazeno.TabIndex = 0;
            // 
            // btnPriradit
            // 
            this.btnPriradit.Location = new System.Drawing.Point(400, 246);
            this.btnPriradit.Name = "btnPriradit";
            this.btnPriradit.Size = new System.Drawing.Size(240, 36);
            this.btnPriradit.TabIndex = 4;
            this.btnPriradit.Text = "▲ Přiřadit vybrané k rozvaděči";
            this.btnPriradit.UseVisualStyleBackColor = true;
            this.btnPriradit.Click += new System.EventHandler(this.BtnPriradit_Click);
            // 
            // btnOdebrat
            // 
            this.btnOdebrat.Location = new System.Drawing.Point(660, 246);
            this.btnOdebrat.Name = "btnOdebrat";
            this.btnOdebrat.Size = new System.Drawing.Size(240, 36);
            this.btnOdebrat.TabIndex = 5;
            this.btnOdebrat.Text = "▼ Odebrat vybrané z rozvaděče";
            this.btnOdebrat.UseVisualStyleBackColor = true;
            this.btnOdebrat.Click += new System.EventHandler(this.BtnOdebrat_Click);
            // 
            // lblStatistika
            // 
            this.lblStatistika.AutoSize = true;
            this.lblStatistika.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold);
            this.lblStatistika.Location = new System.Drawing.Point(12, 532);
            this.lblStatistika.Name = "lblStatistika";
            this.lblStatistika.Size = new System.Drawing.Size(324, 20);
            this.lblStatistika.TabIndex = 6;
            this.lblStatistika.Text = "Statistika: Přiřazeno: 0  |  Chybí přiřadit: 0";
            // 
            // btnZavrit
            // 
            this.btnZavrit.Location = new System.Drawing.Point(922, 525);
            this.btnZavrit.Name = "btnZavrit";
            this.btnZavrit.Size = new System.Drawing.Size(150, 32);
            this.btnZavrit.TabIndex = 7;
            this.btnZavrit.Text = "Zavřít";
            this.btnZavrit.UseVisualStyleBackColor = true;
            this.btnZavrit.Click += new System.EventHandler(this.BtnZavrit_Click);
            // 
            // FormRozvadece
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1084, 569);
            this.Controls.Add(this.btnZavrit);
            this.Controls.Add(this.lblStatistika);
            this.Controls.Add(this.btnOdebrat);
            this.Controls.Add(this.btnPriradit);
            this.Controls.Add(this.groupBoxNeprirazeno);
            this.Controls.Add(this.groupBoxPrirazeno);
            this.Controls.Add(this.groupBoxNovy);
            this.Controls.Add(this.listBoxRozvadece);
            this.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormRozvadece";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Přiřazení zařízení k rozvaděčům";
            this.Load += new System.EventHandler(this.FormRozvadece_Load);
            this.groupBoxNovy.ResumeLayout(false);
            this.groupBoxNovy.PerformLayout();
            this.groupBoxPrirazeno.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPrirazeno)).EndInit();
            this.groupBoxNeprirazeno.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewNeprirazeno)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
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
