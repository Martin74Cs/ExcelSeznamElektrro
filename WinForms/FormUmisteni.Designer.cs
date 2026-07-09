namespace WinForms
{
    partial class FormUmisteni
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
            splitContainerMain = new SplitContainer();
            labelVlastnost = new Label();
            comboBoxVlastnost = new ComboBox();
            listBoxHodnoty = new ListBox();
            groupBoxNovy = new GroupBox();
            btnSloucit = new Button();
            btnPridat = new Button();
            txtNovaHodnota = new TextBox();
            lblStatistika = new Label();
            btnZavrit = new Button();
            splitContainerRight = new SplitContainer();
            groupBoxPrirazeno = new GroupBox();
            dataGridViewPrirazeno = new DataGridView();
            panelButtons = new Panel();
            btnOdebrat = new Button();
            btnPriradit = new Button();
            groupBoxNeprirazeno = new GroupBox();
            dataGridViewNeprirazeno = new DataGridView();
            ((System.ComponentModel.ISupportInitialize)splitContainerMain).BeginInit();
            splitContainerMain.Panel1.SuspendLayout();
            splitContainerMain.Panel2.SuspendLayout();
            splitContainerMain.SuspendLayout();
            groupBoxNovy.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainerRight).BeginInit();
            splitContainerRight.Panel1.SuspendLayout();
            splitContainerRight.Panel2.SuspendLayout();
            splitContainerRight.SuspendLayout();
            groupBoxPrirazeno.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewPrirazeno).BeginInit();
            panelButtons.SuspendLayout();
            groupBoxNeprirazeno.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewNeprirazeno).BeginInit();
            SuspendLayout();
            // 
            // splitContainerMain
            // 
            splitContainerMain.Dock = DockStyle.Fill;
            splitContainerMain.Location = new System.Drawing.Point(0, 0);
            splitContainerMain.Name = "splitContainerMain";
            // 
            // splitContainerMain.Panel1
            // 
            splitContainerMain.Panel1.Controls.Add(labelVlastnost);
            splitContainerMain.Panel1.Controls.Add(comboBoxVlastnost);
            splitContainerMain.Panel1.Controls.Add(listBoxHodnoty);
            splitContainerMain.Panel1.Controls.Add(groupBoxNovy);
            splitContainerMain.Panel1.Controls.Add(lblStatistika);
            splitContainerMain.Panel1.Controls.Add(btnZavrit);
            splitContainerMain.Panel1MinSize = 240;
            // 
            // splitContainerMain.Panel2
            // 
            splitContainerMain.Panel2.Controls.Add(splitContainerRight);
            splitContainerMain.Panel2MinSize = 400;
            splitContainerMain.Size = new System.Drawing.Size(1100, 650);
            splitContainerMain.SplitterDistance = 260;
            splitContainerMain.TabIndex = 0;
            // 
            // labelVlastnost
            // 
            labelVlastnost.AutoSize = true;
            labelVlastnost.Location = new System.Drawing.Point(12, 9);
            labelVlastnost.Name = "labelVlastnost";
            labelVlastnost.Size = new System.Drawing.Size(157, 21);
            labelVlastnost.TabIndex = 0;
            labelVlastnost.Text = "Spravovaná vlastnost:";
            // 
            // comboBoxVlastnost
            // 
            comboBoxVlastnost.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            comboBoxVlastnost.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxVlastnost.FormattingEnabled = true;
            comboBoxVlastnost.Location = new System.Drawing.Point(12, 33);
            comboBoxVlastnost.Name = "comboBoxVlastnost";
            comboBoxVlastnost.Size = new System.Drawing.Size(236, 29);
            comboBoxVlastnost.TabIndex = 1;
            comboBoxVlastnost.SelectedIndexChanged += ComboBoxVlastnost_SelectedIndexChanged;
            // 
            // listBoxHodnoty
            // 
            listBoxHodnoty.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listBoxHodnoty.FormattingEnabled = true;
            listBoxHodnoty.ItemHeight = 21;
            listBoxHodnoty.Location = new System.Drawing.Point(12, 75);
            listBoxHodnoty.Name = "listBoxHodnoty";
            listBoxHodnoty.Size = new System.Drawing.Size(236, 319);
            listBoxHodnoty.TabIndex = 2;
            listBoxHodnoty.SelectedIndexChanged += ListBoxHodnoty_SelectedIndexChanged;
            // 
            // groupBoxNovy
            // 
            groupBoxNovy.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxNovy.Controls.Add(btnSloucit);
            groupBoxNovy.Controls.Add(btnPridat);
            groupBoxNovy.Controls.Add(txtNovaHodnota);
            groupBoxNovy.Location = new System.Drawing.Point(12, 405);
            groupBoxNovy.Name = "groupBoxNovy";
            groupBoxNovy.Size = new System.Drawing.Size(236, 150);
            groupBoxNovy.TabIndex = 3;
            groupBoxNovy.TabStop = false;
            groupBoxNovy.Text = "Úprava / Nový";
            // 
            // btnSloucit
            // 
            btnSloucit.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnSloucit.Location = new System.Drawing.Point(6, 108);
            btnSloucit.Name = "btnSloucit";
            btnSloucit.Size = new System.Drawing.Size(224, 34);
            btnSloucit.TabIndex = 2;
            btnSloucit.Text = "Přejmenovat / Sloučit";
            btnSloucit.UseVisualStyleBackColor = true;
            btnSloucit.Click += BtnSloucit_Click;
            // 
            // btnPridat
            // 
            btnPridat.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            btnPridat.Location = new System.Drawing.Point(6, 68);
            btnPridat.Name = "btnPridat";
            btnPridat.Size = new System.Drawing.Size(224, 34);
            btnPridat.TabIndex = 1;
            btnPridat.Text = "Přidat hodnotu";
            btnPridat.UseVisualStyleBackColor = true;
            btnPridat.Click += BtnPridat_Click;
            // 
            // txtNovaHodnota
            // 
            txtNovaHodnota.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtNovaHodnota.Location = new System.Drawing.Point(6, 28);
            txtNovaHodnota.Name = "txtNovaHodnota";
            txtNovaHodnota.Size = new System.Drawing.Size(224, 29);
            txtNovaHodnota.TabIndex = 0;
            // 
            // lblStatistika
            // 
            lblStatistika.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            lblStatistika.AutoSize = true;
            lblStatistika.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            lblStatistika.Location = new System.Drawing.Point(12, 570);
            lblStatistika.Name = "lblStatistika";
            lblStatistika.Size = new System.Drawing.Size(298, 19);
            lblStatistika.TabIndex = 4;
            lblStatistika.Text = "Statistika: Přiřazeno: 0  |  Nepřiřazeno: 0";
            // 
            // btnZavrit
            // 
            btnZavrit.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            btnZavrit.Location = new System.Drawing.Point(12, 603);
            btnZavrit.Name = "btnZavrit";
            btnZavrit.Size = new System.Drawing.Size(236, 34);
            btnZavrit.TabIndex = 5;
            btnZavrit.Text = "Zavřít";
            btnZavrit.UseVisualStyleBackColor = true;
            btnZavrit.Click += BtnZavrit_Click;
            // 
            // splitContainerRight
            // 
            splitContainerRight.Dock = DockStyle.Fill;
            splitContainerRight.Location = new System.Drawing.Point(0, 0);
            splitContainerRight.Name = "splitContainerRight";
            splitContainerRight.Orientation = Orientation.Horizontal;
            // 
            // splitContainerRight.Panel1
            // 
            splitContainerRight.Panel1.Controls.Add(groupBoxPrirazeno);
            // 
            // splitContainerRight.Panel2
            // 
            splitContainerRight.Panel2.Controls.Add(groupBoxNeprirazeno);
            splitContainerRight.Panel2.Controls.Add(panelButtons);
            splitContainerRight.Size = new System.Drawing.Size(836, 650);
            splitContainerRight.SplitterDistance = 300;
            splitContainerRight.TabIndex = 0;
            // 
            // groupBoxPrirazeno
            // 
            groupBoxPrirazeno.Controls.Add(dataGridViewPrirazeno);
            groupBoxPrirazeno.Dock = DockStyle.Fill;
            groupBoxPrirazeno.Location = new System.Drawing.Point(0, 0);
            groupBoxPrirazeno.Name = "groupBoxPrirazeno";
            groupBoxPrirazeno.Size = new System.Drawing.Size(836, 300);
            groupBoxPrirazeno.TabIndex = 0;
            groupBoxPrirazeno.TabStop = false;
            groupBoxPrirazeno.Text = "Zařízení s přiřazenou hodnotou";
            // 
            // dataGridViewPrirazeno
            // 
            dataGridViewPrirazeno.AllowUserToAddRows = false;
            dataGridViewPrirazeno.AllowUserToDeleteRows = false;
            dataGridViewPrirazeno.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewPrirazeno.Dock = DockStyle.Fill;
            dataGridViewPrirazeno.Location = new System.Drawing.Point(3, 25);
            dataGridViewPrirazeno.Name = "dataGridViewPrirazeno";
            dataGridViewPrirazeno.ReadOnly = true;
            dataGridViewPrirazeno.RowHeadersWidth = 51;
            dataGridViewPrirazeno.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewPrirazeno.Size = new System.Drawing.Size(830, 272);
            dataGridViewPrirazeno.TabIndex = 0;
            // 
            // panelButtons
            // 
            panelButtons.Controls.Add(btnOdebrat);
            panelButtons.Controls.Add(btnPriradit);
            panelButtons.Dock = DockStyle.Top;
            panelButtons.Location = new System.Drawing.Point(0, 0);
            panelButtons.Name = "panelButtons";
            panelButtons.Size = new System.Drawing.Size(836, 45);
            panelButtons.TabIndex = 0;
            // 
            // btnOdebrat
            // 
            btnOdebrat.Location = new System.Drawing.Point(260, 4);
            btnOdebrat.Name = "btnOdebrat";
            btnOdebrat.Size = new System.Drawing.Size(240, 36);
            btnOdebrat.TabIndex = 1;
            btnOdebrat.Text = "▼ Odebrat vybrané";
            btnOdebrat.UseVisualStyleBackColor = true;
            btnOdebrat.Click += BtnOdebrat_Click;
            // 
            // btnPriradit
            // 
            btnPriradit.Location = new System.Drawing.Point(10, 4);
            btnPriradit.Name = "btnPriradit";
            btnPriradit.Size = new System.Drawing.Size(240, 36);
            btnPriradit.TabIndex = 0;
            btnPriradit.Text = "▲ Přiřadit vybrané";
            btnPriradit.UseVisualStyleBackColor = true;
            btnPriradit.Click += BtnPriradit_Click;
            // 
            // groupBoxNeprirazeno
            // 
            groupBoxNeprirazeno.Controls.Add(dataGridViewNeprirazeno);
            groupBoxNeprirazeno.Dock = DockStyle.Fill;
            groupBoxNeprirazeno.Location = new System.Drawing.Point(0, 45);
            groupBoxNeprirazeno.Name = "groupBoxNeprirazeno";
            groupBoxNeprirazeno.Size = new System.Drawing.Size(836, 301);
            groupBoxNeprirazeno.TabIndex = 1;
            groupBoxNeprirazeno.TabStop = false;
            groupBoxNeprirazeno.Text = "Nepřiřazená zařízení (chybí hodnota)";
            // 
            // dataGridViewNeprirazeno
            // 
            dataGridViewNeprirazeno.AllowUserToAddRows = false;
            dataGridViewNeprirazeno.AllowUserToDeleteRows = false;
            dataGridViewNeprirazeno.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewNeprirazeno.Dock = DockStyle.Fill;
            dataGridViewNeprirazeno.Location = new System.Drawing.Point(3, 25);
            dataGridViewNeprirazeno.Name = "dataGridViewNeprirazeno";
            dataGridViewNeprirazeno.ReadOnly = true;
            dataGridViewNeprirazeno.RowHeadersWidth = 51;
            dataGridViewNeprirazeno.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewNeprirazeno.Size = new System.Drawing.Size(830, 273);
            dataGridViewNeprirazeno.TabIndex = 0;
            // 
            // FormUmisteni
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1100, 650);
            Controls.Add(splitContainerMain);
            Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 238);
            Margin = new Padding(4);
            MinimumSize = new System.Drawing.Size(900, 600);
            Name = "FormUmisteni";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "Hromadná správa vlastností a umístění";
            Load += FormUmisteni_Load;
            splitContainerMain.Panel1.ResumeLayout(false);
            splitContainerMain.Panel1.PerformLayout();
            splitContainerMain.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainerMain).EndInit();
            splitContainerMain.ResumeLayout(false);
            groupBoxNovy.ResumeLayout(false);
            groupBoxNovy.PerformLayout();
            splitContainerRight.Panel1.ResumeLayout(false);
            splitContainerRight.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainerRight).EndInit();
            splitContainerRight.ResumeLayout(false);
            groupBoxPrirazeno.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewPrirazeno).EndInit();
            panelButtons.ResumeLayout(false);
            groupBoxNeprirazeno.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewNeprirazeno).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainerMain;
        private System.Windows.Forms.Label labelVlastnost;
        private System.Windows.Forms.ComboBox comboBoxVlastnost;
        private System.Windows.Forms.ListBox listBoxHodnoty;
        private System.Windows.Forms.GroupBox groupBoxNovy;
        private System.Windows.Forms.Button btnSloucit;
        private System.Windows.Forms.Button btnPridat;
        private System.Windows.Forms.TextBox txtNovaHodnota;
        private System.Windows.Forms.Label lblStatistika;
        private System.Windows.Forms.Button btnZavrit;
        private System.Windows.Forms.SplitContainer splitContainerRight;
        private System.Windows.Forms.GroupBox groupBoxPrirazeno;
        private System.Windows.Forms.DataGridView dataGridViewPrirazeno;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnOdebrat;
        private System.Windows.Forms.Button btnPriradit;
        private System.Windows.Forms.GroupBox groupBoxNeprirazeno;
        private System.Windows.Forms.DataGridView dataGridViewNeprirazeno;
    }
}
