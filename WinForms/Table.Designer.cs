namespace WinForms
{
    partial class Table
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            dataGridView1 = new DataGridView();
            menuStrip1 = new MenuStrip();
            souborToolStripMenuItem = new ToolStripMenuItem();
            uložitToolStripMenuItem = new ToolStripMenuItem();
            zavřítToolStripMenuItem = new ToolStripMenuItem();
            upravyToolStripMenuItem = new ToolStripMenuItem();
            pridatToolStripMenuItem = new ToolStripMenuItem();
            smazatToolStripMenuItem = new ToolStripMenuItem();
            vypoctyToolStripMenuItem = new ToolStripMenuItem();
            proudToolStripMenuItem = new ToolStripMenuItem();
            prurezToolStripMenuItem = new ToolStripMenuItem();
            zobrazeniToolStripMenuItem = new ToolStripMenuItem();
            vsechnySloupceToolStripMenuItem = new ToolStripMenuItem();
            rozvadecSloupceToolStripMenuItem = new ToolStripMenuItem();
            datoveSloupceToolStripMenuItem = new ToolStripMenuItem();
            panelFilters = new FlowLayoutPanel();
            lblFiltersHeader = new Label();
            lblPid = new Label();
            comboBox4Pid = new ComboBox();
            lblEtapa = new Label();
            comboBox2 = new ComboBox();
            lblRozvadec = new Label();
            comboBox3 = new ComboBox();
            lblPatro = new Label();
            comboBox1 = new ComboBox();
            splitContainer1 = new SplitContainer();
            propertyGrid1 = new PropertyGrid();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            menuStrip1.SuspendLayout();
            panelFilters.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.Location = new Point(0, 0);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.Size = new Size(896, 677);
            dataGridView1.TabIndex = 0;
            dataGridView1.CellContentClick += DataGridView1_CellContentClick;
            dataGridView1.CellFormatting += DataGridView1_CellFormatting;
            dataGridView1.CellMouseUp += DataGridView1_CellMouseUp;
            dataGridView1.CurrentCellChanged += DataGridView1_CurrentCellChanged;
            dataGridView1.RowsAdded += DataGridView1_RowsAdded;
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new Size(20, 20);
            menuStrip1.Items.AddRange(new ToolStripItem[] { souborToolStripMenuItem, upravyToolStripMenuItem, vypoctyToolStripMenuItem, zobrazeniToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(1300, 24);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // souborToolStripMenuItem
            // 
            souborToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { uložitToolStripMenuItem, zavřítToolStripMenuItem });
            souborToolStripMenuItem.Name = "souborToolStripMenuItem";
            souborToolStripMenuItem.Size = new Size(57, 20);
            souborToolStripMenuItem.Text = "Soubor";
            // 
            // uložitToolStripMenuItem
            // 
            uložitToolStripMenuItem.Name = "uložitToolStripMenuItem";
            uložitToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.S;
            uložitToolStripMenuItem.Size = new Size(180, 22);
            uložitToolStripMenuItem.Text = "Uložit (Save)";
            uložitToolStripMenuItem.Click += Button2_Click;
            // 
            // zavřítToolStripMenuItem
            // 
            zavřítToolStripMenuItem.Name = "zavřítToolStripMenuItem";
            zavřítToolStripMenuItem.Size = new Size(180, 22);
            zavřítToolStripMenuItem.Text = "Zavřít (Konec)";
            zavřítToolStripMenuItem.Click += Button1_Click;
            // 
            // upravyToolStripMenuItem
            // 
            upravyToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { pridatToolStripMenuItem, smazatToolStripMenuItem });
            upravyToolStripMenuItem.Name = "upravyToolStripMenuItem";
            upravyToolStripMenuItem.Size = new Size(56, 20);
            upravyToolStripMenuItem.Text = "Úpravy";
            // 
            // pridatToolStripMenuItem
            // 
            pridatToolStripMenuItem.Name = "pridatToolStripMenuItem";
            pridatToolStripMenuItem.Size = new Size(181, 22);
            pridatToolStripMenuItem.Text = "Přidat kopii";
            pridatToolStripMenuItem.Click += BtnAdd_Click;
            // 
            // smazatToolStripMenuItem
            // 
            smazatToolStripMenuItem.Name = "smazatToolStripMenuItem";
            smazatToolStripMenuItem.ShortcutKeys = Keys.Delete;
            smazatToolStripMenuItem.Size = new Size(181, 22);
            smazatToolStripMenuItem.Text = "Smazat vybrané";
            smazatToolStripMenuItem.Click += Button7_Click;
            // 
            // vypoctyToolStripMenuItem
            // 
            vypoctyToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { proudToolStripMenuItem, prurezToolStripMenuItem });
            vypoctyToolStripMenuItem.Name = "vypoctyToolStripMenuItem";
            vypoctyToolStripMenuItem.Size = new Size(62, 20);
            vypoctyToolStripMenuItem.Text = "Výpočty";
            // 
            // proudToolStripMenuItem
            // 
            proudToolStripMenuItem.Name = "proudToolStripMenuItem";
            proudToolStripMenuItem.Size = new Size(107, 22);
            proudToolStripMenuItem.Text = "Proud";
            proudToolStripMenuItem.Click += Button3_Click;
            // 
            // prurezToolStripMenuItem
            // 
            prurezToolStripMenuItem.Name = "prurezToolStripMenuItem";
            prurezToolStripMenuItem.Size = new Size(107, 22);
            prurezToolStripMenuItem.Text = "Průřez";
            prurezToolStripMenuItem.Click += Button4_Click;
            // 
            // zobrazeniToolStripMenuItem
            // 
            zobrazeniToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { vsechnySloupceToolStripMenuItem, rozvadecSloupceToolStripMenuItem, datoveSloupceToolStripMenuItem });
            zobrazeniToolStripMenuItem.Name = "zobrazeniToolStripMenuItem";
            zobrazeniToolStripMenuItem.Size = new Size(71, 20);
            zobrazeniToolStripMenuItem.Text = "Zobrazení";
            // 
            // vsechnySloupceToolStripMenuItem
            // 
            vsechnySloupceToolStripMenuItem.Name = "vsechnySloupceToolStripMenuItem";
            vsechnySloupceToolStripMenuItem.Size = new Size(172, 22);
            vsechnySloupceToolStripMenuItem.Text = "Všechny sloupce";
            vsechnySloupceToolStripMenuItem.Click += Button6_Click;
            // 
            // rozvadecSloupceToolStripMenuItem
            // 
            rozvadecSloupceToolStripMenuItem.Name = "rozvadecSloupceToolStripMenuItem";
            rozvadecSloupceToolStripMenuItem.Size = new Size(172, 22);
            rozvadecSloupceToolStripMenuItem.Text = "Sloupce rozvaděče";
            rozvadecSloupceToolStripMenuItem.Click += Button5_Click;
            // 
            // datoveSloupceToolStripMenuItem
            // 
            datoveSloupceToolStripMenuItem.Name = "datoveSloupceToolStripMenuItem";
            datoveSloupceToolStripMenuItem.Size = new Size(172, 22);
            datoveSloupceToolStripMenuItem.Text = "Datové sloupce";
            datoveSloupceToolStripMenuItem.Click += Button8_Click;
            // 
            // panelFilters
            // 
            panelFilters.Controls.Add(lblFiltersHeader);
            panelFilters.Controls.Add(lblPid);
            panelFilters.Controls.Add(comboBox4Pid);
            panelFilters.Controls.Add(lblEtapa);
            panelFilters.Controls.Add(comboBox2);
            panelFilters.Controls.Add(lblRozvadec);
            panelFilters.Controls.Add(comboBox3);
            panelFilters.Controls.Add(lblPatro);
            panelFilters.Controls.Add(comboBox1);
            panelFilters.Dock = DockStyle.Top;
            panelFilters.Location = new Point(0, 24);
            panelFilters.Name = "panelFilters";
            panelFilters.Padding = new Padding(10, 5, 10, 5);
            panelFilters.Size = new Size(1300, 45);
            panelFilters.TabIndex = 2;
            // 
            // lblFiltersHeader
            // 
            lblFiltersHeader.AutoSize = true;
            lblFiltersHeader.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            lblFiltersHeader.Location = new Point(10, 10);
            lblFiltersHeader.Margin = new Padding(0, 5, 15, 0);
            lblFiltersHeader.Name = "lblFiltersHeader";
            lblFiltersHeader.Size = new Size(54, 21);
            lblFiltersHeader.TabIndex = 0;
            lblFiltersHeader.Text = "Filtry:";
            // 
            // lblPid
            // 
            lblPid.AutoSize = true;
            lblPid.Location = new Point(79, 10);
            lblPid.Margin = new Padding(0, 5, 5, 0);
            lblPid.Name = "lblPid";
            lblPid.Size = new Size(37, 21);
            lblPid.TabIndex = 1;
            lblPid.Text = "PID:";
            // 
            // comboBox4Pid
            // 
            comboBox4Pid.FormattingEnabled = true;
            comboBox4Pid.Location = new Point(121, 5);
            comboBox4Pid.Margin = new Padding(0, 0, 15, 0);
            comboBox4Pid.Name = "comboBox4Pid";
            comboBox4Pid.Size = new Size(120, 29);
            comboBox4Pid.TabIndex = 3;
            comboBox4Pid.SelectedIndexChanged += ComboBox4Pid_SelectedIndexChanged;
            comboBox4Pid.MouseClick += ComboBox4Pid_MouseClick;
            // 
            // lblEtapa
            // 
            lblEtapa.AutoSize = true;
            lblEtapa.Location = new Point(256, 10);
            lblEtapa.Margin = new Padding(0, 5, 5, 0);
            lblEtapa.Name = "lblEtapa";
            lblEtapa.Size = new Size(51, 21);
            lblEtapa.TabIndex = 4;
            lblEtapa.Text = "Etapa:";
            // 
            // comboBox2
            // 
            comboBox2.FormattingEnabled = true;
            comboBox2.Location = new Point(312, 5);
            comboBox2.Margin = new Padding(0, 0, 15, 0);
            comboBox2.Name = "comboBox2";
            comboBox2.Size = new Size(120, 29);
            comboBox2.TabIndex = 4;
            comboBox2.SelectedIndexChanged += ComboBox2_SelectedIndexChanged;
            comboBox2.MouseClick += ComboBox2_MouseClick;
            // 
            // lblRozvadec
            // 
            lblRozvadec.AutoSize = true;
            lblRozvadec.Location = new Point(447, 10);
            lblRozvadec.Margin = new Padding(0, 5, 5, 0);
            lblRozvadec.Name = "lblRozvadec";
            lblRozvadec.Size = new Size(79, 21);
            lblRozvadec.TabIndex = 5;
            lblRozvadec.Text = "Rozvaděč:";
            // 
            // comboBox3
            // 
            comboBox3.FormattingEnabled = true;
            comboBox3.Location = new Point(531, 5);
            comboBox3.Margin = new Padding(0, 0, 15, 0);
            comboBox3.Name = "comboBox3";
            comboBox3.Size = new Size(120, 29);
            comboBox3.TabIndex = 5;
            comboBox3.SelectedIndexChanged += ComboBox3_SelectedIndexChanged;
            comboBox3.MouseClick += ComboBox3_MouseClick;
            // 
            // lblPatro
            // 
            lblPatro.AutoSize = true;
            lblPatro.Location = new Point(666, 10);
            lblPatro.Margin = new Padding(0, 5, 5, 0);
            lblPatro.Name = "lblPatro";
            lblPatro.Size = new Size(49, 21);
            lblPatro.TabIndex = 6;
            lblPatro.Text = "Patro:";
            // 
            // comboBox1
            // 
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(723, 8);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(120, 29);
            comboBox1.TabIndex = 6;
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;
            comboBox1.MouseClick += ComboBox1_MouseClick;
            // 
            // splitContainer1
            // 
            splitContainer1.BorderStyle = BorderStyle.Fixed3D;
            splitContainer1.Dock = DockStyle.Fill;
            splitContainer1.Location = new Point(0, 69);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(dataGridView1);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(propertyGrid1);
            splitContainer1.Size = new Size(1300, 681);
            splitContainer1.SplitterDistance = 900;
            splitContainer1.TabIndex = 0;
            // 
            // propertyGrid1
            // 
            propertyGrid1.BackColor = SystemColors.Control;
            propertyGrid1.Dock = DockStyle.Fill;
            propertyGrid1.Location = new Point(0, 0);
            propertyGrid1.Name = "propertyGrid1";
            propertyGrid1.PropertySort = PropertySort.Categorized;
            propertyGrid1.Size = new Size(392, 677);
            propertyGrid1.TabIndex = 0;
            // 
            // Table
            // 
            AutoScaleDimensions = new SizeF(9F, 21F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1300, 750);
            Controls.Add(splitContainer1);
            Controls.Add(panelFilters);
            Controls.Add(menuStrip1);
            Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 238);
            MainMenuStrip = menuStrip1;
            Margin = new Padding(4);
            Name = "Table";
            Text = "Seznam zařízení";
            FormClosing += Table_FormClosing;
            Load += Table_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            panelFilters.ResumeLayout(false);
            panelFilters.PerformLayout();
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public DataGridView dataGridView1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem souborToolStripMenuItem;
        private ToolStripMenuItem uložitToolStripMenuItem;
        private ToolStripMenuItem zavřítToolStripMenuItem;
        private ToolStripMenuItem upravyToolStripMenuItem;
        private ToolStripMenuItem pridatToolStripMenuItem;
        private ToolStripMenuItem smazatToolStripMenuItem;
        private ToolStripMenuItem vypoctyToolStripMenuItem;
        private ToolStripMenuItem proudToolStripMenuItem;
        private ToolStripMenuItem prurezToolStripMenuItem;
        private ToolStripMenuItem zobrazeniToolStripMenuItem;
        private ToolStripMenuItem vsechnySloupceToolStripMenuItem;
        private ToolStripMenuItem rozvadecSloupceToolStripMenuItem;
        private ToolStripMenuItem datoveSloupceToolStripMenuItem;
        private FlowLayoutPanel panelFilters;
        private Label lblFiltersHeader;
        private Label lblPid;
        private Label lblEtapa;
        private Label lblRozvadec;
        private Label lblPatro;
        private ComboBox comboBox1;
        private ComboBox comboBox2;
        private ComboBox comboBox3;
        private ComboBox comboBox4Pid;
        public SplitContainer splitContainer1;
        public PropertyGrid propertyGrid1;
    }
}