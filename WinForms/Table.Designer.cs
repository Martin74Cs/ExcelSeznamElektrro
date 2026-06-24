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
            openToolStripMenuItem = new ToolStripMenuItem();
            upravyToolStripMenuItem = new ToolStripMenuItem();
            pridatToolStripMenuItem = new ToolStripMenuItem();
            smazatToolStripMenuItem = new ToolStripMenuItem();
            kabelyToolStripMenuItem1 = new ToolStripMenuItem();
            spravaKabeluToolStripMenuItem = new ToolStripMenuItem();
            rozvadeceToolStripMenuItem = new ToolStripMenuItem();
            prirazeniKRozvadecumToolStripMenuItem = new ToolStripMenuItem();
            prehledToolStripMenuItem = new ToolStripMenuItem();
            vypoctyToolStripMenuItem = new ToolStripMenuItem();
            proudToolStripMenuItem = new ToolStripMenuItem();
            prurezToolStripMenuItem = new ToolStripMenuItem();
            zobrazeniToolStripMenuItem = new ToolStripMenuItem();
            vsechnySloupceToolStripMenuItem = new ToolStripMenuItem();
            rozvadecSloupceToolStripMenuItem = new ToolStripMenuItem();
            datoveSloupceToolStripMenuItem = new ToolStripMenuItem();
            filtToolStripMenuItem = new ToolStripMenuItem();
            smazatToolStripMenuItem1 = new ToolStripMenuItem();
            bezKWToolStripMenuItem = new ToolStripMenuItem();
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
            lblIsExist = new Label();
            comboBoxIsExist = new ComboBox();
            lblIsExistElektro = new Label();
            comboBoxIsExistElektro = new ComboBox();
            lblSearch = new Label();
            textBoxSearch = new TextBox();
            splitContainer1 = new SplitContainer();
            propertyGrid1 = new PropertyGrid();
            flowLayoutPanelButtons = new FlowLayoutPanel();
            BtnAdd = new Button();
            button7 = new Button();
            button3 = new Button();
            Button4 = new Button();
            button5 = new Button();
            button6 = new Button();
            button8 = new Button();
            Button2 = new Button();
            Button1 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            menuStrip1.SuspendLayout();
            panelFilters.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            flowLayoutPanelButtons.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.Location = new Point(0, 0);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.Size = new Size(794, 355);
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
            menuStrip1.Items.AddRange(new ToolStripItem[] { souborToolStripMenuItem, upravyToolStripMenuItem, kabelyToolStripMenuItem1, rozvadeceToolStripMenuItem, vypoctyToolStripMenuItem, zobrazeniToolStripMenuItem, smazatToolStripMenuItem1 });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(1154, 24);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // souborToolStripMenuItem
            // 
            souborToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { uložitToolStripMenuItem, zavřítToolStripMenuItem, openToolStripMenuItem });
            souborToolStripMenuItem.Name = "souborToolStripMenuItem";
            souborToolStripMenuItem.Size = new Size(57, 20);
            souborToolStripMenuItem.Text = "Soubor";
            // 
            // uložitToolStripMenuItem
            // 
            uložitToolStripMenuItem.Name = "uložitToolStripMenuItem";
            uložitToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.S;
            uložitToolStripMenuItem.Size = new Size(179, 22);
            uložitToolStripMenuItem.Text = "Uložit (Save)";
            uložitToolStripMenuItem.Click += Button2_Click;
            // 
            // zavřítToolStripMenuItem
            // 
            zavřítToolStripMenuItem.Name = "zavřítToolStripMenuItem";
            zavřítToolStripMenuItem.Size = new Size(179, 22);
            zavřítToolStripMenuItem.Text = "Zavřít (Konec)";
            zavřítToolStripMenuItem.Click += Button1_Click;
            // 
            // openToolStripMenuItem
            // 
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.Size = new Size(179, 22);
            openToolStripMenuItem.Text = "Open";
            openToolStripMenuItem.Click += openToolStripMenuItem_Click;
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
            pridatToolStripMenuItem.Size = new Size(157, 22);
            pridatToolStripMenuItem.Text = "Přidat kopii";
            pridatToolStripMenuItem.Click += BtnAdd_Click;
            // 
            // smazatToolStripMenuItem
            // 
            smazatToolStripMenuItem.Name = "smazatToolStripMenuItem";
            smazatToolStripMenuItem.Size = new Size(157, 22);
            smazatToolStripMenuItem.Text = "Smazat vybrané";
            smazatToolStripMenuItem.Click += Button7_Click;
            // 
            // kabelyToolStripMenuItem1
            // 
            kabelyToolStripMenuItem1.DropDownItems.AddRange(new ToolStripItem[] { spravaKabeluToolStripMenuItem });
            kabelyToolStripMenuItem1.Name = "kabelyToolStripMenuItem1";
            kabelyToolStripMenuItem1.Size = new Size(54, 20);
            kabelyToolStripMenuItem1.Text = "Kabely";
            // 
            // spravaKabeluToolStripMenuItem
            // 
            spravaKabeluToolStripMenuItem.Name = "spravaKabeluToolStripMenuItem";
            spravaKabeluToolStripMenuItem.Size = new Size(147, 22);
            spravaKabeluToolStripMenuItem.Text = "Správa kabelů";
            spravaKabeluToolStripMenuItem.Click += SpravaKabeluToolStripMenuItem_Click;
            // 
            // rozvadeceToolStripMenuItem
            // 
            rozvadeceToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { prirazeniKRozvadecumToolStripMenuItem, prehledToolStripMenuItem });
            rozvadeceToolStripMenuItem.Name = "rozvadeceToolStripMenuItem";
            rozvadeceToolStripMenuItem.Size = new Size(75, 20);
            rozvadeceToolStripMenuItem.Text = "Rozvaděče";
            // 
            // prirazeniKRozvadecumToolStripMenuItem
            // 
            prirazeniKRozvadecumToolStripMenuItem.Name = "prirazeniKRozvadecumToolStripMenuItem";
            prirazeniKRozvadecumToolStripMenuItem.Size = new Size(196, 22);
            prirazeniKRozvadecumToolStripMenuItem.Text = "Přiřazení k rozvaděčům";
            prirazeniKRozvadecumToolStripMenuItem.Click += PrirazeniKRozvadecumToolStripMenuItem_Click;
            // 
            // prehledToolStripMenuItem
            // 
            prehledToolStripMenuItem.Name = "prehledToolStripMenuItem";
            prehledToolStripMenuItem.Size = new Size(196, 22);
            prehledToolStripMenuItem.Text = "Přehled rozvaděčů";
            prehledToolStripMenuItem.Click += PrehledToolStripMenuItem_Click;
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
            zobrazeniToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { vsechnySloupceToolStripMenuItem, rozvadecSloupceToolStripMenuItem, datoveSloupceToolStripMenuItem, filtToolStripMenuItem });
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
            // filtToolStripMenuItem
            // 
            filtToolStripMenuItem.Name = "filtToolStripMenuItem";
            filtToolStripMenuItem.Size = new Size(172, 22);
            filtToolStripMenuItem.Text = "Filtr";
            filtToolStripMenuItem.Click += FiltToolStripMenuItem_Click;
            // 
            // smazatToolStripMenuItem1
            // 
            smazatToolStripMenuItem1.DropDownItems.AddRange(new ToolStripItem[] { bezKWToolStripMenuItem });
            smazatToolStripMenuItem1.Name = "smazatToolStripMenuItem1";
            smazatToolStripMenuItem1.Size = new Size(57, 20);
            smazatToolStripMenuItem1.Text = "Smazat";
            // 
            // bezKWToolStripMenuItem
            // 
            bezKWToolStripMenuItem.Name = "bezKWToolStripMenuItem";
            bezKWToolStripMenuItem.Size = new Size(112, 22);
            bezKWToolStripMenuItem.Text = "Bez kW";
            bezKWToolStripMenuItem.Click += bezKWToolStripMenuItem_Click;
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
            panelFilters.Controls.Add(lblIsExist);
            panelFilters.Controls.Add(comboBoxIsExist);
            panelFilters.Controls.Add(lblIsExistElektro);
            panelFilters.Controls.Add(comboBoxIsExistElektro);
            panelFilters.Controls.Add(lblSearch);
            panelFilters.Controls.Add(textBoxSearch);
            panelFilters.Dock = DockStyle.Top;
            panelFilters.Location = new Point(0, 24);
            panelFilters.Name = "panelFilters";
            panelFilters.Padding = new Padding(10, 5, 10, 5);
            panelFilters.Size = new Size(1154, 92);
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
            // lblIsExist
            // 
            lblIsExist.AutoSize = true;
            lblIsExist.Location = new Point(846, 10);
            lblIsExist.Margin = new Padding(0, 5, 5, 0);
            lblIsExist.Name = "lblIsExist";
            lblIsExist.Size = new Size(85, 21);
            lblIsExist.TabIndex = 9;
            lblIsExist.Text = "V projektu:";
            // 
            // comboBoxIsExist
            // 
            comboBoxIsExist.FormattingEnabled = true;
            comboBoxIsExist.Location = new Point(936, 5);
            comboBoxIsExist.Margin = new Padding(0, 0, 15, 0);
            comboBoxIsExist.Name = "comboBoxIsExist";
            comboBoxIsExist.Size = new Size(100, 29);
            comboBoxIsExist.TabIndex = 10;
            comboBoxIsExist.SelectedIndexChanged += ComboBoxIsExist_SelectedIndexChanged;
            // 
            // lblIsExistElektro
            // 
            lblIsExistElektro.AutoSize = true;
            lblIsExistElektro.Location = new Point(1051, 10);
            lblIsExistElektro.Margin = new Padding(0, 5, 5, 0);
            lblIsExistElektro.Name = "lblIsExistElektro";
            lblIsExistElektro.Size = new Size(75, 21);
            lblIsExistElektro.TabIndex = 11;
            lblIsExistElektro.Text = "V Elektro:";
            // 
            // comboBoxIsExistElektro
            // 
            comboBoxIsExistElektro.FormattingEnabled = true;
            comboBoxIsExistElektro.Location = new Point(10, 34);
            comboBoxIsExistElektro.Margin = new Padding(0, 0, 15, 0);
            comboBoxIsExistElektro.Name = "comboBoxIsExistElektro";
            comboBoxIsExistElektro.Size = new Size(100, 29);
            comboBoxIsExistElektro.TabIndex = 12;
            comboBoxIsExistElektro.SelectedIndexChanged += ComboBoxIsExistElektro_SelectedIndexChanged;
            // 
            // lblSearch
            // 
            lblSearch.AutoSize = true;
            lblSearch.Location = new Point(125, 39);
            lblSearch.Margin = new Padding(0, 5, 5, 0);
            lblSearch.Name = "lblSearch";
            lblSearch.Size = new Size(74, 21);
            lblSearch.TabIndex = 7;
            lblSearch.Text = "Vyhledat:";
            // 
            // textBoxSearch
            // 
            textBoxSearch.Location = new Point(204, 34);
            textBoxSearch.Margin = new Padding(0, 0, 15, 0);
            textBoxSearch.Name = "textBoxSearch";
            textBoxSearch.Size = new Size(201, 29);
            textBoxSearch.TabIndex = 8;
            textBoxSearch.TextChanged += TextBoxSearch_TextChanged;
            // 
            // splitContainer1
            // 
            splitContainer1.BorderStyle = BorderStyle.Fixed3D;
            splitContainer1.Dock = DockStyle.Fill;
            splitContainer1.Location = new Point(0, 116);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(dataGridView1);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(propertyGrid1);
            splitContainer1.Size = new Size(1154, 359);
            splitContainer1.SplitterDistance = 798;
            splitContainer1.TabIndex = 0;
            // 
            // propertyGrid1
            // 
            propertyGrid1.BackColor = SystemColors.Control;
            propertyGrid1.Dock = DockStyle.Fill;
            propertyGrid1.Location = new Point(0, 0);
            propertyGrid1.Name = "propertyGrid1";
            propertyGrid1.PropertySort = PropertySort.Categorized;
            propertyGrid1.Size = new Size(348, 355);
            propertyGrid1.TabIndex = 0;
            // 
            // flowLayoutPanelButtons
            // 
            flowLayoutPanelButtons.Controls.Add(BtnAdd);
            flowLayoutPanelButtons.Controls.Add(button7);
            flowLayoutPanelButtons.Controls.Add(button3);
            flowLayoutPanelButtons.Controls.Add(Button4);
            flowLayoutPanelButtons.Controls.Add(button5);
            flowLayoutPanelButtons.Controls.Add(button6);
            flowLayoutPanelButtons.Controls.Add(button8);
            flowLayoutPanelButtons.Controls.Add(Button2);
            flowLayoutPanelButtons.Controls.Add(Button1);
            flowLayoutPanelButtons.Dock = DockStyle.Bottom;
            flowLayoutPanelButtons.Location = new Point(0, 475);
            flowLayoutPanelButtons.Name = "flowLayoutPanelButtons";
            flowLayoutPanelButtons.Padding = new Padding(10, 5, 10, 5);
            flowLayoutPanelButtons.Size = new Size(1154, 60);
            flowLayoutPanelButtons.TabIndex = 3;
            // 
            // BtnAdd
            // 
            BtnAdd.Location = new Point(13, 8);
            BtnAdd.Name = "BtnAdd";
            BtnAdd.Size = new Size(130, 36);
            BtnAdd.TabIndex = 0;
            BtnAdd.Text = "Přidat kopii";
            BtnAdd.UseVisualStyleBackColor = true;
            BtnAdd.Click += BtnAdd_Click;
            // 
            // button7
            // 
            button7.Location = new Point(149, 8);
            button7.Name = "button7";
            button7.Size = new Size(100, 36);
            button7.TabIndex = 1;
            button7.Text = "Smazat";
            button7.UseVisualStyleBackColor = true;
            button7.Click += Button7_Click;
            // 
            // button3
            // 
            button3.Location = new Point(255, 8);
            button3.Name = "button3";
            button3.Size = new Size(100, 36);
            button3.TabIndex = 2;
            button3.Text = "Proud";
            button3.UseVisualStyleBackColor = true;
            button3.Click += Button3_Click;
            // 
            // Button4
            // 
            Button4.Location = new Point(361, 8);
            Button4.Name = "Button4";
            Button4.Size = new Size(100, 36);
            Button4.TabIndex = 3;
            Button4.Text = "Průřez";
            Button4.UseVisualStyleBackColor = true;
            Button4.Click += Button4_Click;
            // 
            // button5
            // 
            button5.Location = new Point(467, 8);
            button5.Name = "button5";
            button5.Size = new Size(170, 36);
            button5.TabIndex = 4;
            button5.Text = "Sloupce rozvaděče";
            button5.UseVisualStyleBackColor = true;
            button5.Click += Button5_Click;
            // 
            // button6
            // 
            button6.Location = new Point(643, 8);
            button6.Name = "button6";
            button6.Size = new Size(150, 36);
            button6.TabIndex = 5;
            button6.Text = "Všechny sloupce";
            button6.UseVisualStyleBackColor = true;
            button6.Click += Button6_Click;
            // 
            // button8
            // 
            button8.Location = new Point(799, 8);
            button8.Name = "button8";
            button8.Size = new Size(150, 36);
            button8.TabIndex = 6;
            button8.Text = "Datové sloupce";
            button8.UseVisualStyleBackColor = true;
            button8.Click += Button8_Click;
            // 
            // Button2
            // 
            Button2.Location = new Point(955, 8);
            Button2.Name = "Button2";
            Button2.Size = new Size(100, 36);
            Button2.TabIndex = 7;
            Button2.Text = "Uložit";
            Button2.UseVisualStyleBackColor = true;
            Button2.Click += Button2_Click;
            // 
            // Button1
            // 
            Button1.Anchor = AnchorStyles.Right;
            Button1.Location = new Point(13, 50);
            Button1.Name = "Button1";
            Button1.Size = new Size(100, 36);
            Button1.TabIndex = 8;
            Button1.Text = "Zavřít";
            Button1.UseVisualStyleBackColor = true;
            Button1.Click += Button1_Click;
            // 
            // Table
            // 
            AutoScaleDimensions = new SizeF(9F, 21F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1154, 535);
            Controls.Add(splitContainer1);
            Controls.Add(flowLayoutPanelButtons);
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
            flowLayoutPanelButtons.ResumeLayout(false);
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
        private ToolStripMenuItem filtToolStripMenuItem;
        private ToolStripMenuItem kabelyToolStripMenuItem1;
        private ToolStripMenuItem spravaKabeluToolStripMenuItem;
        private ToolStripMenuItem rozvadeceToolStripMenuItem;
        private ToolStripMenuItem prirazeniKRozvadecumToolStripMenuItem;
        private ToolStripMenuItem prehledToolStripMenuItem;
        private Button Button1;
        private Button Button2;
        private Button button3;
        private Button Button4;
        private Button button5;
        private Button button6;
        private Button button7;
        private Button button8;
        private Button BtnAdd;
        private FlowLayoutPanel flowLayoutPanelButtons;
        private Label lblSearch;
        private TextBox textBoxSearch;
        private Label lblIsExist;
        private ComboBox comboBoxIsExist;
        private Label lblIsExistElektro;
        private ComboBox comboBoxIsExistElektro;
        private ToolStripMenuItem openToolStripMenuItem;
        private ToolStripMenuItem smazatToolStripMenuItem1;
        private ToolStripMenuItem bezKWToolStripMenuItem;
    }
}