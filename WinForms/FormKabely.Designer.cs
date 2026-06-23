namespace WinForms
{
    partial class FormKabely
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
            groupBoxFiltry = new GroupBox();
            FiltrText = new TextBox();
            CheckBoxIsExist = new CheckBox();
            lblFilterExistElektro = new Label();
            comboBoxFilterExistElektro = new ComboBox();
            lblFilterEtapa = new Label();
            comboBoxFilterEtapa = new ComboBox();
            label1 = new Label();
            comboBoxZarizeni = new ComboBox();
            groupBoxInfoZarizeni = new GroupBox();
            lblInfoPopis = new Label();
            txtInfoPopis = new TextBox();
            lblInfoPrikon = new Label();
            txtInfoPrikon = new TextBox();
            lblInfoPrikonStroj = new Label();
            txtInfoPrikonStroj = new TextBox();
            lblInfoProud = new Label();
            txtInfoProud = new TextBox();
            lblInfoNapeti = new Label();
            txtInfoNapeti = new TextBox();
            lblInfoMenic = new Label();
            txtInfoMenic = new TextBox();
            lblInfoBalena = new Label();
            txtInfoBalena = new TextBox();
            lblInfoDruh = new Label();
            txtInfoDruh = new TextBox();
            dataGridViewKabely = new DataGridView();
            groupBoxPridat = new GroupBox();
            lblPopis = new Label();
            txtPopis = new TextBox();
            label6 = new Label();
            txtDelka = new TextBox();
            label5 = new Label();
            txtPrurez = new TextBox();
            label4 = new Label();
            txtPocetZil = new TextBox();
            label3 = new Label();
            txtTyp = new TextBox();
            label2 = new Label();
            txtOznaceni = new TextBox();
            comboBoxZnacka = new ComboBox();
            lblZnacka = new Label();
            btnPridat = new Button();
            btnStorno = new Button();
            groupBoxRychlePridat = new GroupBox();
            label7 = new Label();
            txtPrefixPower = new TextBox();
            button1 = new Button();
            lblSekcePrefixy = new Label();
            lblPrefixPTC = new Label();
            txtPrefixPTC = new TextBox();
            lblPrefixOvladani = new Label();
            txtPrefixOvladani = new TextBox();
            lblPrefixUTP = new Label();
            txtPrefixUTP = new TextBox();
            lblPrefixBinarni = new Label();
            txtPrefixBinarni = new TextBox();
            lblPrefixBlokovani = new Label();
            txtPrefixBlokovani = new TextBox();
            btnRychlyPTC = new Button();
            btnRychlyOvladani5 = new Button();
            btnRychlyOvladani7 = new Button();
            btnRychlyOvladani12 = new Button();
            btnRychlyUTP = new Button();
            btnRychlyBinarni = new Button();
            btnRychlyBlokovani = new Button();
            btnSmazat = new Button();
            lblStatistika = new Label();
            btnZavrit = new Button();
            groupBoxFiltry.SuspendLayout();
            groupBoxInfoZarizeni.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewKabely).BeginInit();
            groupBoxPridat.SuspendLayout();
            groupBoxRychlePridat.SuspendLayout();
            SuspendLayout();
            // 
            // groupBoxFiltry
            // 
            groupBoxFiltry.Controls.Add(FiltrText);
            groupBoxFiltry.Controls.Add(CheckBoxIsExist);
            groupBoxFiltry.Controls.Add(lblFilterExistElektro);
            groupBoxFiltry.Controls.Add(comboBoxFilterExistElektro);
            groupBoxFiltry.Controls.Add(lblFilterEtapa);
            groupBoxFiltry.Controls.Add(comboBoxFilterEtapa);
            groupBoxFiltry.Controls.Add(label1);
            groupBoxFiltry.Controls.Add(comboBoxZarizeni);
            groupBoxFiltry.Location = new Point(12, 12);
            groupBoxFiltry.Name = "groupBoxFiltry";
            groupBoxFiltry.Size = new Size(1136, 90);
            groupBoxFiltry.TabIndex = 0;
            groupBoxFiltry.TabStop = false;
            groupBoxFiltry.Text = "Filtry a výběr zařízení";
            // 
            // FiltrText
            // 
            FiltrText.Location = new Point(948, 47);
            FiltrText.Name = "FiltrText";
            FiltrText.Size = new Size(164, 34);
            FiltrText.TabIndex = 9;
            FiltrText.TextChanged += textBox1_TextChanged;
            // 
            // CheckBoxIsExist
            // 
            CheckBoxIsExist.AutoSize = true;
            CheckBoxIsExist.Location = new Point(12, 47);
            CheckBoxIsExist.Name = "CheckBoxIsExist";
            CheckBoxIsExist.Size = new Size(86, 32);
            CheckBoxIsExist.TabIndex = 8;
            CheckBoxIsExist.Text = "IsExist";
            CheckBoxIsExist.UseVisualStyleBackColor = true;
            CheckBoxIsExist.CheckedChanged += CheckBoxIsExiste_CheckedChanged;
            // 
            // lblFilterExistElektro
            // 
            lblFilterExistElektro.AutoSize = true;
            lblFilterExistElektro.Location = new Point(165, 23);
            lblFilterExistElektro.Name = "lblFilterExistElektro";
            lblFilterExistElektro.Size = new Size(121, 28);
            lblFilterExistElektro.TabIndex = 2;
            lblFilterExistElektro.Text = "Elektro blok:";
            // 
            // comboBoxFilterExistElektro
            // 
            comboBoxFilterExistElektro.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxFilterExistElektro.FormattingEnabled = true;
            comboBoxFilterExistElektro.Location = new Point(165, 47);
            comboBoxFilterExistElektro.Name = "comboBoxFilterExistElektro";
            comboBoxFilterExistElektro.Size = new Size(140, 36);
            comboBoxFilterExistElektro.TabIndex = 3;
            comboBoxFilterExistElektro.SelectedIndexChanged += ComboBoxFilter_SelectedIndexChanged;
            // 
            // lblFilterEtapa
            // 
            lblFilterEtapa.AutoSize = true;
            lblFilterEtapa.Location = new Point(318, 23);
            lblFilterEtapa.Name = "lblFilterEtapa";
            lblFilterEtapa.Size = new Size(136, 28);
            lblFilterEtapa.TabIndex = 4;
            lblFilterEtapa.Text = "Fáze výstavby:";
            // 
            // comboBoxFilterEtapa
            // 
            comboBoxFilterEtapa.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxFilterEtapa.FormattingEnabled = true;
            comboBoxFilterEtapa.Location = new Point(318, 47);
            comboBoxFilterEtapa.Name = "comboBoxFilterEtapa";
            comboBoxFilterEtapa.Size = new Size(140, 36);
            comboBoxFilterEtapa.TabIndex = 5;
            comboBoxFilterEtapa.SelectedIndexChanged += ComboBoxFilter_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(471, 23);
            label1.Name = "label1";
            label1.Size = new Size(84, 28);
            label1.TabIndex = 6;
            label1.Text = "Zařízení:";
            // 
            // comboBoxZarizeni
            // 
            comboBoxZarizeni.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxZarizeni.FormattingEnabled = true;
            comboBoxZarizeni.Location = new Point(471, 47);
            comboBoxZarizeni.Name = "comboBoxZarizeni";
            comboBoxZarizeni.Size = new Size(461, 36);
            comboBoxZarizeni.TabIndex = 7;
            comboBoxZarizeni.SelectedIndexChanged += ComboBoxZarizeni_SelectedIndexChanged;
            // 
            // groupBoxInfoZarizeni
            // 
            groupBoxInfoZarizeni.Controls.Add(lblInfoPopis);
            groupBoxInfoZarizeni.Controls.Add(txtInfoPopis);
            groupBoxInfoZarizeni.Controls.Add(lblInfoPrikon);
            groupBoxInfoZarizeni.Controls.Add(txtInfoPrikon);
            groupBoxInfoZarizeni.Controls.Add(lblInfoPrikonStroj);
            groupBoxInfoZarizeni.Controls.Add(txtInfoPrikonStroj);
            groupBoxInfoZarizeni.Controls.Add(lblInfoProud);
            groupBoxInfoZarizeni.Controls.Add(txtInfoProud);
            groupBoxInfoZarizeni.Controls.Add(lblInfoNapeti);
            groupBoxInfoZarizeni.Controls.Add(txtInfoNapeti);
            groupBoxInfoZarizeni.Controls.Add(lblInfoMenic);
            groupBoxInfoZarizeni.Controls.Add(txtInfoMenic);
            groupBoxInfoZarizeni.Controls.Add(lblInfoBalena);
            groupBoxInfoZarizeni.Controls.Add(txtInfoBalena);
            groupBoxInfoZarizeni.Controls.Add(lblInfoDruh);
            groupBoxInfoZarizeni.Controls.Add(txtInfoDruh);
            groupBoxInfoZarizeni.Location = new Point(12, 110);
            groupBoxInfoZarizeni.Name = "groupBoxInfoZarizeni";
            groupBoxInfoZarizeni.Size = new Size(1136, 95);
            groupBoxInfoZarizeni.TabIndex = 1;
            groupBoxInfoZarizeni.TabStop = false;
            groupBoxInfoZarizeni.Text = "Základní informace o zařízení";
            // 
            // lblInfoPopis
            // 
            lblInfoPopis.AutoSize = true;
            lblInfoPopis.Location = new Point(12, 23);
            lblInfoPopis.Name = "lblInfoPopis";
            lblInfoPopis.Size = new Size(134, 28);
            lblInfoPopis.TabIndex = 0;
            lblInfoPopis.Text = "Název (Popis):";
            // 
            // txtInfoPopis
            // 
            txtInfoPopis.Location = new Point(12, 46);
            txtInfoPopis.Name = "txtInfoPopis";
            txtInfoPopis.ReadOnly = true;
            txtInfoPopis.Size = new Size(311, 34);
            txtInfoPopis.TabIndex = 1;
            // 
            // lblInfoPrikon
            // 
            lblInfoPrikon.AutoSize = true;
            lblInfoPrikon.Location = new Point(324, 23);
            lblInfoPrikon.Name = "lblInfoPrikon";
            lblInfoPrikon.Size = new Size(123, 28);
            lblInfoPrikon.TabIndex = 2;
            lblInfoPrikon.Text = "kW (elektro):";
            // 
            // txtInfoPrikon
            // 
            txtInfoPrikon.Location = new Point(329, 46);
            txtInfoPrikon.Name = "txtInfoPrikon";
            txtInfoPrikon.ReadOnly = true;
            txtInfoPrikon.Size = new Size(90, 34);
            txtInfoPrikon.TabIndex = 3;
            // 
            // lblInfoPrikonStroj
            // 
            lblInfoPrikonStroj.AutoSize = true;
            lblInfoPrikonStroj.Location = new Point(441, 23);
            lblInfoPrikonStroj.Name = "lblInfoPrikonStroj";
            lblInfoPrikonStroj.Size = new Size(101, 28);
            lblInfoPrikonStroj.TabIndex = 4;
            lblInfoPrikonStroj.Text = "kW (stroj):";
            // 
            // txtInfoPrikonStroj
            // 
            txtInfoPrikonStroj.Location = new Point(443, 46);
            txtInfoPrikonStroj.Name = "txtInfoPrikonStroj";
            txtInfoPrikonStroj.ReadOnly = true;
            txtInfoPrikonStroj.Size = new Size(90, 34);
            txtInfoPrikonStroj.TabIndex = 5;
            // 
            // lblInfoProud
            // 
            lblInfoProud.AutoSize = true;
            lblInfoProud.Location = new Point(542, 23);
            lblInfoProud.Name = "lblInfoProud";
            lblInfoProud.Size = new Size(99, 28);
            lblInfoProud.TabIndex = 6;
            lblInfoProud.Text = "Proud [A]:";
            // 
            // txtInfoProud
            // 
            txtInfoProud.Location = new Point(549, 46);
            txtInfoProud.Name = "txtInfoProud";
            txtInfoProud.ReadOnly = true;
            txtInfoProud.Size = new Size(80, 34);
            txtInfoProud.TabIndex = 7;
            // 
            // lblInfoNapeti
            // 
            lblInfoNapeti.AutoSize = true;
            lblInfoNapeti.Location = new Point(637, 23);
            lblInfoNapeti.Name = "lblInfoNapeti";
            lblInfoNapeti.Size = new Size(104, 28);
            lblInfoNapeti.TabIndex = 8;
            lblInfoNapeti.Text = "Napětí [V]:";
            // 
            // txtInfoNapeti
            // 
            txtInfoNapeti.Location = new Point(645, 46);
            txtInfoNapeti.Name = "txtInfoNapeti";
            txtInfoNapeti.ReadOnly = true;
            txtInfoNapeti.Size = new Size(80, 34);
            txtInfoNapeti.TabIndex = 9;
            // 
            // lblInfoMenic
            // 
            lblInfoMenic.AutoSize = true;
            lblInfoMenic.Location = new Point(739, 23);
            lblInfoMenic.Name = "lblInfoMenic";
            lblInfoMenic.Size = new Size(69, 28);
            lblInfoMenic.TabIndex = 10;
            lblInfoMenic.Text = "Měnič:";
            // 
            // txtInfoMenic
            // 
            txtInfoMenic.Location = new Point(738, 46);
            txtInfoMenic.Name = "txtInfoMenic";
            txtInfoMenic.ReadOnly = true;
            txtInfoMenic.Size = new Size(81, 34);
            txtInfoMenic.TabIndex = 11;
            // 
            // lblInfoBalena
            // 
            lblInfoBalena.AutoSize = true;
            lblInfoBalena.Location = new Point(825, 23);
            lblInfoBalena.Name = "lblInfoBalena";
            lblInfoBalena.Size = new Size(109, 28);
            lblInfoBalena.TabIndex = 12;
            lblInfoBalena.Text = "Balená jed.:";
            // 
            // txtInfoBalena
            // 
            txtInfoBalena.Location = new Point(834, 46);
            txtInfoBalena.Name = "txtInfoBalena";
            txtInfoBalena.ReadOnly = true;
            txtInfoBalena.Size = new Size(98, 34);
            txtInfoBalena.TabIndex = 13;
            // 
            // lblInfoDruh
            // 
            lblInfoDruh.AutoSize = true;
            lblInfoDruh.Location = new Point(942, 23);
            lblInfoDruh.Name = "lblInfoDruh";
            lblInfoDruh.Size = new Size(130, 28);
            lblInfoDruh.TabIndex = 14;
            lblInfoDruh.Text = "Druh zařízení:";
            // 
            // txtInfoDruh
            // 
            txtInfoDruh.Location = new Point(942, 46);
            txtInfoDruh.Name = "txtInfoDruh";
            txtInfoDruh.ReadOnly = true;
            txtInfoDruh.Size = new Size(180, 34);
            txtInfoDruh.TabIndex = 15;
            // 
            // dataGridViewKabely
            // 
            dataGridViewKabely.AllowUserToAddRows = false;
            dataGridViewKabely.AllowUserToDeleteRows = false;
            dataGridViewKabely.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewKabely.Location = new Point(12, 215);
            dataGridViewKabely.MultiSelect = false;
            dataGridViewKabely.Name = "dataGridViewKabely";
            dataGridViewKabely.ReadOnly = true;
            dataGridViewKabely.RowHeadersWidth = 51;
            dataGridViewKabely.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewKabely.Size = new Size(796, 520);
            dataGridViewKabely.TabIndex = 2;
            dataGridViewKabely.CellContentClick += dataGridViewKabely_CellContentClick;
            dataGridViewKabely.CellDoubleClick += DataGridViewKabely_CellDoubleClick;
            // 
            // groupBoxPridat
            // 
            groupBoxPridat.Controls.Add(lblPopis);
            groupBoxPridat.Controls.Add(txtPopis);
            groupBoxPridat.Controls.Add(label6);
            groupBoxPridat.Controls.Add(txtDelka);
            groupBoxPridat.Controls.Add(label5);
            groupBoxPridat.Controls.Add(txtPrurez);
            groupBoxPridat.Controls.Add(label4);
            groupBoxPridat.Controls.Add(txtPocetZil);
            groupBoxPridat.Controls.Add(label3);
            groupBoxPridat.Controls.Add(txtTyp);
            groupBoxPridat.Controls.Add(label2);
            groupBoxPridat.Controls.Add(txtOznaceni);
            groupBoxPridat.Controls.Add(comboBoxZnacka);
            groupBoxPridat.Controls.Add(lblZnacka);
            groupBoxPridat.Controls.Add(btnPridat);
            groupBoxPridat.Controls.Add(btnStorno);
            groupBoxPridat.Location = new Point(820, 215);
            groupBoxPridat.Name = "groupBoxPridat";
            groupBoxPridat.Size = new Size(328, 325);
            groupBoxPridat.TabIndex = 3;
            groupBoxPridat.TabStop = false;
            groupBoxPridat.Text = "Nový kabel";
            groupBoxPridat.Enter += groupBoxPridat_Enter;
            // 
            // lblPopis
            // 
            lblPopis.AutoSize = true;
            lblPopis.Location = new Point(7, 238);
            lblPopis.Name = "lblPopis";
            lblPopis.Size = new Size(119, 28);
            lblPopis.TabIndex = 12;
            lblPopis.Text = "Popis/pozn.:";
            // 
            // txtPopis
            // 
            txtPopis.Location = new Point(143, 235);
            txtPopis.Name = "txtPopis";
            txtPopis.Size = new Size(164, 34);
            txtPopis.TabIndex = 13;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(7, 203);
            label6.Name = "label6";
            label6.Size = new Size(99, 28);
            label6.TabIndex = 10;
            label6.Text = "Délka [m]:";
            // 
            // txtDelka
            // 
            txtDelka.Location = new Point(143, 200);
            txtDelka.Name = "txtDelka";
            txtDelka.Size = new Size(164, 34);
            txtDelka.TabIndex = 11;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(7, 168);
            label5.Name = "label5";
            label5.Size = new Size(133, 28);
            label5.TabIndex = 8;
            label5.Text = "Průřez [mm2]:";
            // 
            // txtPrurez
            // 
            txtPrurez.Location = new Point(143, 165);
            txtPrurez.Name = "txtPrurez";
            txtPrurez.Size = new Size(164, 34);
            txtPrurez.TabIndex = 9;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(7, 133);
            label4.Name = "label4";
            label4.Size = new Size(88, 28);
            label4.TabIndex = 6;
            label4.Text = "Počet žil:";
            // 
            // txtPocetZil
            // 
            txtPocetZil.Location = new Point(143, 130);
            txtPocetZil.Name = "txtPocetZil";
            txtPocetZil.Size = new Size(164, 34);
            txtPocetZil.TabIndex = 7;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(7, 98);
            label3.Name = "label3";
            label3.Size = new Size(110, 28);
            label3.TabIndex = 4;
            label3.Text = "Typ kabelu:";
            // 
            // txtTyp
            // 
            txtTyp.Location = new Point(143, 95);
            txtTyp.Name = "txtTyp";
            txtTyp.Size = new Size(164, 34);
            txtTyp.TabIndex = 5;
            txtTyp.Text = "CYKY-J";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(7, 63);
            label2.Name = "label2";
            label2.Size = new Size(96, 28);
            label2.TabIndex = 2;
            label2.Text = "Označení:";
            // 
            // txtOznaceni
            // 
            txtOznaceni.Location = new Point(143, 60);
            txtOznaceni.Name = "txtOznaceni";
            txtOznaceni.Size = new Size(164, 34);
            txtOznaceni.TabIndex = 3;
            // 
            // comboBoxZnacka
            // 
            comboBoxZnacka.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBoxZnacka.FormattingEnabled = true;
            comboBoxZnacka.Location = new Point(143, 25);
            comboBoxZnacka.Name = "comboBoxZnacka";
            comboBoxZnacka.Size = new Size(164, 36);
            comboBoxZnacka.TabIndex = 1;
            comboBoxZnacka.SelectedIndexChanged += ComboBoxZnacka_SelectedIndexChanged;
            // 
            // lblZnacka
            // 
            lblZnacka.AutoSize = true;
            lblZnacka.Location = new Point(7, 28);
            lblZnacka.Name = "lblZnacka";
            lblZnacka.Size = new Size(77, 28);
            lblZnacka.TabIndex = 0;
            lblZnacka.Text = "Značka:";
            // 
            // btnPridat
            // 
            btnPridat.Location = new Point(133, 275);
            btnPridat.Name = "btnPridat";
            btnPridat.Size = new Size(83, 35);
            btnPridat.TabIndex = 14;
            btnPridat.Text = "Přidat kabel";
            btnPridat.UseVisualStyleBackColor = true;
            btnPridat.Click += BtnPridat_Click;
            // 
            // btnStorno
            // 
            btnStorno.Location = new Point(222, 275);
            btnStorno.Name = "btnStorno";
            btnStorno.Size = new Size(82, 35);
            btnStorno.TabIndex = 15;
            btnStorno.Text = "Storno";
            btnStorno.UseVisualStyleBackColor = true;
            btnStorno.Visible = false;
            btnStorno.Click += BtnStorno_Click;
            // 
            // groupBoxRychlePridat
            // 
            groupBoxRychlePridat.Controls.Add(label7);
            groupBoxRychlePridat.Controls.Add(txtPrefixPower);
            groupBoxRychlePridat.Controls.Add(button1);
            groupBoxRychlePridat.Controls.Add(lblSekcePrefixy);
            groupBoxRychlePridat.Controls.Add(lblPrefixPTC);
            groupBoxRychlePridat.Controls.Add(txtPrefixPTC);
            groupBoxRychlePridat.Controls.Add(lblPrefixOvladani);
            groupBoxRychlePridat.Controls.Add(txtPrefixOvladani);
            groupBoxRychlePridat.Controls.Add(lblPrefixUTP);
            groupBoxRychlePridat.Controls.Add(txtPrefixUTP);
            groupBoxRychlePridat.Controls.Add(lblPrefixBinarni);
            groupBoxRychlePridat.Controls.Add(txtPrefixBinarni);
            groupBoxRychlePridat.Controls.Add(lblPrefixBlokovani);
            groupBoxRychlePridat.Controls.Add(txtPrefixBlokovani);
            groupBoxRychlePridat.Controls.Add(btnRychlyPTC);
            groupBoxRychlePridat.Controls.Add(btnRychlyOvladani5);
            groupBoxRychlePridat.Controls.Add(btnRychlyOvladani7);
            groupBoxRychlePridat.Controls.Add(btnRychlyOvladani12);
            groupBoxRychlePridat.Controls.Add(btnRychlyUTP);
            groupBoxRychlePridat.Controls.Add(btnRychlyBinarni);
            groupBoxRychlePridat.Controls.Add(btnRychlyBlokovani);
            groupBoxRychlePridat.Location = new Point(820, 545);
            groupBoxRychlePridat.Name = "groupBoxRychlePridat";
            groupBoxRychlePridat.Size = new Size(328, 250);
            groupBoxRychlePridat.TabIndex = 4;
            groupBoxRychlePridat.TabStop = false;
            groupBoxRychlePridat.Text = "Rychlé přidání";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("Segoe UI", 9F);
            label7.Location = new Point(191, 75);
            label7.Name = "label7";
            label7.Size = new Size(33, 20);
            label7.TabIndex = 23;
            label7.Text = "kW:";
            // 
            // txtPrefixPower
            // 
            txtPrefixPower.Font = new Font("Segoe UI", 9F);
            txtPrefixPower.Location = new Point(230, 77);
            txtPrefixPower.Name = "txtPrefixPower";
            txtPrefixPower.Size = new Size(40, 27);
            txtPrefixPower.TabIndex = 22;
            txtPrefixPower.Text = "WL";
            // 
            // button1
            // 
            button1.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            button1.Location = new Point(170, 215);
            button1.Name = "button1";
            button1.Size = new Size(148, 30);
            button1.TabIndex = 21;
            button1.Text = "JZ-600-Y-CY(FM)";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // lblSekcePrefixy
            // 
            lblSekcePrefixy.AutoSize = true;
            lblSekcePrefixy.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            lblSekcePrefixy.Location = new Point(10, 20);
            lblSekcePrefixy.Name = "lblSekcePrefixy";
            lblSekcePrefixy.Size = new Size(135, 23);
            lblSekcePrefixy.TabIndex = 10;
            lblSekcePrefixy.Text = "Prefixy značení:";
            // 
            // lblPrefixPTC
            // 
            lblPrefixPTC.AutoSize = true;
            lblPrefixPTC.Font = new Font("Segoe UI", 9F);
            lblPrefixPTC.Location = new Point(10, 45);
            lblPrefixPTC.Name = "lblPrefixPTC";
            lblPrefixPTC.Size = new Size(36, 20);
            lblPrefixPTC.TabIndex = 11;
            lblPrefixPTC.Text = "PTC:";
            // 
            // txtPrefixPTC
            // 
            txtPrefixPTC.Font = new Font("Segoe UI", 9F);
            txtPrefixPTC.Location = new Point(45, 42);
            txtPrefixPTC.Name = "txtPrefixPTC";
            txtPrefixPTC.Size = new Size(40, 27);
            txtPrefixPTC.TabIndex = 12;
            txtPrefixPTC.Text = "WS";
            // 
            // lblPrefixOvladani
            // 
            lblPrefixOvladani.AutoSize = true;
            lblPrefixOvladani.Font = new Font("Segoe UI", 9F);
            lblPrefixOvladani.Location = new Point(95, 45);
            lblPrefixOvladani.Name = "lblPrefixOvladani";
            lblPrefixOvladani.Size = new Size(54, 20);
            lblPrefixOvladani.TabIndex = 13;
            lblPrefixOvladani.Text = "Ovlád.:";
            // 
            // txtPrefixOvladani
            // 
            txtPrefixOvladani.Font = new Font("Segoe UI", 9F);
            txtPrefixOvladani.Location = new Point(145, 42);
            txtPrefixOvladani.Name = "txtPrefixOvladani";
            txtPrefixOvladani.Size = new Size(40, 27);
            txtPrefixOvladani.TabIndex = 14;
            txtPrefixOvladani.Text = "WS";
            // 
            // lblPrefixUTP
            // 
            lblPrefixUTP.AutoSize = true;
            lblPrefixUTP.Font = new Font("Segoe UI", 9F);
            lblPrefixUTP.Location = new Point(194, 45);
            lblPrefixUTP.Name = "lblPrefixUTP";
            lblPrefixUTP.Size = new Size(38, 20);
            lblPrefixUTP.TabIndex = 15;
            lblPrefixUTP.Text = "UTP:";
            // 
            // txtPrefixUTP
            // 
            txtPrefixUTP.Font = new Font("Segoe UI", 9F);
            txtPrefixUTP.Location = new Point(230, 42);
            txtPrefixUTP.Name = "txtPrefixUTP";
            txtPrefixUTP.Size = new Size(40, 27);
            txtPrefixUTP.TabIndex = 16;
            txtPrefixUTP.Text = "WD";
            // 
            // lblPrefixBinarni
            // 
            lblPrefixBinarni.AutoSize = true;
            lblPrefixBinarni.Font = new Font("Segoe UI", 9F);
            lblPrefixBinarni.Location = new Point(10, 75);
            lblPrefixBinarni.Name = "lblPrefixBinarni";
            lblPrefixBinarni.Size = new Size(36, 20);
            lblPrefixBinarni.TabIndex = 17;
            lblPrefixBinarni.Text = "Bin.:";
            // 
            // txtPrefixBinarni
            // 
            txtPrefixBinarni.Font = new Font("Segoe UI", 9F);
            txtPrefixBinarni.Location = new Point(45, 72);
            txtPrefixBinarni.Name = "txtPrefixBinarni";
            txtPrefixBinarni.Size = new Size(40, 27);
            txtPrefixBinarni.TabIndex = 18;
            txtPrefixBinarni.Text = "WA";
            // 
            // lblPrefixBlokovani
            // 
            lblPrefixBlokovani.AutoSize = true;
            lblPrefixBlokovani.Font = new Font("Segoe UI", 9F);
            lblPrefixBlokovani.Location = new Point(95, 75);
            lblPrefixBlokovani.Name = "lblPrefixBlokovani";
            lblPrefixBlokovani.Size = new Size(44, 20);
            lblPrefixBlokovani.TabIndex = 19;
            lblPrefixBlokovani.Text = "Blok.:";
            // 
            // txtPrefixBlokovani
            // 
            txtPrefixBlokovani.Font = new Font("Segoe UI", 9F);
            txtPrefixBlokovani.Location = new Point(145, 72);
            txtPrefixBlokovani.Name = "txtPrefixBlokovani";
            txtPrefixBlokovani.Size = new Size(40, 27);
            txtPrefixBlokovani.TabIndex = 20;
            txtPrefixBlokovani.Text = "WB";
            // 
            // btnRychlyPTC
            // 
            btnRychlyPTC.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            btnRychlyPTC.Location = new Point(10, 110);
            btnRychlyPTC.Name = "btnRychlyPTC";
            btnRychlyPTC.Size = new Size(148, 30);
            btnRychlyPTC.TabIndex = 0;
            btnRychlyPTC.Text = "+ PTC (2 vodiče)";
            btnRychlyPTC.UseVisualStyleBackColor = true;
            btnRychlyPTC.Click += BtnRychlyPTC_Click;
            // 
            // btnRychlyOvladani5
            // 
            btnRychlyOvladani5.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            btnRychlyOvladani5.Location = new Point(10, 145);
            btnRychlyOvladani5.Name = "btnRychlyOvladani5";
            btnRychlyOvladani5.Size = new Size(148, 30);
            btnRychlyOvladani5.TabIndex = 2;
            btnRychlyOvladani5.Text = "+ Ovládání (5 v.)";
            btnRychlyOvladani5.UseVisualStyleBackColor = true;
            btnRychlyOvladani5.Click += BtnRychlyOvladani5_Click;
            // 
            // btnRychlyOvladani7
            // 
            btnRychlyOvladani7.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            btnRychlyOvladani7.Location = new Point(170, 145);
            btnRychlyOvladani7.Name = "btnRychlyOvladani7";
            btnRychlyOvladani7.Size = new Size(148, 30);
            btnRychlyOvladani7.TabIndex = 3;
            btnRychlyOvladani7.Text = "+ Ovládání (7 v.)";
            btnRychlyOvladani7.UseVisualStyleBackColor = true;
            btnRychlyOvladani7.Click += BtnRychlyOvladani7_Click;
            // 
            // btnRychlyOvladani12
            // 
            btnRychlyOvladani12.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            btnRychlyOvladani12.Location = new Point(10, 180);
            btnRychlyOvladani12.Name = "btnRychlyOvladani12";
            btnRychlyOvladani12.Size = new Size(148, 30);
            btnRychlyOvladani12.TabIndex = 4;
            btnRychlyOvladani12.Text = "+ Ovládání (12 v.)";
            btnRychlyOvladani12.UseVisualStyleBackColor = true;
            btnRychlyOvladani12.Click += BtnRychlyOvladani12_Click;
            // 
            // btnRychlyUTP
            // 
            btnRychlyUTP.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            btnRychlyUTP.Location = new Point(170, 110);
            btnRychlyUTP.Name = "btnRychlyUTP";
            btnRychlyUTP.Size = new Size(148, 30);
            btnRychlyUTP.TabIndex = 1;
            btnRychlyUTP.Text = "+ UTP Cat6";
            btnRychlyUTP.UseVisualStyleBackColor = true;
            btnRychlyUTP.Click += BtnRychlyUTP_Click;
            // 
            // btnRychlyBinarni
            // 
            btnRychlyBinarni.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            btnRychlyBinarni.Location = new Point(170, 180);
            btnRychlyBinarni.Name = "btnRychlyBinarni";
            btnRychlyBinarni.Size = new Size(148, 30);
            btnRychlyBinarni.TabIndex = 5;
            btnRychlyBinarni.Text = "+ Binární";
            btnRychlyBinarni.UseVisualStyleBackColor = true;
            btnRychlyBinarni.Click += BtnRychlyBinarni_Click;
            // 
            // btnRychlyBlokovani
            // 
            btnRychlyBlokovani.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 238);
            btnRychlyBlokovani.Location = new Point(10, 215);
            btnRychlyBlokovani.Name = "btnRychlyBlokovani";
            btnRychlyBlokovani.Size = new Size(148, 30);
            btnRychlyBlokovani.TabIndex = 6;
            btnRychlyBlokovani.Text = "+ Blokování";
            btnRychlyBlokovani.UseVisualStyleBackColor = true;
            btnRychlyBlokovani.Click += BtnRychlyBlokovani_Click;
            // 
            // btnSmazat
            // 
            btnSmazat.Location = new Point(12, 745);
            btnSmazat.Name = "btnSmazat";
            btnSmazat.Size = new Size(160, 35);
            btnSmazat.TabIndex = 5;
            btnSmazat.Text = "Smazat vybraný";
            btnSmazat.UseVisualStyleBackColor = true;
            btnSmazat.Click += BtnSmazat_Click;
            // 
            // lblStatistika
            // 
            lblStatistika.AutoSize = true;
            lblStatistika.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            lblStatistika.Location = new Point(190, 752);
            lblStatistika.Name = "lblStatistika";
            lblStatistika.Size = new Size(220, 25);
            lblStatistika.TabIndex = 6;
            lblStatistika.Text = "Počet kabelů zařízení: 0";
            // 
            // btnZavrit
            // 
            btnZavrit.Location = new Point(998, 805);
            btnZavrit.Name = "btnZavrit";
            btnZavrit.Size = new Size(150, 35);
            btnZavrit.TabIndex = 7;
            btnZavrit.Text = "Zavřít";
            btnZavrit.UseVisualStyleBackColor = true;
            btnZavrit.Click += BtnZavrit_Click;
            // 
            // FormKabely
            // 
            AutoScaleDimensions = new SizeF(11F, 28F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1160, 855);
            Controls.Add(btnZavrit);
            Controls.Add(lblStatistika);
            Controls.Add(btnSmazat);
            Controls.Add(groupBoxRychlePridat);
            Controls.Add(groupBoxPridat);
            Controls.Add(dataGridViewKabely);
            Controls.Add(groupBoxInfoZarizeni);
            Controls.Add(groupBoxFiltry);
            Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 238);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Margin = new Padding(4);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FormKabely";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "Správa kabelů zařízení";
            Load += FormKabely_Load;
            groupBoxFiltry.ResumeLayout(false);
            groupBoxFiltry.PerformLayout();
            groupBoxInfoZarizeni.ResumeLayout(false);
            groupBoxInfoZarizeni.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewKabely).EndInit();
            groupBoxPridat.ResumeLayout(false);
            groupBoxPridat.PerformLayout();
            groupBoxRychlePridat.ResumeLayout(false);
            groupBoxRychlePridat.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxZarizeni;
        private System.Windows.Forms.DataGridView dataGridViewKabely;
        private System.Windows.Forms.GroupBox groupBoxPridat;
        private System.Windows.Forms.Label lblZnacka;
        private System.Windows.Forms.ComboBox comboBoxZnacka;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOznaceni;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtTyp;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPocetZil;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtPrurez;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtDelka;
        private System.Windows.Forms.Label lblPopis;
        private System.Windows.Forms.TextBox txtPopis;
        private System.Windows.Forms.Button btnPridat;
        private System.Windows.Forms.Button btnSmazat;
        private System.Windows.Forms.Label lblStatistika;
        private System.Windows.Forms.Button btnZavrit;

        // Nové filtry
        private System.Windows.Forms.GroupBox groupBoxFiltry;
        private System.Windows.Forms.Label lblFilterExistElektro;
        private System.Windows.Forms.ComboBox comboBoxFilterExistElektro;
        private System.Windows.Forms.Label lblFilterEtapa;
        private System.Windows.Forms.ComboBox comboBoxFilterEtapa;

        // Informace o zařízení
        private System.Windows.Forms.GroupBox groupBoxInfoZarizeni;
        private System.Windows.Forms.Label lblInfoPopis;
        private System.Windows.Forms.TextBox txtInfoPopis;
        private System.Windows.Forms.Label lblInfoPrikon;
        private System.Windows.Forms.TextBox txtInfoPrikon;
        private System.Windows.Forms.Label lblInfoPrikonStroj;
        private System.Windows.Forms.TextBox txtInfoPrikonStroj;
        private System.Windows.Forms.Label lblInfoProud;
        private System.Windows.Forms.TextBox txtInfoProud;
        private System.Windows.Forms.Label lblInfoNapeti;
        private System.Windows.Forms.TextBox txtInfoNapeti;
        private System.Windows.Forms.Label lblInfoMenic;
        private System.Windows.Forms.TextBox txtInfoMenic;
        private System.Windows.Forms.Label lblInfoBalena;
        private System.Windows.Forms.TextBox txtInfoBalena;
        private System.Windows.Forms.Label lblInfoDruh;
        private System.Windows.Forms.TextBox txtInfoDruh;

        // Rychlé přidávání kabelů
        private System.Windows.Forms.GroupBox groupBoxRychlePridat;
        private System.Windows.Forms.Label lblSekcePrefixy;
        private System.Windows.Forms.Label lblPrefixPTC;
        private System.Windows.Forms.TextBox txtPrefixPTC;
        private System.Windows.Forms.Label lblPrefixOvladani;
        private System.Windows.Forms.TextBox txtPrefixOvladani;
        private System.Windows.Forms.Label lblPrefixUTP;
        private System.Windows.Forms.TextBox txtPrefixUTP;
        private System.Windows.Forms.Label lblPrefixBinarni;
        private System.Windows.Forms.TextBox txtPrefixBinarni;
        private System.Windows.Forms.Label lblPrefixBlokovani;
        private System.Windows.Forms.TextBox txtPrefixBlokovani;
        private System.Windows.Forms.Button btnRychlyPTC;
        private System.Windows.Forms.Button btnRychlyOvladani5;
        private System.Windows.Forms.Button btnRychlyOvladani7;
        private System.Windows.Forms.Button btnRychlyOvladani12;
        private System.Windows.Forms.Button btnRychlyUTP;
        private System.Windows.Forms.Button btnRychlyBinarni;
        private System.Windows.Forms.Button btnRychlyBlokovani;
        private System.Windows.Forms.Button btnStorno;
        private CheckBox CheckBoxIsExist;
        private TextBox FiltrText;
        private Button button1;
        private TextBox txtPrefixPower;
        private Label label7;
    }
}
