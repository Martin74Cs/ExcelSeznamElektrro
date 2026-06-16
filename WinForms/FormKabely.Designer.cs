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

        private void InitializeComponent()
        {
            this.groupBoxFiltry = new System.Windows.Forms.GroupBox();
            this.lblFilterExist = new System.Windows.Forms.Label();
            this.comboBoxFilterExist = new System.Windows.Forms.ComboBox();
            this.lblFilterExistElektro = new System.Windows.Forms.Label();
            this.comboBoxFilterExistElektro = new System.Windows.Forms.ComboBox();
            this.lblFilterEtapa = new System.Windows.Forms.Label();
            this.comboBoxFilterEtapa = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxZarizeni = new System.Windows.Forms.ComboBox();
            this.groupBoxInfoZarizeni = new System.Windows.Forms.GroupBox();
            this.lblInfoPopis = new System.Windows.Forms.Label();
            this.txtInfoPopis = new System.Windows.Forms.TextBox();
            this.lblInfoPrikon = new System.Windows.Forms.Label();
            this.txtInfoPrikon = new System.Windows.Forms.TextBox();
            this.lblInfoPrikonStroj = new System.Windows.Forms.Label();
            this.txtInfoPrikonStroj = new System.Windows.Forms.TextBox();
            this.lblInfoProud = new System.Windows.Forms.Label();
            this.txtInfoProud = new System.Windows.Forms.TextBox();
            this.lblInfoNapeti = new System.Windows.Forms.Label();
            this.txtInfoNapeti = new System.Windows.Forms.TextBox();
            this.lblInfoMenic = new System.Windows.Forms.Label();
            this.txtInfoMenic = new System.Windows.Forms.TextBox();
            this.lblInfoBalena = new System.Windows.Forms.Label();
            this.txtInfoBalena = new System.Windows.Forms.TextBox();
            this.lblInfoDruh = new System.Windows.Forms.Label();
            this.txtInfoDruh = new System.Windows.Forms.TextBox();
            this.dataGridViewKabely = new System.Windows.Forms.DataGridView();
            this.groupBoxPridat = new System.Windows.Forms.GroupBox();
            this.lblPopis = new System.Windows.Forms.Label();
            this.txtPopis = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtDelka = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtPrurez = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPocetZil = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtTyp = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOznaceni = new System.Windows.Forms.TextBox();
            this.comboBoxZnacka = new System.Windows.Forms.ComboBox();
            this.lblZnacka = new System.Windows.Forms.Label();
            this.btnPridat = new System.Windows.Forms.Button();
            this.btnStorno = new System.Windows.Forms.Button();
            this.groupBoxRychlePridat = new System.Windows.Forms.GroupBox();
            this.lblSekcePrefixy = new System.Windows.Forms.Label();
            this.lblPrefixPTC = new System.Windows.Forms.Label();
            this.txtPrefixPTC = new System.Windows.Forms.TextBox();
            this.lblPrefixOvladani = new System.Windows.Forms.Label();
            this.txtPrefixOvladani = new System.Windows.Forms.TextBox();
            this.lblPrefixUTP = new System.Windows.Forms.Label();
            this.txtPrefixUTP = new System.Windows.Forms.TextBox();
            this.lblPrefixBinarni = new System.Windows.Forms.Label();
            this.txtPrefixBinarni = new System.Windows.Forms.TextBox();
            this.lblPrefixBlokovani = new System.Windows.Forms.Label();
            this.txtPrefixBlokovani = new System.Windows.Forms.TextBox();
            this.btnRychlyPTC = new System.Windows.Forms.Button();
            this.btnRychlyOvladani5 = new System.Windows.Forms.Button();
            this.btnRychlyOvladani7 = new System.Windows.Forms.Button();
            this.btnRychlyOvladani12 = new System.Windows.Forms.Button();
            this.btnRychlyUTP = new System.Windows.Forms.Button();
            this.btnRychlyBinarni = new System.Windows.Forms.Button();
            this.btnRychlyBlokovani = new System.Windows.Forms.Button();
            this.btnSmazat = new System.Windows.Forms.Button();
            this.lblStatistika = new System.Windows.Forms.Label();
            this.btnZavrit = new System.Windows.Forms.Button();
            this.groupBoxFiltry.SuspendLayout();
            this.groupBoxInfoZarizeni.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewKabely)).BeginInit();
            this.groupBoxPridat.SuspendLayout();
            this.groupBoxRychlePridat.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxFiltry
            // 
            this.groupBoxFiltry.Controls.Add(this.lblFilterExist);
            this.groupBoxFiltry.Controls.Add(this.comboBoxFilterExist);
            this.groupBoxFiltry.Controls.Add(this.lblFilterExistElektro);
            this.groupBoxFiltry.Controls.Add(this.comboBoxFilterExistElektro);
            this.groupBoxFiltry.Controls.Add(this.lblFilterEtapa);
            this.groupBoxFiltry.Controls.Add(this.comboBoxFilterEtapa);
            this.groupBoxFiltry.Controls.Add(this.label1);
            this.groupBoxFiltry.Controls.Add(this.comboBoxZarizeni);
            this.groupBoxFiltry.Location = new System.Drawing.Point(12, 12);
            this.groupBoxFiltry.Name = "groupBoxFiltry";
            this.groupBoxFiltry.Size = new System.Drawing.Size(1136, 90);
            this.groupBoxFiltry.TabIndex = 0;
            this.groupBoxFiltry.TabStop = false;
            this.groupBoxFiltry.Text = "Filtry a výběr zařízení";
            // 
            // lblFilterExist
            // 
            this.lblFilterExist.AutoSize = true;
            this.lblFilterExist.Location = new System.Drawing.Point(12, 23);
            this.lblFilterExist.Name = "lblFilterExist";
            this.lblFilterExist.Size = new System.Drawing.Size(124, 21);
            this.lblFilterExist.TabIndex = 0;
            this.lblFilterExist.Text = "Existence v proj.:";
            // 
            // comboBoxFilterExist
            // 
            this.comboBoxFilterExist.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxFilterExist.FormattingEnabled = true;
            this.comboBoxFilterExist.Location = new System.Drawing.Point(12, 47);
            this.comboBoxFilterExist.Name = "comboBoxFilterExist";
            this.comboBoxFilterExist.Size = new System.Drawing.Size(140, 29);
            this.comboBoxFilterExist.TabIndex = 1;
            this.comboBoxFilterExist.SelectedIndexChanged += new System.EventHandler(this.ComboBoxFilter_SelectedIndexChanged);
            // 
            // lblFilterExistElektro
            // 
            this.lblFilterExistElektro.AutoSize = true;
            this.lblFilterExistElektro.Location = new System.Drawing.Point(165, 23);
            this.lblFilterExistElektro.Name = "lblFilterExistElektro";
            this.lblFilterExistElektro.Size = new System.Drawing.Size(95, 21);
            this.lblFilterExistElektro.TabIndex = 2;
            this.lblFilterExistElektro.Text = "Elektro blok:";
            // 
            // comboBoxFilterExistElektro
            // 
            this.comboBoxFilterExistElektro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxFilterExistElektro.FormattingEnabled = true;
            this.comboBoxFilterExistElektro.Location = new System.Drawing.Point(165, 47);
            this.comboBoxFilterExistElektro.Name = "comboBoxFilterExistElektro";
            this.comboBoxFilterExistElektro.Size = new System.Drawing.Size(140, 29);
            this.comboBoxFilterExistElektro.TabIndex = 3;
            this.comboBoxFilterExistElektro.SelectedIndexChanged += new System.EventHandler(this.ComboBoxFilter_SelectedIndexChanged);
            // 
            // lblFilterEtapa
            // 
            this.lblFilterEtapa.AutoSize = true;
            this.lblFilterEtapa.Location = new System.Drawing.Point(318, 23);
            this.lblFilterEtapa.Name = "lblFilterEtapa";
            this.lblFilterEtapa.Size = new System.Drawing.Size(108, 21);
            this.lblFilterEtapa.TabIndex = 4;
            this.lblFilterEtapa.Text = "Fáze výstavby:";
            // 
            // comboBoxFilterEtapa
            // 
            this.comboBoxFilterEtapa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxFilterEtapa.FormattingEnabled = true;
            this.comboBoxFilterEtapa.Location = new System.Drawing.Point(318, 47);
            this.comboBoxFilterEtapa.Name = "comboBoxFilterEtapa";
            this.comboBoxFilterEtapa.Size = new System.Drawing.Size(140, 29);
            this.comboBoxFilterEtapa.TabIndex = 5;
            this.comboBoxFilterEtapa.SelectedIndexChanged += new System.EventHandler(this.ComboBoxFilter_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(471, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 21);
            this.label1.TabIndex = 6;
            this.label1.Text = "Zařízení:";
            // 
            // comboBoxZarizeni
            // 
            this.comboBoxZarizeni.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxZarizeni.FormattingEnabled = true;
            this.comboBoxZarizeni.Location = new System.Drawing.Point(471, 47);
            this.comboBoxZarizeni.Name = "comboBoxZarizeni";
            this.comboBoxZarizeni.Size = new System.Drawing.Size(650, 29);
            this.comboBoxZarizeni.TabIndex = 7;
            this.comboBoxZarizeni.SelectedIndexChanged += new System.EventHandler(this.ComboBoxZarizeni_SelectedIndexChanged);
            // 
            // groupBoxInfoZarizeni
            // 
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoPopis);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoPopis);
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoPrikon);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoPrikon);
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoPrikonStroj);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoPrikonStroj);
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoProud);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoProud);
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoNapeti);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoNapeti);
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoMenic);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoMenic);
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoBalena);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoBalena);
            this.groupBoxInfoZarizeni.Controls.Add(this.lblInfoDruh);
            this.groupBoxInfoZarizeni.Controls.Add(this.txtInfoDruh);
            this.groupBoxInfoZarizeni.Location = new System.Drawing.Point(12, 110);
            this.groupBoxInfoZarizeni.Name = "groupBoxInfoZarizeni";
            this.groupBoxInfoZarizeni.Size = new System.Drawing.Size(1136, 95);
            this.groupBoxInfoZarizeni.TabIndex = 1;
            this.groupBoxInfoZarizeni.TabStop = false;
            this.groupBoxInfoZarizeni.Text = "Základní informace o zařízení";
            // 
            // lblInfoPopis
            // 
            this.lblInfoPopis.AutoSize = true;
            this.lblInfoPopis.Location = new System.Drawing.Point(12, 23);
            this.lblInfoPopis.Name = "lblInfoPopis";
            this.lblInfoPopis.Size = new System.Drawing.Size(107, 21);
            this.lblInfoPopis.TabIndex = 0;
            this.lblInfoPopis.Text = "Název (Popis):";
            // 
            // txtInfoPopis
            // 
            this.txtInfoPopis.Location = new System.Drawing.Point(12, 46);
            this.txtInfoPopis.Name = "txtInfoPopis";
            this.txtInfoPopis.ReadOnly = true;
            this.txtInfoPopis.Size = new System.Drawing.Size(250, 29);
            this.txtInfoPopis.TabIndex = 1;
            // 
            // lblInfoPrikon
            // 
            this.lblInfoPrikon.AutoSize = true;
            this.lblInfoPrikon.Location = new System.Drawing.Point(272, 23);
            this.lblInfoPrikon.Name = "lblInfoPrikon";
            this.lblInfoPrikon.Size = new System.Drawing.Size(95, 21);
            this.lblInfoPrikon.TabIndex = 2;
            this.lblInfoPrikon.Text = "kW (elektro):";
            // 
            // txtInfoPrikon
            // 
            this.txtInfoPrikon.Location = new System.Drawing.Point(272, 46);
            this.txtInfoPrikon.Name = "txtInfoPrikon";
            this.txtInfoPrikon.ReadOnly = true;
            this.txtInfoPrikon.Size = new System.Drawing.Size(90, 29);
            this.txtInfoPrikon.TabIndex = 3;
            // 
            // lblInfoPrikonStroj
            // 
            this.lblInfoPrikonStroj.AutoSize = true;
            this.lblInfoPrikonStroj.Location = new System.Drawing.Point(372, 23);
            this.lblInfoPrikonStroj.Name = "lblInfoPrikonStroj";
            this.lblInfoPrikonStroj.Size = new System.Drawing.Size(83, 21);
            this.lblInfoPrikonStroj.TabIndex = 4;
            this.lblInfoPrikonStroj.Text = "kW (stroj):";
            // 
            // txtInfoPrikonStroj
            // 
            this.txtInfoPrikonStroj.Location = new System.Drawing.Point(372, 46);
            this.txtInfoPrikonStroj.Name = "txtInfoPrikonStroj";
            this.txtInfoPrikonStroj.ReadOnly = true;
            this.txtInfoPrikonStroj.Size = new System.Drawing.Size(90, 29);
            this.txtInfoPrikonStroj.TabIndex = 5;
            // 
            // lblInfoProud
            // 
            this.lblInfoProud.AutoSize = true;
            this.lblInfoProud.Location = new System.Drawing.Point(472, 23);
            this.lblInfoProud.Name = "lblInfoProud";
            this.lblInfoProud.Size = new System.Drawing.Size(78, 21);
            this.lblInfoProud.TabIndex = 6;
            this.lblInfoProud.Text = "Proud [A]:";
            // 
            // txtInfoProud
            // 
            this.txtInfoProud.Location = new System.Drawing.Point(472, 46);
            this.txtInfoProud.Name = "txtInfoProud";
            this.txtInfoProud.ReadOnly = true;
            this.txtInfoProud.Size = new System.Drawing.Size(80, 29);
            this.txtInfoProud.TabIndex = 7;
            // 
            // lblInfoNapeti
            // 
            this.lblInfoNapeti.AutoSize = true;
            this.lblInfoNapeti.Location = new System.Drawing.Point(562, 23);
            this.lblInfoNapeti.Name = "lblInfoNapeti";
            this.lblInfoNapeti.Size = new System.Drawing.Size(85, 21);
            this.lblInfoNapeti.TabIndex = 8;
            this.lblInfoNapeti.Text = "Napětí [V]:";
            // 
            // txtInfoNapeti
            // 
            this.txtInfoNapeti.Location = new System.Drawing.Point(562, 46);
            this.txtInfoNapeti.Name = "txtInfoNapeti";
            this.txtInfoNapeti.ReadOnly = true;
            this.txtInfoNapeti.Size = new System.Drawing.Size(80, 29);
            this.txtInfoNapeti.TabIndex = 9;
            // 
            // lblInfoMenic
            // 
            this.lblInfoMenic.AutoSize = true;
            this.lblInfoMenic.Location = new System.Drawing.Point(652, 23);
            this.lblInfoMenic.Name = "lblInfoMenic";
            this.lblInfoMenic.Size = new System.Drawing.Size(57, 21);
            this.lblInfoMenic.TabIndex = 10;
            this.lblInfoMenic.Text = "Měnič:";
            // 
            // txtInfoMenic
            // 
            this.txtInfoMenic.Location = new System.Drawing.Point(652, 46);
            this.txtInfoMenic.Name = "txtInfoMenic";
            this.txtInfoMenic.ReadOnly = true;
            this.txtInfoMenic.Size = new System.Drawing.Size(110, 29);
            this.txtInfoMenic.TabIndex = 11;
            // 
            // lblInfoBalena
            // 
            this.lblInfoBalena.AutoSize = true;
            this.lblInfoBalena.Location = new System.Drawing.Point(772, 23);
            this.lblInfoBalena.Name = "lblInfoBalena";
            this.lblInfoBalena.Size = new System.Drawing.Size(126, 21);
            this.lblInfoBalena.TabIndex = 12;
            this.lblInfoBalena.Text = "Balená jednotka:";
            // 
            // txtInfoBalena
            // 
            this.txtInfoBalena.Location = new System.Drawing.Point(772, 46);
            this.txtInfoBalena.Name = "txtInfoBalena";
            this.txtInfoBalena.ReadOnly = true;
            this.txtInfoBalena.Size = new System.Drawing.Size(160, 29);
            this.txtInfoBalena.TabIndex = 13;
            // 
            // lblInfoDruh
            // 
            this.lblInfoDruh.AutoSize = true;
            this.lblInfoDruh.Location = new System.Drawing.Point(942, 23);
            this.lblInfoDruh.Name = "lblInfoDruh";
            this.lblInfoDruh.Size = new System.Drawing.Size(111, 21);
            this.lblInfoDruh.TabIndex = 14;
            this.lblInfoDruh.Text = "Druh zařízení:";
            // 
            // txtInfoDruh
            // 
            this.txtInfoDruh.Location = new System.Drawing.Point(942, 46);
            this.txtInfoDruh.Name = "txtInfoDruh";
            this.txtInfoDruh.ReadOnly = true;
            this.txtInfoDruh.Size = new System.Drawing.Size(180, 29);
            this.txtInfoDruh.TabIndex = 15;
            // 
            // dataGridViewKabely
            // 
            this.dataGridViewKabely.AllowUserToAddRows = false;
            this.dataGridViewKabely.AllowUserToDeleteRows = false;
            this.dataGridViewKabely.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewKabely.Location = new System.Drawing.Point(12, 215);
            this.dataGridViewKabely.MultiSelect = false;
            this.dataGridViewKabely.Name = "dataGridViewKabely";
            this.dataGridViewKabely.ReadOnly = true;
            this.dataGridViewKabely.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewKabely.Size = new System.Drawing.Size(796, 520);
            this.dataGridViewKabely.TabIndex = 2;
            this.dataGridViewKabely.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridViewKabely_CellDoubleClick);
            // 
            // groupBoxPridat
            // 
            this.groupBoxPridat.Controls.Add(this.lblPopis);
            this.groupBoxPridat.Controls.Add(this.txtPopis);
            this.groupBoxPridat.Controls.Add(this.label6);
            this.groupBoxPridat.Controls.Add(this.txtDelka);
            this.groupBoxPridat.Controls.Add(this.label5);
            this.groupBoxPridat.Controls.Add(this.txtPrurez);
            this.groupBoxPridat.Controls.Add(this.label4);
            this.groupBoxPridat.Controls.Add(this.txtPocetZil);
            this.groupBoxPridat.Controls.Add(this.label3);
            this.groupBoxPridat.Controls.Add(this.txtTyp);
            this.groupBoxPridat.Controls.Add(this.label2);
            this.groupBoxPridat.Controls.Add(this.txtOznaceni);
            this.groupBoxPridat.Controls.Add(this.comboBoxZnacka);
            this.groupBoxPridat.Controls.Add(this.lblZnacka);
            this.groupBoxPridat.Controls.Add(this.btnPridat);
            this.groupBoxPridat.Controls.Add(this.btnStorno);
            this.groupBoxPridat.Location = new System.Drawing.Point(820, 215);
            this.groupBoxPridat.Name = "groupBoxPridat";
            this.groupBoxPridat.Size = new System.Drawing.Size(328, 325);
            this.groupBoxPridat.TabIndex = 3;
            this.groupBoxPridat.TabStop = false;
            this.groupBoxPridat.Text = "Nový kabel";
            // 
            // lblPopis
            // 
            this.lblPopis.AutoSize = true;
            this.lblPopis.Location = new System.Drawing.Point(16, 238);
            this.lblPopis.Name = "lblPopis";
            this.lblPopis.Size = new System.Drawing.Size(95, 21);
            this.lblPopis.TabIndex = 12;
            this.lblPopis.Text = "Popis/pozn.:";
            // 
            // txtPopis
            // 
            this.txtPopis.Location = new System.Drawing.Point(130, 235);
            this.txtPopis.Name = "txtPopis";
            this.txtPopis.Size = new System.Drawing.Size(182, 29);
            this.txtPopis.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(16, 203);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(81, 21);
            this.label6.TabIndex = 10;
            this.label6.Text = "Délka [m]:";
            // 
            // txtDelka
            // 
            this.txtDelka.Location = new System.Drawing.Point(130, 200);
            this.txtDelka.Name = "txtDelka";
            this.txtDelka.Size = new System.Drawing.Size(182, 29);
            this.txtDelka.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 168);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(106, 21);
            this.label5.TabIndex = 8;
            this.label5.Text = "Průřez [mm2]:";
            // 
            // txtPrurez
            // 
            this.txtPrurez.Location = new System.Drawing.Point(130, 165);
            this.txtPrurez.Name = "txtPrurez";
            this.txtPrurez.Size = new System.Drawing.Size(182, 29);
            this.txtPrurez.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 133);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 21);
            this.label4.TabIndex = 6;
            this.label4.Text = "Počet žil:";
            // 
            // txtPocetZil
            // 
            this.txtPocetZil.Location = new System.Drawing.Point(130, 130);
            this.txtPocetZil.Name = "txtPocetZil";
            this.txtPocetZil.Size = new System.Drawing.Size(182, 29);
            this.txtPocetZil.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 21);
            this.label3.TabIndex = 4;
            this.label3.Text = "Typ kabelu:";
            // 
            // txtTyp
            // 
            this.txtTyp.Location = new System.Drawing.Point(130, 95);
            this.txtTyp.Name = "txtTyp";
            this.txtTyp.Size = new System.Drawing.Size(182, 29);
            this.txtTyp.TabIndex = 5;
            this.txtTyp.Text = "CYKY-J";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 21);
            this.label2.TabIndex = 2;
            this.label2.Text = "Označení:";
            // 
            // txtOznaceni
            // 
            this.txtOznaceni.Location = new System.Drawing.Point(130, 60);
            this.txtOznaceni.Name = "txtOznaceni";
            this.txtOznaceni.Size = new System.Drawing.Size(182, 29);
            this.txtOznaceni.TabIndex = 3;
            // 
            // comboBoxZnacka
            // 
            this.comboBoxZnacka.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxZnacka.FormattingEnabled = true;
            this.comboBoxZnacka.Location = new System.Drawing.Point(130, 25);
            this.comboBoxZnacka.Name = "comboBoxZnacka";
            this.comboBoxZnacka.Size = new System.Drawing.Size(182, 29);
            this.comboBoxZnacka.TabIndex = 1;
            this.comboBoxZnacka.SelectedIndexChanged += new System.EventHandler(this.ComboBoxZnacka_SelectedIndexChanged);
            // 
            // lblZnacka
            // 
            this.lblZnacka.AutoSize = true;
            this.lblZnacka.Location = new System.Drawing.Point(16, 28);
            this.lblZnacka.Name = "lblZnacka";
            this.lblZnacka.Size = new System.Drawing.Size(62, 21);
            this.lblZnacka.TabIndex = 0;
            this.lblZnacka.Text = "Značka:";
            // 
            // btnPridat
            // 
            this.btnPridat.Location = new System.Drawing.Point(130, 275);
            this.btnPridat.Name = "btnPridat";
            this.btnPridat.Size = new System.Drawing.Size(100, 35);
            this.btnPridat.TabIndex = 14;
            this.btnPridat.Text = "Přidat kabel";
            this.btnPridat.UseVisualStyleBackColor = true;
            this.btnPridat.Click += new System.EventHandler(this.BtnPridat_Click);
            // 
            // btnStorno
            // 
            this.btnStorno.Location = new System.Drawing.Point(236, 275);
            this.btnStorno.Name = "btnStorno";
            this.btnStorno.Size = new System.Drawing.Size(76, 35);
            this.btnStorno.TabIndex = 15;
            this.btnStorno.Text = "Storno";
            this.btnStorno.UseVisualStyleBackColor = true;
            this.btnStorno.Visible = false;
            this.btnStorno.Click += new System.EventHandler(this.BtnStorno_Click);
            // 
            // groupBoxRychlePridat
            // 
            this.groupBoxRychlePridat.Controls.Add(this.lblSekcePrefixy);
            this.groupBoxRychlePridat.Controls.Add(this.lblPrefixPTC);
            this.groupBoxRychlePridat.Controls.Add(this.txtPrefixPTC);
            this.groupBoxRychlePridat.Controls.Add(this.lblPrefixOvladani);
            this.groupBoxRychlePridat.Controls.Add(this.txtPrefixOvladani);
            this.groupBoxRychlePridat.Controls.Add(this.lblPrefixUTP);
            this.groupBoxRychlePridat.Controls.Add(this.txtPrefixUTP);
            this.groupBoxRychlePridat.Controls.Add(this.lblPrefixBinarni);
            this.groupBoxRychlePridat.Controls.Add(this.txtPrefixBinarni);
            this.groupBoxRychlePridat.Controls.Add(this.lblPrefixBlokovani);
            this.groupBoxRychlePridat.Controls.Add(this.txtPrefixBlokovani);
            this.groupBoxRychlePridat.Controls.Add(this.btnRychlyPTC);
            this.groupBoxRychlePridat.Controls.Add(this.btnRychlyOvladani5);
            this.groupBoxRychlePridat.Controls.Add(this.btnRychlyOvladani7);
            this.groupBoxRychlePridat.Controls.Add(this.btnRychlyOvladani12);
            this.groupBoxRychlePridat.Controls.Add(this.btnRychlyUTP);
            this.groupBoxRychlePridat.Controls.Add(this.btnRychlyBinarni);
            this.groupBoxRychlePridat.Controls.Add(this.btnRychlyBlokovani);
            this.groupBoxRychlePridat.Location = new System.Drawing.Point(820, 545);
            this.groupBoxRychlePridat.Name = "groupBoxRychlePridat";
            this.groupBoxRychlePridat.Size = new System.Drawing.Size(328, 250);
            this.groupBoxRychlePridat.TabIndex = 4;
            this.groupBoxRychlePridat.TabStop = false;
            this.groupBoxRychlePridat.Text = "Rychlé přidání";
            // 
            // lblSekcePrefixy
            // 
            this.lblSekcePrefixy.AutoSize = true;
            this.lblSekcePrefixy.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblSekcePrefixy.Location = new System.Drawing.Point(10, 20);
            this.lblSekcePrefixy.Name = "lblSekcePrefixy";
            this.lblSekcePrefixy.Size = new System.Drawing.Size(107, 17);
            this.lblSekcePrefixy.TabIndex = 10;
            this.lblSekcePrefixy.Text = "Prefixy značení:";
            // 
            // lblPrefixPTC
            // 
            this.lblPrefixPTC.AutoSize = true;
            this.lblPrefixPTC.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPrefixPTC.Location = new System.Drawing.Point(10, 45);
            this.lblPrefixPTC.Name = "lblPrefixPTC";
            this.lblPrefixPTC.Size = new System.Drawing.Size(30, 15);
            this.lblPrefixPTC.TabIndex = 11;
            this.lblPrefixPTC.Text = "PTC:";
            // 
            // txtPrefixPTC
            // 
            this.txtPrefixPTC.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPrefixPTC.Location = new System.Drawing.Point(45, 42);
            this.txtPrefixPTC.Name = "txtPrefixPTC";
            this.txtPrefixPTC.Size = new System.Drawing.Size(40, 23);
            this.txtPrefixPTC.TabIndex = 12;
            this.txtPrefixPTC.Text = "WH";
            // 
            // lblPrefixOvladani
            // 
            this.lblPrefixOvladani.AutoSize = true;
            this.lblPrefixOvladani.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPrefixOvladani.Location = new System.Drawing.Point(95, 45);
            this.lblPrefixOvladani.Name = "lblPrefixOvladani";
            this.lblPrefixOvladani.Size = new System.Drawing.Size(45, 15);
            this.lblPrefixOvladani.TabIndex = 13;
            this.lblPrefixOvladani.Text = "Ovlád.:";
            // 
            // txtPrefixOvladani
            // 
            this.txtPrefixOvladani.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPrefixOvladani.Location = new System.Drawing.Point(145, 42);
            this.txtPrefixOvladani.Name = "txtPrefixOvladani";
            this.txtPrefixOvladani.Size = new System.Drawing.Size(40, 23);
            this.txtPrefixOvladani.TabIndex = 14;
            this.txtPrefixOvladani.Text = "WS";
            // 
            // lblPrefixUTP
            // 
            this.lblPrefixUTP.AutoSize = true;
            this.lblPrefixUTP.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPrefixUTP.Location = new System.Drawing.Point(195, 45);
            this.lblPrefixUTP.Name = "lblPrefixUTP";
            this.lblPrefixUTP.Size = new System.Drawing.Size(31, 15);
            this.lblPrefixUTP.TabIndex = 15;
            this.lblPrefixUTP.Text = "UTP:";
            // 
            // txtPrefixUTP
            // 
            this.txtPrefixUTP.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPrefixUTP.Location = new System.Drawing.Point(230, 42);
            this.txtPrefixUTP.Name = "txtPrefixUTP";
            this.txtPrefixUTP.Size = new System.Drawing.Size(40, 23);
            this.txtPrefixUTP.TabIndex = 16;
            this.txtPrefixUTP.Text = "WD";
            // 
            // lblPrefixBinarni
            // 
            this.lblPrefixBinarni.AutoSize = true;
            this.lblPrefixBinarni.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPrefixBinarni.Location = new System.Drawing.Point(10, 75);
            this.lblPrefixBinarni.Name = "lblPrefixBinarni";
            this.lblPrefixBinarni.Size = new System.Drawing.Size(30, 15);
            this.lblPrefixBinarni.TabIndex = 17;
            this.lblPrefixBinarni.Text = "Bin.:";
            // 
            // txtPrefixBinarni
            // 
            this.txtPrefixBinarni.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPrefixBinarni.Location = new System.Drawing.Point(45, 72);
            this.txtPrefixBinarni.Name = "txtPrefixBinarni";
            this.txtPrefixBinarni.Size = new System.Drawing.Size(40, 23);
            this.txtPrefixBinarni.TabIndex = 18;
            this.txtPrefixBinarni.Text = "XB";
            // 
            // lblPrefixBlokovani
            // 
            this.lblPrefixBlokovani.AutoSize = true;
            this.lblPrefixBlokovani.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblPrefixBlokovani.Location = new System.Drawing.Point(95, 75);
            this.lblPrefixBlokovani.Name = "lblPrefixBlokovani";
            this.lblPrefixBlokovani.Size = new System.Drawing.Size(36, 15);
            this.lblPrefixBlokovani.TabIndex = 19;
            this.lblPrefixBlokovani.Text = "Blok.:";
            // 
            // txtPrefixBlokovani
            // 
            this.txtPrefixBlokovani.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPrefixBlokovani.Location = new System.Drawing.Point(145, 72);
            this.txtPrefixBlokovani.Name = "txtPrefixBlokovani";
            this.txtPrefixBlokovani.Size = new System.Drawing.Size(40, 23);
            this.txtPrefixBlokovani.TabIndex = 20;
            this.txtPrefixBlokovani.Text = "WB";
            // 
            // btnRychlyPTC
            // 
            this.btnRychlyPTC.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnRychlyPTC.Location = new System.Drawing.Point(10, 110);
            this.btnRychlyPTC.Name = "btnRychlyPTC";
            this.btnRychlyPTC.Size = new System.Drawing.Size(148, 30);
            this.btnRychlyPTC.TabIndex = 0;
            this.btnRychlyPTC.Text = "+ PTC (2 vodiče)";
            this.btnRychlyPTC.UseVisualStyleBackColor = true;
            this.btnRychlyPTC.Click += new System.EventHandler(this.BtnRychlyPTC_Click);
            // 
            // btnRychlyUTP
            // 
            this.btnRychlyUTP.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnRychlyUTP.Location = new System.Drawing.Point(170, 110);
            this.btnRychlyUTP.Name = "btnRychlyUTP";
            this.btnRychlyUTP.Size = new System.Drawing.Size(148, 30);
            this.btnRychlyUTP.TabIndex = 1;
            this.btnRychlyUTP.Text = "+ UTP Cat6";
            this.btnRychlyUTP.UseVisualStyleBackColor = true;
            this.btnRychlyUTP.Click += new System.EventHandler(this.BtnRychlyUTP_Click);
            // 
            // btnRychlyOvladani5
            // 
            this.btnRychlyOvladani5.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnRychlyOvladani5.Location = new System.Drawing.Point(10, 145);
            this.btnRychlyOvladani5.Name = "btnRychlyOvladani5";
            this.btnRychlyOvladani5.Size = new System.Drawing.Size(148, 30);
            this.btnRychlyOvladani5.TabIndex = 2;
            this.btnRychlyOvladani5.Text = "+ Ovládání (5 v.)";
            this.btnRychlyOvladani5.UseVisualStyleBackColor = true;
            this.btnRychlyOvladani5.Click += new System.EventHandler(this.BtnRychlyOvladani5_Click);
            // 
            // btnRychlyOvladani7
            // 
            this.btnRychlyOvladani7.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnRychlyOvladani7.Location = new System.Drawing.Point(170, 145);
            this.btnRychlyOvladani7.Name = "btnRychlyOvladani7";
            this.btnRychlyOvladani7.Size = new System.Drawing.Size(148, 30);
            this.btnRychlyOvladani7.TabIndex = 3;
            this.btnRychlyOvladani7.Text = "+ Ovládání (7 v.)";
            this.btnRychlyOvladani7.UseVisualStyleBackColor = true;
            this.btnRychlyOvladani7.Click += new System.EventHandler(this.BtnRychlyOvladani7_Click);
            // 
            // btnRychlyOvladani12
            // 
            this.btnRychlyOvladani12.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnRychlyOvladani12.Location = new System.Drawing.Point(10, 180);
            this.btnRychlyOvladani12.Name = "btnRychlyOvladani12";
            this.btnRychlyOvladani12.Size = new System.Drawing.Size(148, 30);
            this.btnRychlyOvladani12.TabIndex = 4;
            this.btnRychlyOvladani12.Text = "+ Ovládání (12 v.)";
            this.btnRychlyOvladani12.UseVisualStyleBackColor = true;
            this.btnRychlyOvladani12.Click += new System.EventHandler(this.BtnRychlyOvladani12_Click);
            // 
            // btnRychlyBinarni
            // 
            this.btnRychlyBinarni.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnRychlyBinarni.Location = new System.Drawing.Point(170, 180);
            this.btnRychlyBinarni.Name = "btnRychlyBinarni";
            this.btnRychlyBinarni.Size = new System.Drawing.Size(148, 30);
            this.btnRychlyBinarni.TabIndex = 5;
            this.btnRychlyBinarni.Text = "+ Binární";
            this.btnRychlyBinarni.UseVisualStyleBackColor = true;
            this.btnRychlyBinarni.Click += new System.EventHandler(this.BtnRychlyBinarni_Click);
            // 
            // btnRychlyBlokovani
            // 
            this.btnRychlyBlokovani.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.btnRychlyBlokovani.Location = new System.Drawing.Point(10, 215);
            this.btnRychlyBlokovani.Name = "btnRychlyBlokovani";
            this.btnRychlyBlokovani.Size = new System.Drawing.Size(148, 30);
            this.btnRychlyBlokovani.TabIndex = 6;
            this.btnRychlyBlokovani.Text = "+ Blokování";
            this.btnRychlyBlokovani.UseVisualStyleBackColor = true;
            this.btnRychlyBlokovani.Click += new System.EventHandler(this.BtnRychlyBlokovani_Click);
            // 
            // btnSmazat
            // 
            this.btnSmazat.Location = new System.Drawing.Point(12, 745);
            this.btnSmazat.Name = "btnSmazat";
            this.btnSmazat.Size = new System.Drawing.Size(160, 35);
            this.btnSmazat.TabIndex = 5;
            this.btnSmazat.Text = "Smazat vybraný";
            this.btnSmazat.UseVisualStyleBackColor = true;
            this.btnSmazat.Click += new System.EventHandler(this.BtnSmazat_Click);
            // 
            // lblStatistika
            // 
            this.lblStatistika.AutoSize = true;
            this.lblStatistika.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold);
            this.lblStatistika.Location = new System.Drawing.Point(190, 752);
            this.lblStatistika.Name = "lblStatistika";
            this.lblStatistika.Size = new System.Drawing.Size(183, 20);
            this.lblStatistika.TabIndex = 6;
            this.lblStatistika.Text = "Počet kabelů zařízení: 0";
            // 
            // btnZavrit
            // 
            this.btnZavrit.Location = new System.Drawing.Point(998, 805);
            this.btnZavrit.Name = "btnZavrit";
            this.btnZavrit.Size = new System.Drawing.Size(150, 35);
            this.btnZavrit.TabIndex = 7;
            this.btnZavrit.Text = "Zavřít";
            this.btnZavrit.UseVisualStyleBackColor = true;
            this.btnZavrit.Click += new System.EventHandler(this.BtnZavrit_Click);
            // 
            // FormKabely
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1160, 855);
            this.Controls.Add(this.btnZavrit);
            this.Controls.Add(this.lblStatistika);
            this.Controls.Add(this.btnSmazat);
            this.Controls.Add(this.groupBoxRychlePridat);
            this.Controls.Add(this.groupBoxPridat);
            this.Controls.Add(this.dataGridViewKabely);
            this.Controls.Add(this.groupBoxInfoZarizeni);
            this.Controls.Add(this.groupBoxFiltry);
            this.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormKabely";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Správa kabelů zařízení";
            this.Load += new System.EventHandler(this.FormKabely_Load);
            this.groupBoxFiltry.ResumeLayout(false);
            this.groupBoxFiltry.PerformLayout();
            this.groupBoxInfoZarizeni.ResumeLayout(false);
            this.groupBoxInfoZarizeni.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewKabely)).EndInit();
            this.groupBoxPridat.ResumeLayout(false);
            this.groupBoxPridat.PerformLayout();
            this.groupBoxRychlePridat.ResumeLayout(false);
            this.groupBoxRychlePridat.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
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
        private System.Windows.Forms.Label lblFilterExist;
        private System.Windows.Forms.ComboBox comboBoxFilterExist;
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
    }
}
