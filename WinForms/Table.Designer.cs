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
            Button1 = new Button();
            Button2 = new Button();
            button3 = new Button();
            Button4 = new Button();
            button5 = new Button();
            button6 = new Button();
            label1 = new Label();
            comboBox1 = new ComboBox();
            BtnAdd = new Button();
            button7 = new Button();
            comboBox2 = new ComboBox();
            comboBox3 = new ComboBox();
            comboBox4Pid = new ComboBox();
            button8 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(15, 59);
            dataGridView1.Margin = new Padding(4);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(1100, 627);
            dataGridView1.TabIndex = 0;
            dataGridView1.CellContentClick += DataGridView1_CellContentClick;
            dataGridView1.CellFormatting += DataGridView1_CellFormatting;
            dataGridView1.CellMouseUp += DataGridView1_CellMouseUp;
            dataGridView1.CurrentCellChanged += DataGridView1_CurrentCellChanged;
            dataGridView1.RowsAdded += DataGridView1_RowsAdded;
            // 
            // Button1
            // 
            Button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Button1.Location = new Point(1123, 654);
            Button1.Margin = new Padding(4);
            Button1.Name = "Button1";
            Button1.Size = new Size(96, 32);
            Button1.TabIndex = 1;
            Button1.Text = "Konec";
            Button1.UseVisualStyleBackColor = true;
            Button1.Click += Button1_Click;
            // 
            // Button2
            // 
            Button2.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Button2.Location = new Point(1123, 614);
            Button2.Margin = new Padding(4);
            Button2.Name = "Button2";
            Button2.Size = new Size(96, 32);
            Button2.TabIndex = 2;
            Button2.Text = "Save";
            Button2.UseVisualStyleBackColor = true;
            Button2.Click += Button2_Click;
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button3.Location = new Point(1123, 408);
            button3.Margin = new Padding(4);
            button3.Name = "button3";
            button3.Size = new Size(96, 32);
            button3.TabIndex = 3;
            button3.Text = "Proud";
            button3.UseVisualStyleBackColor = true;
            button3.Click += Button3_Click;
            // 
            // Button4
            // 
            Button4.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Button4.Location = new Point(1123, 448);
            Button4.Margin = new Padding(4);
            Button4.Name = "Button4";
            Button4.Size = new Size(96, 32);
            Button4.TabIndex = 4;
            Button4.Text = "Průřez";
            Button4.UseVisualStyleBackColor = true;
            Button4.Click += Button4_Click;
            // 
            // button5
            // 
            button5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button5.Location = new Point(1123, 99);
            button5.Margin = new Padding(4);
            button5.Name = "button5";
            button5.Size = new Size(96, 32);
            button5.TabIndex = 5;
            button5.Text = "Rozvaděč";
            button5.UseVisualStyleBackColor = true;
            button5.Click += Button5_Click;
            // 
            // button6
            // 
            button6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button6.Location = new Point(1123, 59);
            button6.Margin = new Padding(4);
            button6.Name = "button6";
            button6.Size = new Size(96, 32);
            button6.TabIndex = 6;
            button6.Text = "Vše";
            button6.UseVisualStyleBackColor = true;
            button6.Click += Button6_Click;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            label1.AutoSize = true;
            label1.Location = new Point(1123, 383);
            label1.Name = "label1";
            label1.Size = new Size(73, 21);
            label1.TabIndex = 7;
            label1.Text = "Výpočet :";
            // 
            // comboBox1
            // 
            comboBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(1122, 497);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(96, 29);
            comboBox1.TabIndex = 8;
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;
            comboBox1.MouseClick += ComboBox1_MouseClick;
            // 
            // BtnAdd
            // 
            BtnAdd.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            BtnAdd.Location = new Point(1122, 533);
            BtnAdd.Margin = new Padding(4);
            BtnAdd.Name = "BtnAdd";
            BtnAdd.Size = new Size(96, 32);
            BtnAdd.TabIndex = 9;
            BtnAdd.Text = "Přidat";
            BtnAdd.UseVisualStyleBackColor = true;
            BtnAdd.Click += BtnAdd_Click;
            // 
            // button7
            // 
            button7.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button7.Location = new Point(1122, 574);
            button7.Margin = new Padding(4);
            button7.Name = "button7";
            button7.Size = new Size(96, 32);
            button7.TabIndex = 10;
            button7.Text = "Delete";
            button7.UseVisualStyleBackColor = true;
            button7.Click += Button7_Click;
            // 
            // comboBox2
            // 
            comboBox2.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            comboBox2.FormattingEnabled = true;
            comboBox2.Location = new Point(1122, 254);
            comboBox2.Name = "comboBox2";
            comboBox2.Size = new Size(96, 29);
            comboBox2.TabIndex = 11;
            comboBox2.SelectedIndexChanged += ComboBox2_SelectedIndexChanged;
            comboBox2.MouseClick += ComboBox2_MouseClick;
            // 
            // comboBox3
            // 
            comboBox3.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            comboBox3.FormattingEnabled = true;
            comboBox3.Location = new Point(1123, 318);
            comboBox3.Name = "comboBox3";
            comboBox3.Size = new Size(96, 29);
            comboBox3.TabIndex = 12;
            comboBox3.SelectedIndexChanged += ComboBox3_SelectedIndexChanged;
            comboBox3.MouseClick += ComboBox3_MouseClick;
            // 
            // comboBox4Pid
            // 
            comboBox4Pid.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            comboBox4Pid.FormattingEnabled = true;
            comboBox4Pid.Location = new Point(1123, 209);
            comboBox4Pid.Name = "comboBox4Pid";
            comboBox4Pid.Size = new Size(96, 29);
            comboBox4Pid.TabIndex = 13;
            comboBox4Pid.SelectedIndexChanged += ComboBox4Pid_SelectedIndexChanged;
            comboBox4Pid.MouseClick += ComboBox4Pid_MouseClick;
            // 
            // button8
            // 
            button8.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button8.Location = new Point(1123, 139);
            button8.Margin = new Padding(4);
            button8.Name = "button8";
            button8.Size = new Size(96, 32);
            button8.TabIndex = 14;
            button8.Text = "Data";
            button8.UseVisualStyleBackColor = true;
            button8.Click += Button8_Click;
            // 
            // Table
            // 
            AutoScaleDimensions = new SizeF(9F, 21F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1242, 697);
            Controls.Add(button8);
            Controls.Add(comboBox4Pid);
            Controls.Add(comboBox3);
            Controls.Add(comboBox2);
            Controls.Add(button7);
            Controls.Add(BtnAdd);
            Controls.Add(comboBox1);
            Controls.Add(label1);
            Controls.Add(button6);
            Controls.Add(button5);
            Controls.Add(Button4);
            Controls.Add(button3);
            Controls.Add(Button2);
            Controls.Add(Button1);
            Controls.Add(dataGridView1);
            Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 238);
            Margin = new Padding(4);
            Name = "Table";
            Text = "Table";
            FormClosing += Table_FormClosing;
            Load += Table_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public DataGridView dataGridView1;
        private Button Button1;
        private Button Button2;
        private Button button3;
        private Button Button4;
        private Button button5;
        private Button button6;
        private Label label1;
        private ComboBox comboBox1;
        private Button BtnAdd;
        private Button button7;
        private ComboBox comboBox2;
        private ComboBox comboBox3;
        private ComboBox comboBox4Pid;
        private Button button8;
    }
}