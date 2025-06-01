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
        private void InitializeComponent() {
            dataGridView1 = new DataGridView();
            Button1 = new Button();
            Button2 = new Button();
            button3 = new Button();
            Button4 = new Button();
            button5 = new Button();
            button6 = new Button();
            label1 = new Label();
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
            dataGridView1.Size = new Size(1077, 613);
            dataGridView1.TabIndex = 0;
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            dataGridView1.CellMouseUp += dataGridView1_CellMouseUp;
            dataGridView1.CurrentCellChanged += dataGridView1_CurrentCellChanged;
            // 
            // Button1
            // 
            Button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Button1.Location = new Point(1100, 640);
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
            Button2.Location = new Point(1100, 600);
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
            button3.Location = new Point(1100, 400);
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
            Button4.Location = new Point(1100, 440);
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
            button5.Location = new Point(1100, 101);
            button5.Margin = new Padding(4);
            button5.Name = "button5";
            button5.Size = new Size(96, 32);
            button5.TabIndex = 5;
            button5.Text = "Rozvaděč";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // button6
            // 
            button6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button6.Location = new Point(1100, 59);
            button6.Margin = new Padding(4);
            button6.Name = "button6";
            button6.Size = new Size(96, 32);
            button6.TabIndex = 6;
            button6.Text = "Vše";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(1100, 375);
            label1.Name = "label1";
            label1.Size = new Size(73, 21);
            label1.TabIndex = 7;
            label1.Text = "Výpočet :";
            // 
            // Table
            // 
            AutoScaleDimensions = new SizeF(9F, 21F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1219, 683);
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
    }
}