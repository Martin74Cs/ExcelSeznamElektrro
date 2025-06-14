namespace WinForms
{
    partial class Vytvořit
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
            button1 = new Button();
            button7 = new Button();
            label6 = new Label();
            button6 = new Button();
            label5 = new Label();
            button2 = new Button();
            label1 = new Label();
            button3 = new Button();
            button4 = new Button();
            button5 = new Button();
            dataGridView1 = new DataGridView();
            Button8 = new Button();
            Button9 = new Button();
            Button10 = new Button();
            label2 = new Label();
            button11 = new Button();
            button12 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button1.Font = new Font("Segoe UI", 12F);
            button1.Location = new Point(918, 553);
            button1.Name = "button1";
            button1.Size = new Size(75, 36);
            button1.TabIndex = 0;
            button1.Text = "Zavřít";
            button1.UseVisualStyleBackColor = true;
            button1.Click += Button1_Click;
            // 
            // button7
            // 
            button7.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button7.Font = new Font("Segoe UI", 12F);
            button7.Location = new Point(139, 458);
            button7.Name = "button7";
            button7.Size = new Size(159, 38);
            button7.TabIndex = 20;
            button7.Text = "Open table Motory";
            button7.UseVisualStyleBackColor = true;
            button7.Click += Button7_Click;
            // 
            // label6
            // 
            label6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 12F);
            label6.Location = new Point(396, 468);
            label6.Name = "label6";
            label6.Size = new Size(92, 21);
            label6.TabIndex = 19;
            label6.Text = "CSV to Json";
            // 
            // button6
            // 
            button6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button6.Font = new Font("Segoe UI", 12F);
            button6.Location = new Point(139, 501);
            button6.Name = "button6";
            button6.Size = new Size(159, 38);
            button6.TabIndex = 18;
            button6.Text = "Open table FM";
            button6.UseVisualStyleBackColor = true;
            button6.Click += Button6_Click;
            // 
            // label5
            // 
            label5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI", 12F);
            label5.Location = new Point(396, 511);
            label5.Name = "label5";
            label5.Size = new Size(92, 21);
            label5.TabIndex = 17;
            label5.Text = "CSV to Json";
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button2.Font = new Font("Segoe UI", 12F);
            button2.Location = new Point(12, 458);
            button2.Name = "button2";
            button2.Size = new Size(121, 38);
            button2.TabIndex = 22;
            button2.Text = "Otevřít Motory";
            button2.UseVisualStyleBackColor = true;
            button2.Click += Button2_Click;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 12F);
            label1.Location = new Point(396, 556);
            label1.Name = "label1";
            label1.Size = new Size(92, 21);
            label1.TabIndex = 23;
            label1.Text = "CSV to Json";
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button3.Font = new Font("Segoe UI", 12F);
            button3.Location = new Point(12, 501);
            button3.Name = "button3";
            button3.Size = new Size(121, 38);
            button3.TabIndex = 24;
            button3.Text = "Otevřít FM";
            button3.UseVisualStyleBackColor = true;
            button3.Click += Button3_Click;
            // 
            // button4
            // 
            button4.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button4.Font = new Font("Segoe UI", 12F);
            button4.Location = new Point(12, 546);
            button4.Name = "button4";
            button4.Size = new Size(121, 38);
            button4.TabIndex = 25;
            button4.Text = "Otevřít KM";
            button4.UseVisualStyleBackColor = true;
            button4.Click += Button4_Click;
            // 
            // button5
            // 
            button5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button5.Font = new Font("Segoe UI", 12F);
            button5.Location = new Point(139, 546);
            button5.Name = "button5";
            button5.Size = new Size(159, 38);
            button5.TabIndex = 26;
            button5.Text = "Open table KM";
            button5.UseVisualStyleBackColor = true;
            button5.Click += Button5_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(14, 31);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(979, 385);
            dataGridView1.TabIndex = 27;
            // 
            // Button8
            // 
            Button8.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Button8.Font = new Font("Segoe UI", 12F);
            Button8.Location = new Point(304, 547);
            Button8.Name = "Button8";
            Button8.Size = new Size(89, 38);
            Button8.TabIndex = 28;
            Button8.Text = "Save KM";
            Button8.UseVisualStyleBackColor = true;
            Button8.Click += Button8_Click;
            // 
            // Button9
            // 
            Button9.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Button9.Font = new Font("Segoe UI", 12F);
            Button9.Location = new Point(304, 502);
            Button9.Name = "Button9";
            Button9.Size = new Size(89, 38);
            Button9.TabIndex = 29;
            Button9.Text = "Save FM";
            Button9.UseVisualStyleBackColor = true;
            Button9.Click += Button9_Click;
            // 
            // Button10
            // 
            Button10.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Button10.Font = new Font("Segoe UI", 12F);
            Button10.Location = new Point(304, 458);
            Button10.Name = "Button10";
            Button10.Size = new Size(89, 38);
            Button10.TabIndex = 30;
            Button10.Text = "Save M";
            Button10.UseVisualStyleBackColor = true;
            Button10.Click += Button10_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(14, 431);
            label2.Name = "label2";
            label2.Size = new Size(10, 15);
            label2.TabIndex = 31;
            label2.Text = ".";
            // 
            // button11
            // 
            button11.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button11.Font = new Font("Segoe UI", 12F);
            button11.Location = new Point(529, 451);
            button11.Name = "button11";
            button11.Size = new Size(128, 38);
            button11.TabIndex = 32;
            button11.Text = "Open OEZ 3VA";
            button11.UseVisualStyleBackColor = true;
            button11.Click += Button11_Click;
            // 
            // button12
            // 
            button12.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button12.Font = new Font("Segoe UI", 12F);
            button12.Location = new Point(663, 451);
            button12.Name = "button12";
            button12.Size = new Size(89, 38);
            button12.TabIndex = 33;
            button12.Text = "Save M";
            button12.UseVisualStyleBackColor = true;
            button12.Click += Button12_Click;
            // 
            // Vytvořit
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1003, 598);
            Controls.Add(button12);
            Controls.Add(button11);
            Controls.Add(label2);
            Controls.Add(Button10);
            Controls.Add(Button9);
            Controls.Add(Button8);
            Controls.Add(dataGridView1);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(label1);
            Controls.Add(button2);
            Controls.Add(button7);
            Controls.Add(label6);
            Controls.Add(button6);
            Controls.Add(label5);
            Controls.Add(button1);
            Name = "Vytvořit";
            Text = "Vytvořit";
            FormClosing += Vytvořit_FormClosing;
            Load += Vytvořit_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button7;
        private Label label6;
        private Button button6;
        private Label label5;
        private Button button2;
        private Label label1;
        private Button button3;
        private Button button4;
        private Button button5;
        private DataGridView dataGridView1;
        private Button Button8;
        private Button Button9;
        private Button Button10;
        private Label label2;
        private Button button11;
        private Button button12;
    }
}