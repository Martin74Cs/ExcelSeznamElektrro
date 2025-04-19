namespace WinForms
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            label1 = new Label();
            textBox1 = new TextBox();
            button2 = new Button();
            listBox1 = new ListBox();
            button3 = new Button();
            label2 = new Label();
            button4 = new Button();
            label3 = new Label();
            button5 = new Button();
            label4 = new Label();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button1.Font = new Font("Segoe UI", 12F);
            button1.Location = new Point(879, 398);
            button1.Name = "button1";
            button1.Size = new Size(86, 38);
            button1.TabIndex = 0;
            button1.Text = "Konec";
            button1.UseVisualStyleBackColor = true;
            button1.Click += Button1_Click;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 12F);
            label1.Location = new Point(788, 21);
            label1.Name = "label1";
            label1.Size = new Size(177, 21);
            label1.TabIndex = 1;
            label1.Text = "Převod Strojů na Elektro";
            // 
            // textBox1
            // 
            textBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            textBox1.Font = new Font("Segoe UI", 12F);
            textBox1.Location = new Point(15, 407);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(592, 29);
            textBox1.TabIndex = 2;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button2.Font = new Font("Segoe UI", 12F);
            button2.Location = new Point(615, 12);
            button2.Name = "button2";
            button2.Size = new Size(167, 38);
            button2.TabIndex = 3;
            button2.Text = "Převod";
            button2.UseVisualStyleBackColor = true;
            button2.Click += Button2_Click;
            // 
            // listBox1
            // 
            listBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listBox1.FormattingEnabled = true;
            listBox1.Location = new Point(12, 17);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(595, 379);
            listBox1.TabIndex = 4;
            listBox1.SelectedIndexChanged += ListBox1_SelectedIndexChanged_1;
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button3.Font = new Font("Segoe UI", 12F);
            button3.Location = new Point(615, 56);
            button3.Name = "button3";
            button3.Size = new Size(167, 38);
            button3.TabIndex = 6;
            button3.Text = "Doplň Csv -> Json";
            button3.UseVisualStyleBackColor = true;
            button3.Click += Button3_Click;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 12F);
            label2.Location = new Point(788, 65);
            label2.Name = "label2";
            label2.Size = new Size(177, 21);
            label2.TabIndex = 5;
            label2.Text = "Převod Strojů na Elektro";
            // 
            // button4
            // 
            button4.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button4.Font = new Font("Segoe UI", 12F);
            button4.Location = new Point(613, 347);
            button4.Name = "button4";
            button4.Size = new Size(167, 38);
            button4.TabIndex = 8;
            button4.Text = "Kill Excel";
            button4.UseVisualStyleBackColor = true;
            button4.Click += Button4_Click;
            // 
            // label3
            // 
            label3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 12F);
            label3.Location = new Point(786, 356);
            label3.Name = "label3";
            label3.Size = new Size(86, 21);
            label3.TabIndex = 7;
            label3.Text = "Excel Close";
            // 
            // button5
            // 
            button5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button5.Font = new Font("Segoe UI", 12F);
            button5.Location = new Point(615, 184);
            button5.Name = "button5";
            button5.Size = new Size(167, 38);
            button5.TabIndex = 10;
            button5.Text = "Přidat Kabely";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // label4
            // 
            label4.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 12F);
            label4.Location = new Point(788, 193);
            label4.Name = "label4";
            label4.Size = new Size(139, 21);
            label4.TabIndex = 9;
            label4.Text = "dle proudu a délky";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(977, 448);
            Controls.Add(button5);
            Controls.Add(label4);
            Controls.Add(button4);
            Controls.Add(label3);
            Controls.Add(button3);
            Controls.Add(label2);
            Controls.Add(listBox1);
            Controls.Add(button2);
            Controls.Add(textBox1);
            Controls.Add(label1);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Label label1;
        private TextBox textBox1;
        private Button button2;
        private ListBox listBox1;
        private Button button3;
        private Label label2;
        private Button button4;
        private Label label3;
        private Button button5;
        private Label label4;
    }
}
