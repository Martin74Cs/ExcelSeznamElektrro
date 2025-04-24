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
        private void InitializeComponent()
        {
            button1 = new Button();
            button7 = new Button();
            label6 = new Label();
            button6 = new Button();
            label5 = new Label();
            listBox1 = new ListBox();
            button2 = new Button();
            label1 = new Label();
            button3 = new Button();
            button4 = new Button();
            button5 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button1.Font = new Font("Segoe UI", 12F);
            button1.Location = new Point(740, 416);
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
            button7.Location = new Point(548, 12);
            button7.Name = "button7";
            button7.Size = new Size(167, 38);
            button7.TabIndex = 20;
            button7.Text = "Vytvořit Motory";
            button7.UseVisualStyleBackColor = true;
            button7.Click += Button7_Click;
            // 
            // label6
            // 
            label6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 12F);
            label6.Location = new Point(721, 21);
            label6.Name = "label6";
            label6.Size = new Size(92, 21);
            label6.TabIndex = 19;
            label6.Text = "CSV to Json";
            // 
            // button6
            // 
            button6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button6.Font = new Font("Segoe UI", 12F);
            button6.Location = new Point(548, 55);
            button6.Name = "button6";
            button6.Size = new Size(167, 38);
            button6.TabIndex = 18;
            button6.Text = "Vytvořit FM";
            button6.UseVisualStyleBackColor = true;
            button6.Click += Button6_Click;
            // 
            // label5
            // 
            label5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI", 12F);
            label5.Location = new Point(721, 64);
            label5.Name = "label5";
            label5.Size = new Size(92, 21);
            label5.TabIndex = 17;
            label5.Text = "CSV to Json";
            // 
            // listBox1
            // 
            listBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listBox1.FormattingEnabled = true;
            listBox1.Location = new Point(12, 12);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(357, 394);
            listBox1.TabIndex = 21;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button2.Font = new Font("Segoe UI", 12F);
            button2.Location = new Point(375, 12);
            button2.Name = "button2";
            button2.Size = new Size(167, 38);
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
            label1.Location = new Point(721, 109);
            label1.Name = "label1";
            label1.Size = new Size(92, 21);
            label1.TabIndex = 23;
            label1.Text = "CSV to Json";
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button3.Font = new Font("Segoe UI", 12F);
            button3.Location = new Point(375, 55);
            button3.Name = "button3";
            button3.Size = new Size(167, 38);
            button3.TabIndex = 24;
            button3.Text = "Otevřít FM";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button4.Font = new Font("Segoe UI", 12F);
            button4.Location = new Point(375, 100);
            button4.Name = "button4";
            button4.Size = new Size(167, 38);
            button4.TabIndex = 25;
            button4.Text = "Otevřít KM";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // button5
            // 
            button5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button5.Font = new Font("Segoe UI", 12F);
            button5.Location = new Point(548, 100);
            button5.Name = "button5";
            button5.Size = new Size(167, 38);
            button5.TabIndex = 26;
            button5.Text = "Vytvořit KM";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // Vytvořit
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(825, 461);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(label1);
            Controls.Add(button2);
            Controls.Add(listBox1);
            Controls.Add(button7);
            Controls.Add(label6);
            Controls.Add(button6);
            Controls.Add(label5);
            Controls.Add(button1);
            Name = "Vytvořit";
            Text = "Vytvořit";
            Load += Vytvořit_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button7;
        private Label label6;
        private Button button6;
        private Label label5;
        private ListBox listBox1;
        private Button button2;
        private Label label1;
        private Button button3;
        private Button button4;
        private Button button5;
    }
}