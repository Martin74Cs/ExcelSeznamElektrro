using System.Threading.Tasks;

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
        private async Task InitializeComponent()
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
            button8 = new Button();
            button9 = new Button();
            label7 = new Label();
            menuStrip1 = new MenuStrip();
            souborToolStripMenuItem = new ToolStripMenuItem();
            openToolStripMenuItem = new ToolStripMenuItem();
            seznamyToolStripMenuItem = new ToolStripMenuItem();
            místnostiToolStripMenuItem = new ToolStripMenuItem();
            místnostiToolStripMenuItem1 = new ToolStripMenuItem();
            generovatToolStripMenuItem = new ToolStripMenuItem();
            pomocToolStripMenuItem = new ToolStripMenuItem();
            button6 = new Button();
            label5 = new Label();
            button7 = new Button();
            label6 = new Label();
            button10 = new Button();
            label8 = new Label();
            button11 = new Button();
            label9 = new Label();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button1.Font = new Font("Segoe UI", 12F);
            button1.Location = new Point(975, 618);
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
            label1.Location = new Point(882, 41);
            label1.Name = "label1";
            label1.Size = new Size(177, 21);
            label1.TabIndex = 1;
            label1.Text = "Převod Strojů na Elektro";
            // 
            // textBox1
            // 
            textBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            textBox1.Font = new Font("Segoe UI", 12F);
            textBox1.Location = new Point(15, 627);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(688, 29);
            textBox1.TabIndex = 2;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button2.Font = new Font("Segoe UI", 12F);
            button2.Location = new Point(709, 32);
            button2.Name = "button2";
            button2.Size = new Size(167, 38);
            button2.TabIndex = 3;
            button2.Text = "Převod->json,csv";
            button2.UseVisualStyleBackColor = true;
            button2.Click += Button2_Click;
            // 
            // listBox1
            // 
            listBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listBox1.FormattingEnabled = true;
            listBox1.Location = new Point(12, 32);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(691, 574);
            listBox1.TabIndex = 4;
            listBox1.SelectedIndexChanged += ListBox1_SelectedIndexChanged_1;
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button3.Font = new Font("Segoe UI", 12F);
            button3.Location = new Point(709, 167);
            button3.Name = "button3";
            button3.Size = new Size(167, 38);
            button3.TabIndex = 6;
            button3.Text = "Přidat Csv -> Json";
            button3.UseVisualStyleBackColor = true;
            button3.Click += Button3_Click;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 12F);
            label2.Location = new Point(882, 176);
            label2.Name = "label2";
            label2.Size = new Size(162, 21);
            label2.TabIndex = 5;
            label2.Text = "Doplnění  Dat do Json";
            // 
            // button4
            // 
            button4.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button4.Font = new Font("Segoe UI", 12F);
            button4.Location = new Point(709, 618);
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
            label3.Location = new Point(882, 627);
            label3.Name = "label3";
            label3.Size = new Size(86, 21);
            label3.TabIndex = 7;
            label3.Text = "Excel Close";
            // 
            // button5
            // 
            button5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button5.Font = new Font("Segoe UI", 12F);
            button5.Location = new Point(709, 251);
            button5.Name = "button5";
            button5.Size = new Size(167, 38);
            button5.TabIndex = 10;
            button5.Text = "Přidat Kabely";
            button5.UseVisualStyleBackColor = true;
            button5.Click += Button5_Click;
            // 
            // label4
            // 
            label4.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 12F);
            label4.Location = new Point(882, 260);
            label4.Name = "label4";
            label4.Size = new Size(141, 21);
            label4.TabIndex = 9;
            label4.Text = "Dle proudu a délky";
            // 
            // button8
            // 
            button8.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button8.Font = new Font("Segoe UI", 12F);
            button8.Location = new Point(709, 123);
            button8.Name = "button8";
            button8.Size = new Size(167, 38);
            button8.TabIndex = 16;
            button8.Text = "Otevřít Csv";
            button8.UseVisualStyleBackColor = true;
            button8.Click += Button8_Click;
            // 
            // button9
            // 
            button9.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button9.Font = new Font("Segoe UI", 12F);
            button9.Location = new Point(709, 339);
            button9.Name = "button9";
            button9.Size = new Size(167, 38);
            button9.TabIndex = 18;
            button9.Text = "Přidat Vývody.csv ";
            button9.UseVisualStyleBackColor = true;
            button9.Click += Button9_Click;
            // 
            // label7
            // 
            label7.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label7.AutoSize = true;
            label7.Font = new Font("Segoe UI", 12F);
            label7.Location = new Point(882, 304);
            label7.Name = "label7";
            label7.Size = new Size(149, 21);
            label7.TabIndex = 17;
            label7.Text = "Přidat vlasní vývody";
            // 
            // menuStrip1
            // 
            menuStrip1.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 238);
            menuStrip1.Items.AddRange(new ToolStripItem[] { souborToolStripMenuItem, seznamyToolStripMenuItem, místnostiToolStripMenuItem, pomocToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(1073, 29);
            menuStrip1.TabIndex = 19;
            menuStrip1.Text = "menuStrip1";
            // 
            // souborToolStripMenuItem
            // 
            souborToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { openToolStripMenuItem });
            souborToolStripMenuItem.Name = "souborToolStripMenuItem";
            souborToolStripMenuItem.Size = new Size(73, 25);
            souborToolStripMenuItem.Text = "Soubor";
            // 
            // openToolStripMenuItem
            // 
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.Size = new Size(176, 26);
            openToolStripMenuItem.Text = "Otevřít složku";
            openToolStripMenuItem.Click += OpenToolStripMenuItem_Click;
            // 
            // seznamyToolStripMenuItem
            // 
            seznamyToolStripMenuItem.Name = "seznamyToolStripMenuItem";
            seznamyToolStripMenuItem.Size = new Size(85, 25);
            seznamyToolStripMenuItem.Text = "Seznamy";
            seznamyToolStripMenuItem.Click += SeznamyToolStripMenuItem_Click;
            // 
            // místnostiToolStripMenuItem
            // 
            místnostiToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { místnostiToolStripMenuItem1, generovatToolStripMenuItem });
            místnostiToolStripMenuItem.Name = "místnostiToolStripMenuItem";
            místnostiToolStripMenuItem.Size = new Size(86, 25);
            místnostiToolStripMenuItem.Text = "Místnosti";
            // 
            // místnostiToolStripMenuItem1
            // 
            místnostiToolStripMenuItem1.Name = "místnostiToolStripMenuItem1";
            místnostiToolStripMenuItem1.Size = new Size(180, 26);
            místnostiToolStripMenuItem1.Text = "Otevřít složku";
            místnostiToolStripMenuItem1.Click += MístnostiToolStripMenuItem1_Click;
            // 
            // generovatToolStripMenuItem
            // 
            generovatToolStripMenuItem.Name = "generovatToolStripMenuItem";
            generovatToolStripMenuItem.Size = new Size(180, 26);
            generovatToolStripMenuItem.Text = "Generovat ";
            generovatToolStripMenuItem.Click += GenerovatToolStripMenuItem_Click;
            // 
            // pomocToolStripMenuItem
            // 
            pomocToolStripMenuItem.Name = "pomocToolStripMenuItem";
            pomocToolStripMenuItem.Size = new Size(69, 25);
            pomocToolStripMenuItem.Text = "Pomoc";
            // 
            // button6
            // 
            button6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button6.Font = new Font("Segoe UI", 12F);
            button6.Location = new Point(709, 387);
            button6.Name = "button6";
            button6.Size = new Size(167, 38);
            button6.TabIndex = 20;
            button6.Text = "Vytvořit Json->Excel";
            button6.UseVisualStyleBackColor = true;
            button6.Click += Button6_Click;
            // 
            // label5
            // 
            label5.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI", 12F);
            label5.Location = new Point(882, 132);
            label5.Name = "label5";
            label5.Size = new Size(175, 21);
            label5.TabIndex = 21;
            label5.Text = "Úprava dat, kopie strojů";
            // 
            // button7
            // 
            button7.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button7.Font = new Font("Segoe UI", 12F);
            button7.Location = new Point(709, 211);
            button7.Name = "button7";
            button7.Size = new Size(167, 38);
            button7.TabIndex = 22;
            button7.Text = "Přidat Proud ";
            button7.UseVisualStyleBackColor = true;
            button7.Click += Button7_Click;
            // 
            // label6
            // 
            label6.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 12F);
            label6.Location = new Point(882, 220);
            label6.Name = "label6";
            label6.Size = new Size(182, 21);
            label6.TabIndex = 23;
            label6.Text = "Dle příkonu přidat proud";
            // 
            // button10
            // 
            button10.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button10.Font = new Font("Segoe UI", 12F);
            button10.Location = new Point(709, 79);
            button10.Name = "button10";
            button10.Size = new Size(167, 38);
            button10.TabIndex = 24;
            button10.Text = "Otevřít hlavni Json";
            button10.UseVisualStyleBackColor = true;
            button10.Click += Button10_Click;
            // 
            // label8
            // 
            label8.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label8.AutoSize = true;
            label8.Font = new Font("Segoe UI", 12F);
            label8.Location = new Point(886, 88);
            label8.Name = "label8";
            label8.Size = new Size(102, 21);
            label8.TabIndex = 25;
            label8.Text = "Aktualní Json";
            // 
            // button11
            // 
            button11.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            button11.Font = new Font("Segoe UI", 12F);
            button11.Location = new Point(709, 295);
            button11.Name = "button11";
            button11.Size = new Size(167, 38);
            button11.TabIndex = 26;
            button11.Text = "Otevřít Vývody.csv";
            button11.UseVisualStyleBackColor = true;
            button11.Click += Button11_Click;
            // 
            // label9
            // 
            label9.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            label9.AutoSize = true;
            label9.Font = new Font("Segoe UI", 12F);
            label9.Location = new Point(882, 348);
            label9.Name = "label9";
            label9.Size = new Size(172, 21);
            label9.TabIndex = 27;
            label9.Text = "Přidat do aktualno Json";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1073, 668);
            Controls.Add(label9);
            Controls.Add(button11);
            Controls.Add(label8);
            Controls.Add(button10);
            Controls.Add(label6);
            Controls.Add(button7);
            Controls.Add(label5);
            Controls.Add(button6);
            Controls.Add(button9);
            Controls.Add(label7);
            Controls.Add(button8);
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
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
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
        private Button button8;
        private Button button9;
        private Label label7;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem souborToolStripMenuItem;
        private ToolStripMenuItem seznamyToolStripMenuItem;
        private ToolStripMenuItem pomocToolStripMenuItem;
        private ToolStripMenuItem openToolStripMenuItem;
        private ToolStripMenuItem místnostiToolStripMenuItem;
        private ToolStripMenuItem místnostiToolStripMenuItem1;
        private ToolStripMenuItem generovatToolStripMenuItem;
        private Button button6;
        private Label label5;
        private Button button7;
        private Label label6;
        private Button button10;
        private Label label8;
        private Button button11;
        private Label label9;
    }
}
