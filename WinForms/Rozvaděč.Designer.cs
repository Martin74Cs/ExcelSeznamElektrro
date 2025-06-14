namespace WinForms {
    partial class Rozvaděč {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
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
            listViewCategories = new ListView();
            listViewProducts = new ListView();
            SuspendLayout();
            // 
            // listViewCategories
            // 
            listViewCategories.Location = new Point(12, 12);
            listViewCategories.Name = "listViewCategories";
            listViewCategories.Size = new Size(331, 383);
            listViewCategories.TabIndex = 2;
            listViewCategories.UseCompatibleStateImageBehavior = false;
            listViewCategories.SelectedIndexChanged += ListViewCategories_SelectedIndexChanged;
            // 
            // listViewProducts
            // 
            listViewProducts.Location = new Point(376, 12);
            listViewProducts.Name = "listViewProducts";
            listViewProducts.Size = new Size(353, 383);
            listViewProducts.TabIndex = 3;
            listViewProducts.UseCompatibleStateImageBehavior = false;
            // 
            // Rozvaděč
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(listViewProducts);
            Controls.Add(listViewCategories);
            Name = "Rozvaděč";
            Text = "Rozvaděč";
            Load += Rozvaděč_Load;
            ResumeLayout(false);
        }

        #endregion

        private ListView listViewCategories;
        private ListView listViewProducts;
    }
}