using Aplikace.Sdilene;
using Aplikace.Upravy;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinForms
{
    public partial class Vytvořit : Form
    {
        public Vytvořit()
        {
            InitializeComponent();
        }

        private void Vytvořit_Load(object sender, EventArgs e)
        {
            //Console.SetOut(new ListBoxWriter(listBox1));
        }

        //public class ListBoxWriter(ListBox listBox) : TextWriter
        //{
        //    private readonly ListBox _listBox = listBox;
        //    private readonly SynchronizationContext _context = SynchronizationContext.Current;

        //    public override Encoding Encoding => Encoding.UTF8;

        //    public override void WriteLine(string value)
        //    {
        //        //_context.Post(_ => _listBox.Items.Add(value), null);
        //        _context.Post(_ =>
        //        {
        //            _listBox.Items.Add(value);
        //            _listBox.TopIndex = _listBox.Items.Count - 1; // ← automatické scrollování dolů
        //        }, null);
        //    }
        //}

        private void Button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            //Close();
        }

        private async void Button7_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.VyvoritMotor());
        }

        private async void Button6_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.VyvoritFMKM());
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            //Motory open
            string cesta = Cesty.BasePath;
            cesta = System.IO.Path.Combine(cesta, "Data", "Motory");
            System.Diagnostics.Process.Start("explorer.exe", cesta);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Frekvenční měniče open
            string cesta = Cesty.BasePath;
            cesta = System.IO.Path.Combine(cesta, "Data", "FM.csv");
            System.Diagnostics.Process.Start("explorer.exe", cesta);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Stykače open
            string cesta = Cesty.BasePath;
            cesta = System.IO.Path.Combine(cesta, "Data", "KM.csv");
            System.Diagnostics.Process.Start("explorer.exe", cesta);
        }

        private async void button5_Click(object sender, EventArgs e)
        {
            //Stykače prevo
            await Task.Run(() => LigthChem.VyvoritFMKM());
        }
    }
}
