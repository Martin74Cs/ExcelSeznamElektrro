using Aplikace.Sdilene;
using Aplikace.Upravy;
using System.Text;
using System.Windows.Forms;

namespace WinForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Převod
        private async void Button2_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.Hlavni());
            //Console.SetOut(new ListBoxWriter(listBox1));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.SetOut(new ListBoxWriter(listBox1));
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ListBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            var box = (ListBox)sender;
            textBox1.Text = box.Text;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private async void Button3_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.PrevodCsvToJson());
        }

        private async void Button4_Click(object sender, EventArgs e)
        {
            await Task.Run(() => Soubory.KillExcel());
        }

        private async void button5_Click(object sender, EventArgs e)
        {
            await Task.Run(() => LigthChem.KabelyAdd());
        }
    }

    public class ListBoxWriter(ListBox listBox) : TextWriter
    {
        private readonly ListBox _listBox = listBox;
        private readonly SynchronizationContext _context = SynchronizationContext.Current;

        public override Encoding Encoding => Encoding.UTF8;

        public override void WriteLine(string value)
        {
            //_context.Post(_ => _listBox.Items.Add(value), null);
            _context.Post(_ =>
            {
                _listBox.Items.Add(value);
                _listBox.TopIndex = _listBox.Items.Count - 1; // ← automatické scrollování dolů
            }, null);
        }

        //public override void Write(char value)
        //{
        //    // Nepřepisujeme po znacích, pouze řádky (volitelné)
        //}
    }
}
