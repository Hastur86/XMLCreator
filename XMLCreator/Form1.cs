using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XMLCreator
{
    public partial class Form1 : Form
    {
        private string XMLFile { get; set; }
        private string ExcelFile { get; set; }
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.Cancel)
            {
                label1.Text = "Файл не выбран";
                return;
            }

            XMLFile = openFileDialog1.FileName;
            label1.Text = "Шаблон XML выбран";
        }
    }
}
