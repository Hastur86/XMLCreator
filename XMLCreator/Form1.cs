using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XMLCreator
{
    public partial class Form1 : Form
    {
        private string XMLFile { get; set; }
        private string ExcelFile { get; set; }
        private XmlDocument NewXML { get; set; }

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

        private void button2_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.Cancel)
            {
                label2.Text = "Файл не выбран";
                return;
            }

            ExcelFile = openFileDialog1.FileName;
            label2.Text = "Шаблон XML выбран";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.Items.Add("Старт");

            try
            {
                listBox1.Items.Add("Открытие XML файла");
                NewXML = new XmlDocument();
                NewXML.Load(XMLFile);
                listBox1.Items.Add("Успех");
            }
            catch (Exception exception)
            {
                listBox1.Items.Add("Ошибка - " + exception.Message);
                return;
            }

            try
            {

            }
            catch (Exception exception)
            {
                listBox1.Items.Add("Ошибка - " + exception.Message);
                return;
            }
        }
    }
}
