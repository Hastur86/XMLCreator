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
using Excel = Microsoft.Office.Interop.Excel;

namespace XMLCreator
{
    public partial class Form1 : Form
    {
        private string XMLFile { get; set; }
        private string ExcelFile { get; set; }
        private XmlDocument NewXML { get; set; }
        private Excel.Application ex { get; set; }

        public Form1()
        {
            InitializeComponent();
            ex = new Microsoft.Office.Interop.Excel.Application();
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
            label2.Text = "Файл данных выбран";
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

            Excel.Workbook workBook;
            try
            {
                listBox1.Items.Add("Открытие Excel файла");
                workBook = ex.Workbooks.Open(ExcelFile);
            }
            catch (Exception exception)
            {
                listBox1.Items.Add("Ошибка - " + exception.Message);
                return;
            }

            try
            {
                Excel.Worksheet sheet = workBook.Sheets[1];
                int i = 1;
                while (sheet.Cells[i,1].Value2.ToString() != "Конец")
                {
                    string nodeName = sheet.Cells[i, 1].Value2.ToString();
                    string attrName = sheet.Cells[i, 2].Value2 != null ? sheet.Cells[i, 2].Value2.ToString() : "";
                    string nodeValue = sheet.Cells[i, 3].Value2 != null ? sheet.Cells[i, 3].Value2.ToString() : ""; ;

                    if (nodeName != "Спека")
                    {
                        if (attrName != "")
                        {
                            if (nodeName.Contains("ИнфПолФХЖ1"))
                            {
                                listBox1.Items.Add(attrName);
                                XmlNodeList aNodes = NewXML.SelectNodes(nodeName);
                                foreach (XmlNode aNode in aNodes)
                                {
                                    if (aNode.Attributes["Идентиф"].Value == attrName)
                                    {
                                        aNode.Attributes["Значен"].Value = nodeValue;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                listBox1.Items.Add(attrName);
                                NewXML.SelectSingleNode(nodeName).Attributes[attrName].Value = nodeValue;
                            }
                        }
                        else
                        {
                            listBox1.Items.Add(attrName);
                            NewXML.SelectSingleNode(nodeName).InnerText = nodeValue;
                        }
                    }
                    else
                    {
                        i++;
                        XmlNode specNode = NewXML.SelectSingleNode(sheet.Cells[i, 1].Value2.ToString());
                        i = i + 2;

                        while (sheet.Cells[i, 1].Value2.ToString() != "Конец")
                        {
                            listBox1.Items.Add(sheet.Cells[i, 1].Value2.ToString());
                            XmlElement newSvedTov = NewXML.CreateElement("СведТов");

                            XmlAttribute attr1 = NewXML.CreateAttribute("НомСтр");
                            attr1.Value = sheet.Cells[i, 1].Value2.ToString();
                            XmlAttribute attr2 = NewXML.CreateAttribute("НаимТов");
                            attr2.Value = sheet.Cells[i, 2].Value2.ToString();
                            XmlAttribute attr3 = NewXML.CreateAttribute("ОКЕИ_Тов");
                            attr3.Value = sheet.Cells[i, 3].Value2.ToString();
                            XmlAttribute attr4 = NewXML.CreateAttribute("КолТов");
                            attr4.Value = sheet.Cells[i, 4].Value2.ToString();
                            XmlAttribute attr5 = NewXML.CreateAttribute("ЦенаТов");
                            attr5.Value = sheet.Cells[i, 5].Value2.ToString();
                            XmlAttribute attr6 = NewXML.CreateAttribute("СтТовБезНДС");
                            attr6.Value = sheet.Cells[i, 6].Value2.ToString();
                            XmlAttribute attr7 = NewXML.CreateAttribute("НалСт");
                            attr7.Value = sheet.Cells[i, 7].Value2.ToString();
                            XmlAttribute attr8 = NewXML.CreateAttribute("СтТовУчНал");
                            attr8.Value = sheet.Cells[i, 8].Value2.ToString();

                            newSvedTov.Attributes.Append(attr1);
                            newSvedTov.Attributes.Append(attr2);
                            newSvedTov.Attributes.Append(attr3);
                            newSvedTov.Attributes.Append(attr4);
                            newSvedTov.Attributes.Append(attr5);
                            newSvedTov.Attributes.Append(attr6);
                            newSvedTov.Attributes.Append(attr7);
                            newSvedTov.Attributes.Append(attr8);

                            XmlElement newAkciz = NewXML.CreateElement("Акциз");
                            XmlElement newBezAkciz = NewXML.CreateElement("БезАкциз");
                            newBezAkciz.InnerText = sheet.Cells[i, 9].Value2.ToString();
                            newAkciz.AppendChild(newBezAkciz);

                            XmlElement newSumNal = NewXML.CreateElement("СумНал");
                            XmlElement newBezNds = NewXML.CreateElement("БезНДС");
                            newBezNds.InnerText = sheet.Cells[i, 10].Value2.ToString();
                            newSumNal.AppendChild(newBezNds);

                            XmlElement newDopSvedTov = NewXML.CreateElement("ДопСведТов");
                            XmlAttribute attr9 = NewXML.CreateAttribute("НаимЕдИзм");
                            attr9.Value = sheet.Cells[i, 11].Value2.ToString();
                            newDopSvedTov.Attributes.Append(attr9);

                            newSvedTov.AppendChild(newAkciz);
                            newSvedTov.AppendChild(newSumNal);
                            newSvedTov.AppendChild(newDopSvedTov);

                            specNode.PrependChild(newSvedTov);

                            i++;
                        }

                        i--;
                    }

                    i++;
                }
            }
            catch (Exception exception)
            {
                ex.Quit();
                listBox1.Items.Add("Ошибка - " + exception.Message);
                return;
            }

            ex.Quit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var result = saveFileDialog1.ShowDialog();
            if (result == DialogResult.Cancel)
            {
                return;
            }

            try
            {
                NewXML.Save(saveFileDialog1.FileName);
                listBox1.Items.Add(label1.Text = "Файл сохранен");
            }
            catch (Exception exception)
            {
                listBox1.Items.Add("Ошибка - " + exception.Message);
            }
        }
    }
}
