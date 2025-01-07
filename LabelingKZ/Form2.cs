using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System.IO;
using LabelingKZ.Properties;

namespace LabelingKZ
{
    public partial class Form2 : Form
    {
        string[] knopname = new string[8];
        string[] knopcorr = new string[8];
        string[] knopcomm = new string[8];
        int page = 1;
        int rc;
        const int maxlength = 130;
        bool fdone;
        bool tcomp;
        string filePath = "";
        SLDocument doc;
        XmlDocument xmldoc;
        public System.Drawing.Size OldSize { get; set; }
        public Form2()
        {
            InitializeComponent();
            Environment.SetEnvironmentVariable("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS", "--enable-features=msWebView2BrowserHitTransparent");
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            fdone = false;
            KeyPreview = true;
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
                doc = new SLDocument(filePath);
            } else {
                this.Hide();
                using (Form1 form1 = new Form1())
                {
                    form1.ShowDialog();
                }
            }

            //INITIALIZE PROFILE
            xmldoc = new XmlDocument();
            string currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            string[] profiles = Directory.GetFiles(currentDirectory + "/profiles/");
            List<string> profilenames = new List<string>();

            foreach (string profile in profiles)
            {
                comboBox1.Items.Add(Path.GetFileNameWithoutExtension(profile));
                profilenames.Add(Path.GetFileNameWithoutExtension(profile));
            }

            if (profilenames.Contains(Properties.Settings.Default.LastProfile))
            {
                comboBox1.SelectedIndex = comboBox1.FindStringExact(Properties.Settings.Default.LastProfile);
                xmldoc.Load(currentDirectory + "/profiles/" + Properties.Settings.Default.LastProfile + ".xml");
            } else
            {
                comboBox1.SelectedIndex = 0;
                xmldoc.Load(currentDirectory + "/profiles/" + comboBox1.Items[comboBox1.SelectedIndex] + ".xml");
            }

            for (int i = 0; i < 8; i++)
            {
                XmlNode node1 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/name");
                knopname[i] = node1.InnerText;
                XmlNode node2 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/corr");
                knopcorr[i] = node2.InnerText;
                XmlNode node3 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/comm");
                knopcomm[i] = node3.InnerText;
            }
            button4.Text = knopname[0] + " (Q)";
            button11.Text = knopname[1] + " (W)";
            button5.Text = knopname[2] + " (E)";
            button7.Text = knopname[3] + " (R)";
            button8.Text = knopname[4] + " (T)";
            button9.Text = knopname[5] + " (Y)";
            button10.Text = knopname[6] + " (U)";
            button12.Text = knopname[7] + " (A)";

            //PROFILES END

            string str = System.IO.File.ReadAllText(@filePath);
            doc = new SLDocument(filePath);
            var stats = doc.GetWorksheetStatistics(); //sheet::SLDocument
            rc = stats.NumberOfRows;
            for (int i = 2; i <= rc; i++)
            {
                DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                row.Cells[0].Value = i - 1;
                row.Cells[1].Value = Convert.ToString(doc.GetCellValueAsString("E" + i));
                row.Cells[2].Value = Convert.ToString(doc.GetCellValueAsString("F" + i));
                dataGridView1.Rows.Add(row);
            }
            label1.Text = doc.GetCellValueAsString("A2");
            if (label1.Text.Length > maxlength)
            {
                label1.Text = label1.Text.Substring(0, maxlength);
            }
            label2.Text = doc.GetCellValueAsString("C2");
            if (label2.Text.Length > maxlength)
            {
                label2.Text = label2.Text.Substring(0, maxlength);
            }
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E2"));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F2"));
            webView21.ZoomFactor = Properties.Settings.Default.Zoom1;
            webView22.ZoomFactor = Properties.Settings.Default.Zoom2;
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B2")));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D2")) + "?oos_search=false");
            textBox1.Text = Convert.ToString(page);
            this.Text = "Labeling KZ Workspace: " + ofd.SafeFileName;
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Zoom1 = webView21.ZoomFactor;
            Properties.Settings.Default.Zoom2 = webView22.ZoomFactor;
            Properties.Settings.Default.LastProfile = comboBox1.Text;
            Properties.Settings.Default.Save();
            this.Hide();
            using (Form1 form1 = new Form1())
            {
                form1.ShowDialog();
            }
        }

        private void Form2_Resize(object sender, EventArgs e)
        {
            webView21.Width = this.Width / 2;
            webView22.Width = this.Width / 2;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (page < (rc - 1))
            {
                page++;
            }
            else
            {
                page = 1;
            }
            textBox1.Text = Convert.ToString(page);
            label1.Text = Convert.ToString(doc.GetCellValueAsString("A" + (page + 1)));
            if (label1.Text.Length > maxlength)
            {
                label1.Text = label1.Text.Substring(0, maxlength);
            }
            label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (page + 1)));
            if (label2.Text.Length > maxlength)
            {
                label2.Text = label2.Text.Substring(0, maxlength);
            }
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (page + 1)));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (page + 1)));
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (page + 1))));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (page + 1))) + "?oos_search=false");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (page > 1)
            {
                page--;
            }
            else
            {
                page = (rc - 1);
            }
            textBox1.Text = Convert.ToString(page);
            label1.Text = Convert.ToString(doc.GetCellValueAsString("A" + (page + 1)));
            if (label1.Text.Length > maxlength)
            {
                label1.Text = label1.Text.Substring(0, maxlength);
            }
            label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (page + 1)));
            if (label2.Text.Length > maxlength)
            {
                label2.Text = label2.Text.Substring(0, maxlength);
            }
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (page + 1)));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (page + 1)));
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (page + 1))));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (page + 1))) + "?oos_search=false");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "1";
            }
            page = Convert.ToInt32(textBox1.Text);
            if (page < 1)
            {
                page = 1;
                textBox1.Text = Convert.ToString(page);
            }
            else if (page > (rc - 1))
            {
                page = (rc - 1);
                textBox1.Text = Convert.ToString(page);
            }
            label1.Text = Convert.ToString(doc.GetCellValueAsString("A" + (page + 1)));
            if (label1.Text.Length > maxlength)
            {
                label1.Text = label1.Text.Substring(0, maxlength);
            }
            label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (page + 1)));
            if (label2.Text.Length > maxlength)
            {
                label2.Text = label2.Text.Substring(0, maxlength);
            }
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (page + 1)));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (page + 1)));
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (page + 1))));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (page + 1))) + "?oos_search=false");
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), textBox2.Text);
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            doc.SetCellValue("F" + (page + 1), textBox3.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    textBox1.Text = Convert.ToString(j);
                    label1.Text = Convert.ToString(doc.GetCellValueAsString("A" + (j + 1))); if (label1.Text.Length > maxlength)
                    {
                        label1.Text = label1.Text.Substring(0, maxlength);
                    }
                    label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (j + 1)));
                    if (label2.Text.Length > maxlength)
                    {
                        label2.Text = label2.Text.Substring(0, maxlength);
                    }
                    textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                    textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (j + 1)));
                    webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (j + 1))));
                    webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (j + 1))) + "?oos_search=false");
                    tcomp = false;
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[0]);
            if (knopcomm[0] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[0]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            }
            else
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                } else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[2]);
            if (knopcomm[2] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[2]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            }
            else
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[3]);
            if (knopcomm[3] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[3]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            }
            else
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[4]);
            if (knopcomm[4] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[4]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            }
            else
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[5]);
            if (knopcomm[5] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[5]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            }
            else
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[6]);
            if (knopcomm[6] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[6]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            }
            else
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[1]);
            if (knopcomm[1] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[1]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            }
            else
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void label2_Resize(object sender, EventArgs e)
        {
            LabelControl labelControl = sender as LabelControl;
            int diff = label2.Size.Width - OldSize.Width;
            label2.Left -= diff;
        }

        private void label2_TextChanged(object sender, EventArgs e)
        {
            LabelControl labelControl = sender as LabelControl;
            OldSize = label2.Size;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), knopcorr[7]);
            if (knopcomm[7] != "")
            {
                doc.SetCellValue("F" + (page + 1), knopcomm[7]);
            }
            page++;
            if (page < (rc))
            {
                textBox1.Text = Convert.ToString(page);
            } 
            else 
            {
                textBox1.Text = Convert.ToString("1");
            }
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job == "0" || job == "1" || job == "2")
                {
                    label5.Text = "Документ готов!";
                    tcomp = true;
                    continue;
                }
                else
                {
                    label5.Text = "Следующая строка: " + j;
                    tcomp = false;
                    break;
                }
            }
            if (!fdone && tcomp)
            {
                fdone = true;
                System.Media.SystemSounds.Asterisk.Play();
                MessageBox.Show("Документ готов!");
            }
            doc.Save();
            doc = new SLDocument(filePath);
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            if (textBox3.Focused)
            {
                if (e.KeyCode == Keys.Enter) 
                { 
                    textBox2.Focus();
                    doc.Save();
                    doc = new SLDocument(filePath);
                }
            }
            else
            {
                if (e.KeyCode == Keys.Q)
                {
                    button4_Click_1(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.W)
                {
                    button11_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.E)
                {
                    button5_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.R && e.Modifiers != Keys.Control)
                {
                    button7_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.R && e.Modifiers == Keys.Control)
                {
                    button13_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.T)
                {
                    button8_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Y)
                {
                    button9_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.U)
                {
                    button10_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.A)
                {
                    button12_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.S)
                {
                    textBox3.Focus();
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Left)
                {
                    button2_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Right)
                {
                    button3_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Space)
                {
                    button6_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
            }
        }

        private void webView21_KeyDown(object sender, KeyEventArgs e)
        {

            if (textBox3.Focused)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    textBox2.Focus();
                    doc.Save();
                    doc = new SLDocument(filePath);
                }
            }
            else
            {
                if (e.KeyCode == Keys.Q)
                {
                    button4_Click_1(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.W)
                {
                    button11_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.E)
                {
                    button5_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.R && e.Modifiers != Keys.Control)
                {
                    button7_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.R && e.Modifiers == Keys.Control)
                {
                    button13_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.T)
                {
                    button8_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Y)
                {
                    button9_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.U)
                {
                    button10_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.A)
                {
                    button12_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.S)
                {
                    textBox3.Focus();
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Left)
                {
                    button2_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Right)
                {
                    button3_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Space)
                {
                    button6_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
            }
        }

        private void webView22_KeyDown(object sender, KeyEventArgs e)
        {

            if (textBox3.Focused)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    textBox2.Focus();
                    doc.Save();
                    doc = new SLDocument(filePath);
                }
            }
            else
            {
                if (e.KeyCode == Keys.Q)
                {
                    button4_Click_1(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.W)
                {
                    button11_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.E)
                {
                    button5_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.R && e.Modifiers != Keys.Control)
                {
                    button7_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.R && e.Modifiers == Keys.Control)
                {
                    button13_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.T)
                {
                    button8_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Y)
                {
                    button9_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.U)
                {
                    button10_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.A)
                {
                    button12_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.S)
                {
                    textBox3.Focus();
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Left)
                {
                    button2_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Right)
                {
                    button3_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Space)
                {
                    button6_Click(sender, e);
                    e.SuppressKeyPress = true;
                }
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Zoom1 = webView21.ZoomFactor;
            Properties.Settings.Default.Zoom2 = webView22.ZoomFactor;
            Properties.Settings.Default.LastProfile = comboBox1.Text;
            Properties.Settings.Default.Save();
            Application.Exit();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            webView21.Reload();
            webView22.Reload();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            webView21.ZoomFactor -= 0.05;
            webView22.ZoomFactor -= 0.05;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            webView21.ZoomFactor += 0.05;
            webView22.ZoomFactor += 0.05;
        }

        private async void webView22_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e)
        {
            string pageContent = await webView22.ExecuteScriptAsync("document.body.innerText");
            string anscell = textBox2.Text;
            if ((pageContent.Contains("Такой страницы не существует")) && (anscell != "2"))
            {
                textBox2.Text = "2";
                System.Media.SystemSounds.Asterisk.Play();
                AutoClosingMessageBox.Show("Товар отсутствует, пропускаем", "404", 1000);
                page++;
                if (page < (rc))
                {
                    textBox1.Text = Convert.ToString(page);
                }
                else
                {
                    textBox1.Text = Convert.ToString("1");
                }
                for (int j = 1; j < rc; j++)
                {
                    string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                    if (job == "0" || job == "1" || job == "2")
                    {
                        label5.Text = "Документ готов!";
                        tcomp = true;
                        continue;
                    }
                    else
                    {
                        label5.Text = "Следующая строка: " + j;
                        tcomp = false;
                        break;
                    }
                }
                if (!fdone && tcomp)
                {
                    fdone = true;
                    System.Media.SystemSounds.Asterisk.Play();
                    MessageBox.Show("Документ готов!");
                }
                doc.Save();
                doc = new SLDocument(filePath);
            } else if (pageContent.Contains("Этот товар закончился") || pageContent.Contains("Товар не доставляется в ваш регион"))
            {
                await webView22.ExecuteScriptAsync("window.scroll(0, 500)");
            }
        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            textBox1.SelectAll();
        }

        private void webView21_ZoomFactorChanged(object sender, EventArgs e)
        {

            Properties.Settings.Default.Zoom1 = webView21.ZoomFactor;
            Properties.Settings.Default.Save();
        }

        private void webView22_ZoomFactorChanged(object sender, EventArgs e)
        {

            Properties.Settings.Default.Zoom2 = webView22.ZoomFactor;
            Properties.Settings.Default.Save();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Visible)
            {
                dataGridView1.Visible = false;
                dataGridView1.Enabled = false;
            } else
            {
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();
                for (int i = 2; i <= rc; i++)
                {
                    DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                    row.Cells[0].Value = i - 1;
                    row.Cells[1].Value = Convert.ToString(doc.GetCellValueAsString("E" + i));
                    row.Cells[2].Value = Convert.ToString(doc.GetCellValueAsString("F" + i));
                    dataGridView1.Rows.Add(row);
                }
                dataGridView1.Visible = true;
                dataGridView1.Enabled = true;
            }
        }

        private void tableLayoutPanel1_MouseLeave(object sender, EventArgs e)
        {
            dataGridView1.Visible = false;
            dataGridView1.Enabled = false;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[0].Value);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            xmldoc = new XmlDocument();
            string currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            xmldoc.Load(currentDirectory + "/profiles/" + comboBox1.Items[comboBox1.SelectedIndex] + ".xml");
            for (int i = 0; i < 8; i++)
            {
                XmlNode node1 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/name");
                knopname[i] = node1.InnerText;
                XmlNode node2 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/corr");
                knopcorr[i] = node2.InnerText;
                XmlNode node3 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/comm");
                knopcomm[i] = node3.InnerText;
            }
            button4.Text = knopname[0] + " (Q)";
            button11.Text = knopname[1] + " (W)";
            button5.Text = knopname[2] + " (E)";
            button7.Text = knopname[3] + " (R)";
            button8.Text = knopname[4] + " (T)";
            button9.Text = knopname[5] + " (Y)";
            button10.Text = knopname[6] + " (U)";
            button12.Text = knopname[7] + " (A)";
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            int current = comboBox1.SelectedIndex;
            comboBox1.Items.Clear();
            xmldoc = new XmlDocument();
            string currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            string[] profiles = Directory.GetFiles(currentDirectory + "/profiles/");
            List<string> profilenames = new List<string>();

            foreach (string profile in profiles)
            {
                comboBox1.Items.Add(Path.GetFileNameWithoutExtension(profile));
            }
            comboBox1.SelectedIndex = current;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.LastProfile = comboBox1.Text;
            Properties.Settings.Default.Save();
            using (Form3 form3 = new Form3())
            {
                form3.ShowDialog();
            }
            comboBox1.Items.Clear();
            xmldoc = new XmlDocument();
            string currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            string[] profiles = Directory.GetFiles(currentDirectory + "/profiles/");
            List<string> profilenames = new List<string>();

            foreach (string profile in profiles)
            {
                comboBox1.Items.Add(Path.GetFileNameWithoutExtension(profile));
            }
            comboBox1.SelectedIndex = 0;
            comboBox1.SelectedIndex = comboBox1.FindStringExact(Properties.Settings.Default.LastProfile);
        }
    }
}
