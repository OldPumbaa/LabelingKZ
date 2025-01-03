﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace LabelingKZ
{
    public partial class Form2 : Form
    {
        int page = 1;
        int rc;
        SLDocument doc;
        public System.Drawing.Size OldSize { get; set; }
        public Form2()
        {
            InitializeComponent();
            Environment.SetEnvironmentVariable("WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS", "--enable-features=msWebView2BrowserHitTransparent");
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string filePath = "";
            KeyPreview = true;
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
                doc = new SLDocument(filePath);
            }
            string str = System.IO.File.ReadAllText(@filePath);
            doc = new SLDocument(filePath);
            var stats = doc.GetWorksheetStatistics(); //sheet::SLDocument
            rc = stats.NumberOfRows;
            label1.Text = doc.GetCellValueAsString("A2");
            label2.Text = doc.GetCellValueAsString("C2");
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E2"));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F2"));
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B2")));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D2")));
            textBox1.Text = Convert.ToString(page);
            this.Text = "Labeling KZ Workspace: " + ofd.SafeFileName;
            for (int j = 1; j < rc; j++)
            {
                string job = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            doc.Save();
            
            System.Windows.Forms.Application.Exit();
        }

        private void webView21_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Resize(object sender, EventArgs e)
        {
            webView21.Width = this.Width / 2;
            webView22.Width = this.Width / 2;
        }

        private void Form2_ResizeEnd(object sender, EventArgs e)
        {

        }

        private void webView22_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

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
            label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (page + 1)));
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (page + 1)));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (page + 1)));
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (page + 1))));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (page + 1))));
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
            label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (page + 1)));
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (page + 1)));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (page + 1)));
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (page + 1))));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (page + 1))));
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
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
            label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (page + 1)));
            textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (page + 1)));
            textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (page + 1)));
            webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (page + 1))));
            webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (page + 1))));
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
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
                if (job != "0" && job != "1" && job != "2")
                {
                    textBox1.Text = Convert.ToString(j);
                    label1.Text = Convert.ToString(doc.GetCellValueAsString("A" + (j + 1)));
                    label2.Text = Convert.ToString(doc.GetCellValueAsString("C" + (j + 1)));
                    textBox2.Text = Convert.ToString(doc.GetCellValueAsString("E" + (j + 1)));
                    textBox3.Text = Convert.ToString(doc.GetCellValueAsString("F" + (j + 1)));
                    webView21.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("B" + (j + 1))));
                    webView22.Source = new System.Uri(Convert.ToString(doc.GetCellValueAsString("D" + (j + 1))));
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), "1");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), "0");
            doc.SetCellValue("F" + (page + 1), "бренд");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), "0");
            doc.SetCellValue("F" + (page + 1), "модель");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), "0");
            doc.SetCellValue("F" + (page + 1), "тип товара");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), "0");
            doc.SetCellValue("F" + (page + 1), "бу");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), "1");
            doc.SetCellValue("F" + (page + 1), "разные названия одного бренда");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            doc.SetCellValue("E" + (page + 1), "0");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                }
                label5.Text = "Документ готов!";
            }
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {
            
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
            doc.SetCellValue("E" + (page + 1), "2");
            doc.SetCellValue("F" + (page + 1), "отсутствует страница");
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
                if (job != "0" && job != "1" && job != "2")
                {
                    label5.Text = "Следующая строка: " + j;
                    break;
                } 
                label5.Text = "Документ готов!";
            }
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            if (textBox3.Focused)
            {
                if (e.KeyCode == Keys.Enter) 
                { 
                    textBox2.Focus();
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
                else if (e.KeyCode == Keys.R)
                {
                    button7_Click(sender, e);
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
                else if (e.KeyCode == Keys.R)
                {
                    button7_Click(sender, e);
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
                else if (e.KeyCode == Keys.R)
                {
                    button7_Click(sender, e);
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
    }
}
