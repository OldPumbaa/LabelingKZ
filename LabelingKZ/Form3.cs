using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace LabelingKZ
{
    public partial class Form3 : Form
    {
        XmlDocument xmldoc;
        string[] knopname = new string[8];
        string[] knopcorr = new string[8];
        string[] knopcomm = new string[8];
        string currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            xmldoc = new XmlDocument();

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
            }
            else
            {
                comboBox1.SelectedIndex = 0;
            }

            xmldoc.Load(currentDirectory + "/profiles/" + Properties.Settings.Default.LastProfile + ".xml");
            for (int i = 0; i < 8; i++)
            {
                XmlNode node1 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/name");
                knopname[i] = node1.InnerText;
                XmlNode node2 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/corr");
                knopcorr[i] = node2.InnerText;
                XmlNode node3 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/comm");
                knopcomm[i] = node3.InnerText;
            }
            button1.Text = knopname[0] + " (Q)";
            button2.Text = knopname[1] + " (W)";
            button3.Text = knopname[2] + " (E)";
            button4.Text = knopname[3] + " (R)";
            button5.Text = knopname[4] + " (T)";
            button6.Text = knopname[5] + " (Y)";
            button7.Text = knopname[6] + " (U)";
            button8.Text = knopname[7] + " (A)";
            textBox1.Text = knopname[0];
            textBox2.Text = knopname[1];
            textBox3.Text = knopname[2];
            textBox4.Text = knopname[3];
            textBox5.Text = knopname[4];
            textBox6.Text = knopname[5];
            textBox7.Text = knopname[6];
            textBox8.Text = knopname[7];
            textBox16.Text = knopcorr[0];
            textBox15.Text = knopcorr[1];
            textBox14.Text = knopcorr[2];
            textBox13.Text = knopcorr[3];
            textBox12.Text = knopcorr[4];
            textBox11.Text = knopcorr[5];
            textBox10.Text = knopcorr[6];
            textBox9.Text = knopcorr[7];
            textBox24.Text = knopcomm[0];
            textBox23.Text = knopcomm[1];
            textBox22.Text = knopcomm[2];
            textBox21.Text = knopcomm[3];
            textBox20.Text = knopcomm[4];
            textBox19.Text = knopcomm[5];
            textBox18.Text = knopcomm[6];
            textBox17.Text = knopcomm[7];
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.LastProfile = comboBox1.Text;
            Properties.Settings.Default.Save();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            xmldoc = new XmlDocument();
            string currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            xmldoc.Load(currentDirectory + "/profiles/" + comboBox1.Text + ".xml");
            for (int i = 0; i < 8; i++)
            {
                XmlNode node1 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/name");
                knopname[i] = node1.InnerText;
                XmlNode node2 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/corr");
                knopcorr[i] = node2.InnerText;
                XmlNode node3 = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i + 1) + "/comm");
                knopcomm[i] = node3.InnerText;
            }
            button1.Text = knopname[0] + " (Q)";
            button2.Text = knopname[1] + " (W)";
            button3.Text = knopname[2] + " (E)";
            button4.Text = knopname[3] + " (R)";
            button5.Text = knopname[4] + " (T)";
            button6.Text = knopname[5] + " (Y)";
            button7.Text = knopname[6] + " (U)";
            button8.Text = knopname[7] + " (A)";
            textBox1.Text = knopname[0];
            textBox2.Text = knopname[1];
            textBox3.Text = knopname[2];
            textBox4.Text = knopname[3];
            textBox5.Text = knopname[4];
            textBox6.Text = knopname[5];
            textBox7.Text = knopname[6];
            textBox8.Text = knopname[7];
            textBox16.Text = knopcorr[0];
            textBox15.Text = knopcorr[1];
            textBox14.Text = knopcorr[2];
            textBox13.Text = knopcorr[3];
            textBox12.Text = knopcorr[4];
            textBox11.Text = knopcorr[5];
            textBox10.Text = knopcorr[6];
            textBox9.Text = knopcorr[7];
            textBox24.Text = knopcomm[0];
            textBox23.Text = knopcomm[1];
            textBox22.Text = knopcomm[2];
            textBox21.Text = knopcomm[3];
            textBox20.Text = knopcomm[4];
            textBox19.Text = knopcomm[5];
            textBox18.Text = knopcomm[6];
            textBox17.Text = knopcomm[7];
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            button1.Text = textBox1.Text + " (Q)";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            button2.Text = textBox2.Text + " (W)";
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            button3.Text = textBox3.Text + " (E)";
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            button4.Text = textBox4.Text + " (R)";
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            button5.Text = textBox5.Text + " (T)";
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            button6.Text = textBox6.Text + " (Y)";
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            button7.Text = textBox7.Text + " (U)";
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            button8.Text = textBox8.Text + " (A)";
        }

        private void button16_Click(object sender, EventArgs e)
        {
            xmldoc = new XmlDocument();
            xmldoc.Load(currentDirectory + "/profiles/" + comboBox1.Text + ".xml");
            for (int i = 1; i <= 8; i++)
            {
                XmlNode node = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i) + "/name");
                System.Windows.Forms.Control control = Controls["textBox" + i];
                node.InnerText = control.Text;
            }
            for (int i = 1; i <= 8; i++)
            {
                XmlNode node = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i) + "/corr");
                System.Windows.Forms.Control control = Controls["textBox" + (17 - i)];
                node.InnerText = control.Text;
            }
            for (int i = 1; i <= 8; i++)
            {
                XmlNode node = xmldoc.DocumentElement.SelectSingleNode("/buttons/button" + (i) + "/comm");
                System.Windows.Forms.Control control = Controls["textBox" + (25 - i)];
                node.InnerText = control.Text;
            }
            xmldoc.Save(currentDirectory + "/profiles/" + comboBox1.Text + ".xml");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            using (Form4 form4 = new Form4())
            {
                form4.ShowDialog();
            }
            if (Properties.Settings.Default.newprofname != "")
            {
                XDocument xmlnew = new XDocument(new XElement("buttons",
                            new XElement("button1",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", "")),
                            new XElement("button2",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", "")),
                            new XElement("button3",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", "")),
                            new XElement("button4",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", "")),
                            new XElement("button5",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", "")),
                            new XElement("button6",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", "")),
                            new XElement("button7",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", "")),
                            new XElement("button8",
                                new XElement("name", ""),
                                new XElement("corr", ""),
                                new XElement("comm", ""))));
                xmlnew.Save(currentDirectory + "/profiles/" + Properties.Settings.Default.newprofname + ".xml");
                Properties.Settings.Default.newprofname = "";
                Properties.Settings.Default.Save();
            }
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
    }
}
