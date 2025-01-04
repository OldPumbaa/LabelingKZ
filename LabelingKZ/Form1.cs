using LabelingKZ;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LabelingKZ
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Text = "Labeling KZ";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            using (Form2 form2 = new Form2())
            {
                form2.ShowDialog();
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("v1.2:\n -Небольшие визуальные изменения;\n -Изменен зум у браузеров;\n -Добавлены горячие клавиши.\nv1.1.6:\n -Теперь считается количество строк.\nv1.1.5:\n -Добавлена кнопка \"404\";\n -Исправлен зум браузеров.\nv1.1:\n -Переписано под NET Framework; \n -Визуальные доработки; \n -Заменена библиотека (теперь не нужен Microsoft Excel).");
        }
    }
}
