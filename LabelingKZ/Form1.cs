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
            MessageBox.Show("v.1.3:\n -Исправлены некоторые баги;\n -Исправлена ошибка при отмене выбора файла;\n -Добавлен выпадающий список со всеми ответами в таблице;\n -Если товара нет в наличии или не осуществляется доставка на Озон, то страница все равно открывается и прокручивается до товара;\n -Добавлена система профилей для кнопок, теперь можно добавлять свои назначения кнопкам (есть уже заготовленные профили под разные категории, основанные на самых популярных ответах);\n -Имена товаров теперь сокращаются, если они слишком большие.\n\n\nВсе иконки предоставлены сайтом icons8.ru");
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
