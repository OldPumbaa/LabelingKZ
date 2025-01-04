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
            MessageBox.Show("v.1.2.5:\n -Добавлено уведомление о готовности документа;\n -Кнопка выхода из приложения заменена на кнопку выбора другого файла;\n -Сохранение теперь производится автоматически;\n -Окно теперь имеет стандартные кнопки Windows, которые можно использовать;\n -Кнопка РНОБ заменена на Количество (в будущем кнопки будут более гибкими);\n -Добавлены кнопки на изменение зума (можно на Ctrl + колесико мыши) и перезагрузку обоих браузеров (можно на Ctrl + R);\n -Измененный зум сохраняется между сессиями в приложении;\n -При отстутствии страницы на Озон теперь появляется уведомление, результат автоматически заполняется и программа перелистывается на следующее сравнение.\nv1.2:\n -Небольшие визуальные изменения;\n -Изменен зум у браузеров;\n -Добавлены горячие клавиши.\n");
        }
    }
}
