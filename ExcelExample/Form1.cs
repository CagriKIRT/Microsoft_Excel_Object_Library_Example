using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelExample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            OpenFile();
        }
        public void OpenFile()
        {
            Excel excel = new Excel(AppDomain.CurrentDomain.BaseDirectory+ "\\TEST.xlsx", 1);
            MessageBox.Show(excel.ReadCell(0, 0));
        }

    }
}
