using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestApplication
{
    public partial class Insert_Count_Form : Form
    {

        public Insert_Count_Form()
        {
            InitializeComponent();
            label1.Show();
            label2.Show();
            label3.Show();
        }
        public void Print_Count(long Count)
        {
            label3.Text = Convert.ToString(Count);
        }
    }
}
