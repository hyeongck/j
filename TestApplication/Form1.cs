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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Dir.Dir_Directory Dir = new Dir.Dir_Directory("C:\\Automation\\DB\\Yield");
            Dir = new Dir.Dir_Directory("C:\\Automation\\Yield");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Yield_Form Yield = new Yield_Form();
            if (Yield.ActiveControl == null)
            {
                Yield.Show();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Spec_Gen_Form SpecGenForm = new Spec_Gen_Form();
            if (SpecGenForm.ActiveControl == null)
            {
                SpecGenForm.Show();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Get_Spec_Form GetSpec = new Get_Spec_Form();
            if (GetSpec.ActiveControl == null)
            {
                GetSpec.Show();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Merge_Form Merge = new Merge_Form();
            if (Merge.ActiveControl == null)
            {
                Merge.Show();
            }
            int a = 0;
            int b = 0;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Box_Plot_Form Box = new Box_Plot_Form();
            if (Box.ActiveControl == null)
            {
                Box.Show();
            }
        }
    }
}
