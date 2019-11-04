using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ATE
{
    public partial class Marker_Form : Form
    {
        Dictionary<string, Dictionary<string, Dictionary<string, double[]>>> Snp_Data;

        double[] Data;
        System.Windows.Forms.DataVisualization.Charting.Chart chart;
        int i;

        public Marker_Form(string Char, string Freq, string Tab_Text, System.Windows.Forms.DataVisualization.Charting.Chart chart, double[] Data)
        {
            this.chart = chart;
            this.Data = Data;

            if (Form.ActiveForm == null)
            {
                Snp_Data = new Dictionary<string, Dictionary<string, Dictionary<string,double[]>>>();

                Dictionary<string, Dictionary<string, double[]>> Tag = new Dictionary<string, Dictionary<string,double[]>>();
                Dictionary<string, double[]> Row_Data = new Dictionary<string, double[]>();

                if (!Row_Data.ContainsKey(Freq))
                    Row_Data.Add(Freq, Data);

                if (!Tag.ContainsKey(Tab_Text))
                    Tag.Add(Tab_Text, Row_Data);

                if (!Snp_Data.ContainsKey(Char))
                    Snp_Data.Add(Char, Tag);

                InitializeComponent();
                Run();

                listView1.GridLines = true;
                listView1.FullRowSelect = true;

                this.Show();
            }
        }
        public void Run()
        {
            listView1.BeginUpdate();

            listView1.View = View.Details;

 
            for (i = 0; i < Data.Length; i++)
            {
                ListViewItem Ivi = new ListViewItem(chart.Series[i].Name.ToString());

                Ivi.SubItems.Add(Convert.ToString(Data[i]));
                listView1.Items.Add(Ivi);

    
            }

            listView1.Columns.Add("Parameter");

            for (i = 0; i < 1; i ++)
            {
                listView1.Columns.Add("Marker" + (i + 1));
            }

        

            listView1.EndUpdate();

        }
        
    }
}
