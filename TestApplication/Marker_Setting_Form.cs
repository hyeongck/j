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
    public delegate double[] Marker_Set(int index , string Test, string Freq, int Row);

    public partial class Marker_Setting_Form : Form
    {
        public static int Snp_Index;

        int i;
        int Index;
        string Chan;
        string Tab_Text;
        string[] ItemList;

        Dictionary<string, Dictionary<string, Dictionary<string, double[]>>> Snp_Data;
        Dictionary<string, Dictionary<string, string>> _Marker_Data;
        System.Windows.Forms.DataVisualization.Charting.Chart chart;

        public System.Windows.Forms.ListView[] listView1;

        public static event Marker_Set Marker_Set_Send;


        public Marker_Setting_Form(string[] ItemList,int index, System.Windows.Forms.DataVisualization.Charting.Chart chart, string Chan, string Tab_Text)
        {
      
            this.chart = chart;
            this.Chan = Chan;
            this.Tab_Text = Tab_Text;
            this.Index = index;
            this.ItemList = ItemList;
            InitializeComponent();

            listView1 = new ListView[ItemList.Length];

            for (i = 0; i < ItemList.Length; i ++)
            {
                TabPage myTabPage = new TabPage(ItemList[i].ToString());
                tabControl1.TabPages.Add(myTabPage);

                listView1[i] = new ListView();

                listView1[i].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
                listView1[i].Location = new System.Drawing.Point(3, 6);
                listView1[i].Name = "listView1";
                listView1[i].Size = new System.Drawing.Size(535, 448);
                listView1[i].TabIndex = 1;
                listView1[i].UseCompatibleStateImageBehavior = false;

                listView1[i].GridLines = true;
                listView1[i].FullRowSelect = true;

                tabControl1.TabPages[i].Controls.Add(listView1[i]);

             
            }
          //  Run();
        }
 
        public Marker_Setting_Form(string[] ItemList, string Chan, string Tab_Text)
        {


            this.Chan = Chan;
            this.Tab_Text = Tab_Text;
            this.ItemList = ItemList;
            InitializeComponent();
            Flag_Form = true;

            listView1 = new ListView[ItemList.Length];

            for (i = 0; i < ItemList.Length; i++)
            {
                TabPage myTabPage = new TabPage(ItemList[i].ToString());
                tabControl1.TabPages.Add(myTabPage);

                listView1[i] = new ListView();

                listView1[i].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
                listView1[i].Location = new System.Drawing.Point(3, 6);
                listView1[i].Name = "listView1";
                listView1[i].Size = new System.Drawing.Size(535, 448);
                listView1[i].TabIndex = 1;
                listView1[i].UseCompatibleStateImageBehavior = false;

                listView1[i].GridLines = true;
                listView1[i].FullRowSelect = true;

                tabControl1.TabPages[i].Controls.Add(listView1[i]);

        
            }
            //  Run();
        }


        public void Listview_Run(int index, int marker)
        {


            if (Snp_Data.Count == 1)
            {
                listView1[index].Clear();
                listView1[index].BeginUpdate();

                listView1[index].View = View.Details;

                foreach (Dictionary<string, Dictionary<string, double[]>> T in Snp_Data.Values)
                {
                    int j = 0;

                    foreach (Dictionary<string, double[]> _T in T.Values)
                    {
                        if (j == index)
                        {
                            foreach (double[] _T_Data in _T.Values)
                            {
                                i = 0;
                                foreach (double Row in _T_Data)
                                {
                                    ListViewItem Ivi = new ListViewItem(chart.Series[i].Name.ToString());

                                    Ivi.SubItems.Add(Convert.ToString(_T_Data[i]));
                                    listView1[index].Items.Add(Ivi);
                                    i++;
                                }
                                listView1[index].Columns.Add("SN");
                                listView1[index].Columns.Add("Marker" + (marker + 1));

                            }

                        }
                        j++;
                    }

                }

                listView1[index].EndUpdate();
            }
            else
            {
                int x = Snp_Data.Count;
                int y = 0;
                int Unit_Count = 0;

                foreach (Dictionary<string, Dictionary<string, double[]>> T in Snp_Data.Values)
                {
                    int j = 0;

                    foreach (Dictionary<string, double[]> _T in T.Values)
                    {
                        foreach (double[] _T_data in _T.Values)
                        {
                            Unit_Count = _T_data.Length;
                        }
                        y = T.Values.Count;
                        break;
                    }
                    break;

                }

                int x_N = 0;
                int q = 0;
                bool flag_2 = false;
                //  if(index == 0)


                Dictionary<string, string> t = new Dictionary<string, string>();

                t = _Marker_Data[Chan];

                List<string> A = new List<string>();

                foreach (KeyValuePair<string, string> sd in t)
                {
                    if (sd.Value != "")
                    {
                        A.Add(sd.Key);
                    }
                }

                t = new Dictionary<string, string>();


                int Snp_index = 0;

                for (int k = 0; k < y; k++)
                {
                    listView1[k].Clear();
                    listView1[k].BeginUpdate();

                    listView1[k].View = View.Details;
                    listView1[k].Columns.Add("SN");
                }

                int T_Index = 0;

                for (int G = 0; G < Unit_Count; G++)
                {
                    index = 0;

                    for (int k = 0; k < y; k++)
                    {
                        ListViewItem Ivi = new ListViewItem();
                        bool flag = false;

                        Ivi = new ListViewItem(chart.Series[x_N].Name.ToString());

                        foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, double[]>>> T in Snp_Data)
                        {

                            if (flag) T_Index = 0;

                            bool Snp_Flag = false;
                            foreach (KeyValuePair<string, Dictionary<string, double[]>> _T in T.Value)
                            {
                                if (T_Index == Snp_index)
                                {
                                    foreach (KeyValuePair<string, double[]> _T_Data in _T.Value)
                                    {
                                        for (int h = G; h < G + 1; h++)
                                        {
                                            Ivi.SubItems.Add(Convert.ToString(_T_Data.Value[G]));

                                        }
                                        x_N++;
                                        flag = true;

                                        break;
                                    }

                                    if (x_N == x)
                                    {
                                        if (!flag_2)
                                        {
                                            for (int d = 0; d < A.Count; d++)
                                            {
                                                listView1[index].Columns.Add(A[d]);
                                            }
                                            flag_2 = true;

                                        }

                                        listView1[index].Items.Add(Ivi);
                                        x_N = 0;
                                        Snp_index++;

                                    }
                                }
                                else if (!flag_2) T_Index++;
                                else
                                {
                                    T_Index++;
                                }


                                if (flag)
                                {
                                    flag = false;
                                    break;
                                }

                            }

                        }

                        index++;

                    }
                }
                for (int k = 0; k < y; k++)
                {

                    listView1[k].EndUpdate();
                }

            }



        }

        public bool Flag_Form;
        private void Marker_Setting_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            Flag_Form = false;
            Dispose(true);
        }
    }
}
