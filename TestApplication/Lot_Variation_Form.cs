using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace TestApplication
{

    public partial class Lot_Variation_Form : Form
    {
        string[] Length;
        Dictionary<string, int>[] Test_Result;
        int[] Yield;
        int[] Total;
        int[] Bin_Total;
        int Total_sample;

        string Path;

        Dictionary<string, Data_Class.Data_Editing.Clotho_Spec> For_Bin;

        CSV_Class.CSV.INT Csv_Interface;
        CSV_Class.CSV CSV = new CSV_Class.CSV();

        public Lot_Variation_Form(string[] Length, int[] Total, Dictionary<string, int>[] Test_Result, int[] Yield, string Path)
        {
            this.Length = Length;
            this.Test_Result = Test_Result;
            this.Yield = Yield;
            this.Total = Total;
            this.Path = Path;
            InitializeComponent();

            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            dataGridView1.Visible = false;

            Gridview();

            dataGridView1.Visible = true;
        }

        public Lot_Variation_Form(string[] Length, int Total_sample, int[] Bin_Total, Dictionary<string, int>[] Test_Result, int[] Yield, Dictionary<string, Data_Class.Data_Editing.Clotho_Spec> For_Bin)
        {
            this.Length = Length;
            this.Test_Result = Test_Result;
            this.Yield = Yield;
            this.Total_sample = Total_sample;
            this.Bin_Total = Bin_Total;
            this.For_Bin = For_Bin;
            InitializeComponent();

            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);

            dataGridView1.Visible = false;
            Gridview_for_Bin();

            dataGridView1.Visible = true;
        }

        public void Gridview()
        {
            DataTable _dataTable = new DataTable();
            BindingSource _bindingSource = new BindingSource();

            _bindingSource.DataSource = _dataTable;
            dataGridView1.DataSource = _bindingSource;
            dataGridView1.DoubleBuffereds(true);

            _dataTable.BeginLoadData();

            _dataTable.Columns.Add("No", typeof(int));

            dataGridView1.Columns[0].Width = 40;
            string adsa = "ewqewq";
            int i = 0;
            int g = 1;

            _dataTable.Columns.Add(this.Length[i], typeof(string));
            dataGridView1.Columns[i + g].Width = 400; g++;

            for (i = 0; i < this.Length.Length; i++)
            {
                _dataTable.Columns.Add(this.Length[i] + "_N", typeof(double));
                dataGridView1.Columns[i + g].Width = 70;
            }
            _bindingSource.DataMember = _dataTable.TableName;

            List<Dictionary<string, int>> Update_Dic = new List<Dictionary<string, int>>();

            foreach (Dictionary<string, int> item in Test_Result)
            {
                //   int Yeild = 0;

                var Test = new Dictionary<string, int>();

                foreach (KeyValuePair<string, int> data in item)
                {
                    Test.Add(data.Key.ToString(), data.Value);
                }
                //   Dictionary<string, int> Dsec = Test.OrderByDescending(num => num.Value).ToDictionary(t => t.Key, t => t.Value);
                Update_Dic.Add(Test);
            }

            object[] value = new object[Length.Length + 2];

            int k = 0;
            int Start_Index = 0;
            int For_Dic_Index = 0;

            for (k = 0; k < 4; k++)
            {
                int value_index = 0;
                value[value_index] = k + 1;
                value_index++;

                if (k == 0)
                {
                    value[value_index] = "Total"; value_index++;
                    for (int j = 0; j < Length.Length; j++)
                    {
                      
                        value[value_index] = this.Total[j]; value_index++;
                    }

                }
                else if (k == 1)
                {
                    value[value_index] = "Pass"; value_index++;
                    for (int j = 0; j < Length.Length; j++)
                    {
                        value[value_index] = this.Total[j] - this.Yield[j]; value_index++;
                    }
                }
                else if (k == 2)
                {
                    value[value_index] = "Fail"; value_index++;
                    for (int j = 0; j < Length.Length; j++)
                    {
                        value[value_index] = this.Yield[j]; value_index++;
                    }
                }
                else if (k == 3)
                {
                    value[value_index] = "%"; value_index++;

                    for (int j = 0; j < Length.Length; j++)
                    {
                        double Yield = 0f;

                        if (this.Yield[j] == 0)
                        {
                            Yield = 100f;
                        }
                        else
                        {
                            Yield = (Convert.ToDouble(this.Total[j] - this.Yield[j]) / Convert.ToDouble(this.Total[j])) * 100;
                        }

                        Yield = Math.Round(Yield, 5);

                        value[value_index] = Yield; value_index++;
                    }
                }

                _dataTable.Rows.Add(value);
            }




            for (k = k; k < Test_Result[0].Count + 4; k++)
            {
                value = new object[Length.Length + 2];

                value[0] = k + 1;

                int Index = 0;

                value[1] = Test_Result[0].ElementAt(For_Dic_Index).Key;

                for (int j = 0; j < Update_Dic.Count; j++)
                {
                    int Find_Value_Index = 0;
                    Dictionary<string, int> dummy = Update_Dic[j];

                

                    //if (dummy.ElementAt(For_Dic_Index).Value != 0)
                    //{
                    //    value[Start_Index + 1 + j + Find_Value_Index + Index] = dummy.ElementAt(For_Dic_Index).Key; Find_Value_Index++;
                    //    value[Start_Index + 1 + j + Find_Value_Index + Index] = dummy.ElementAt(For_Dic_Index).Value; Find_Value_Index++;
                    //    Index++;
                    //}
                    //else
                    //{

                        value[Index + 2] = dummy.ElementAt(For_Dic_Index).Value;
                        Index++;
                   // }
                }
                For_Dic_Index++;

                _dataTable.Rows.Add(value);
            }

            _dataTable.EndLoadData();
        }

        public void Gridview_for_Bin()
        {
            DataTable _dataTable = new DataTable();
            BindingSource _bindingSource = new BindingSource();

            _bindingSource.DataSource = _dataTable;
            dataGridView1.DataSource = _bindingSource;
            dataGridView1.DoubleBuffereds(true);
            _dataTable.Columns.Add("No", typeof(int));

            dataGridView1.Columns[0].Width = 40;
            int i = 0;
            int g = 1;

            for (i = 0; i < this.Length.Length; i++)
            {
                _dataTable.Columns.Add("Bin" + this.Length[i], typeof(string));
                _dataTable.Columns.Add("Bin" + this.Length[i] + "_Min", typeof(string));
                _dataTable.Columns.Add("Bin" + this.Length[i] + "_Max", typeof(string));
                _dataTable.Columns.Add("Bin" + this.Length[i] + "_N", typeof(int));


                dataGridView1.Columns[i + g].Width = 400; g++;
                dataGridView1.Columns[i + g].Width = 70; g++;
                dataGridView1.Columns[i + g].Width = 70; g++;
                dataGridView1.Columns[i + g].Width = 50;
            }
            _bindingSource.DataMember = _dataTable.TableName;

            List<Dictionary<string, int>> Update_Dic = new List<Dictionary<string, int>>();

            foreach (Dictionary<string, int> item in Test_Result)
            {

                var Test = new Dictionary<string, int>();

                foreach (KeyValuePair<string, int> data in item)
                {
                    Test.Add(data.Key.ToString(), data.Value);
                }
                Dictionary<string, int> Dsec = Test.OrderByDescending(num => num.Value).ToDictionary(t => t.Key, t => t.Value);
                Update_Dic.Add(Dsec);
            }

            object[] value = new object[Length.Length * 4 + 1];

            int k = 0;
            int Start_Index = 0;
            int For_Dic_Index = 0;

            for (k = 0; k < 4; k++)
            {
                int value_index = 0;
                value[value_index] = k + 1;
                value_index++;

                if (k == 0)
                {
                    for (int j = 0; j < Length.Length; j++)
                    {
                        value[value_index] = "Total"; value_index++;
                        value[value_index] = null; value_index++;
                        value[value_index] = null; value_index++;
                        value[value_index] = Total_sample; value_index++;
                    }

                }
                else if (k == 1)
                {
                    for (int j = 0; j < Length.Length; j++)
                    {
                        value[value_index] = "Pass"; value_index++;
                        value[value_index] = null; value_index++;
                        value[value_index] = null; value_index++;
                        value[value_index] = Total_sample - this.Bin_Total[j]; value_index++;
                    }
                }
                else if (k == 2)
                {
                    for (int j = 0; j < Length.Length; j++)
                    {
                        value[value_index] = "Fail"; value_index++;
                        value[value_index] = null; value_index++;
                        value[value_index] = null; value_index++;
                        value[value_index] = this.Bin_Total[j]; value_index++;
                    }
                }
                else if (k == 3)
                {
                    for (int j = 0; j < Length.Length; j++)
                    {
                        double Yield = 0f;

                        value[value_index] = "%"; value_index++;
                        value[value_index] = null; value_index++;
                        value[value_index] = null; value_index++;

                        if (this.Yield[j] == 0)
                        {
                            Yield = 100f;
                        }
                        else
                        {
                            Yield = (Convert.ToDouble(Total_sample - this.Bin_Total[j]) / Convert.ToDouble(Total_sample)) * 100;
                        }

                        Yield = Math.Round(Yield, 4);

                        value[value_index] = Yield; value_index++;
                    }
                }

                _dataTable.Rows.Add(value);
            }

            value = new object[Length.Length * 4 + 1];

            for (k = k; k < 300; k++)
            {
                value[0] = k + 1;

                int Index = 0;
                int Find_Value_Index = 0;
                for (int j = 0; j < Update_Dic.Count; j++)
                {
                    int Count = 0;

                    Dictionary<string, int> dummy = Update_Dic[j];
                    value[Start_Index + 1 + j + Find_Value_Index + Index] = dummy.ElementAt(For_Dic_Index).Key; Find_Value_Index++;
                    value[Start_Index + 1 + j + Find_Value_Index + Index] = For_Bin[dummy.ElementAt(For_Dic_Index).Key].Min[Count]; Find_Value_Index++;
                    value[Start_Index + 1 + j + Find_Value_Index + Index] = For_Bin[dummy.ElementAt(For_Dic_Index).Key].Max[Count]; Find_Value_Index++;
                    value[Start_Index + 1 + j + Find_Value_Index + Index] = dummy.ElementAt(For_Dic_Index).Value;
                    Count++;
                }
                For_Dic_Index++;

                _dataTable.Rows.Add(value);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {


            CSV.Open("YIELD");

            Csv_Interface = CSV.Open("YIELD");

            
            string Filename = Path.Substring(Path.LastIndexOf("\\") + 1);

            int Length = Filename.Length;

            Filename = Path.Substring(0, Path.Length - Length);

            Filename += "Result_" + System.DateTime.Now.ToString("yyyyMMddhhmmss") + ".csv";

         //   Filename = "C:\\1.csv";
            Csv_Interface.Write_Open(Filename);

            string dummy = "";
            for (int i = 0; i < this.Length.Length + 1; i++)
            {
                if (i == 0)
                {
                    dummy += "No,,";
                }
                else if (i == this.Length.Length)
                {
     
                    dummy += this.Length[i - 1];
                }
                else
                {
                    dummy += this.Length[i - 1] + ",";
            
                }
            }

            Csv_Interface.Write(dummy);



            for (int k = 0; k < Test_Result[0].Count + 4; k++)
            {
                dummy = "";

                for (int i = 0; i < this.Length.Length + 2; i++)
                {
                    if (i == this.Length.Length + 2)
                    {
                        dummy += dataGridView1.Rows[k].Cells[i].Value.ToString();
                    }
                    else
                    {
                        dummy += dataGridView1.Rows[k].Cells[i].Value.ToString() + ",";
                    }

                }

                Csv_Interface.Write(dummy);
            }

            Csv_Interface.Write_Close();
        }
    }

    public static class ExtensionMethodss
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.SetProperty);
            pi.SetValue(dgv, setting, null);
        }
    }
}
