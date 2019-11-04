using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace TestApplication
{
    public partial class Get_Spec_Form : Form
    {
        int RowCount = 0;
        int ColumnCount = 0;

        public Get_Spec_Form()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {

            CSV_Class.CSV.INT CSV_Interface;
            CSV_Class.CSV CSV = new CSV_Class.CSV();

            Data_Class.Data_Editing.INT Data_Interface;
            Data_Class.Data_Editing Data_Edit = new Data_Class.Data_Editing();

            EXCEL_Class.Excel_Editing Excel = new EXCEL_Class.Excel_Editing();
            EXCEL_Class.Excel_Editing.INT EXCEl_Interface;

            DB_Class.DB_Editing DB = new DB_Class.DB_Editing();
            DB_Class.DB_Editing.INT DB_Interface;

            string Key = "GETSPEC";

            CSV_Interface = CSV.Open(Key);
            Data_Interface = Data_Edit.Open(Key);
            EXCEl_Interface = Excel.Open(Key);
            DB_Interface = DB.Open(Key);
            string Files1 = textBox4.Text;
            Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

            // string Files1 = "C:\\Users\\hyeongck\\Desktop\\New folder\\Program_tool\\New\\New folder (2)\\AFEM8105_A2A_PROD_REV12.CSV";

            CSV_Interface.Read_Open(Files1);


            int m = 0;
            Data_Interface.Clotho_Spec_Data = new string[40000];
            while (!CSV_Interface.StreamReader.EndOfStream)
            {
                Data_Interface.Clotho_Spec_Data[m] = CSV_Interface.Read_Cloth_Spec();
                m++;
            }
            var Var = Data_Interface.Clotho_Spec_Data;
            Array.Resize(ref Var, m);
            Data_Interface.Clotho_Spec_Data = Var;
            Var = null;

            Data_Interface.Find_Cloth_DataFile(Data_Interface.Clotho_Spec_Data);
            Data_Interface.Reference_Header = Data_Interface.Ref_New_Header;


            //while (!CSV_Interface.StreamReader.EndOfStream)
            //{
            //    CSV_Interface.Read();
            //    bool Flag = Data_Interface.Find_First_Row(CSV_Interface.Get_String);
            //    if (Flag) break;
            //}

            Data_Interface.Define_DB_Count(CSV_Interface.Get_String);

            //Data_Interface.For_GetSpec_Header = new string[Data_Interface.Reference_Header.Length];

            //for (int i = 0; i < Data_Interface.Reference_Header.Length; i++)
            //{
            //    string[] dummy = Data_Interface.Reference_Header[i].Split('_');
            //    string dummy_string = dummy[dummy.Length - 1].Replace('-', '_');
            //    Data_Interface.For_GetSpec_Header[i] = dummy_string;

            //}

            string Files2 = textBox1.Text;

            Files2 = Files2.Replace('\r', ' ').Replace('\n', ' ').Trim();
            //  Files2 = "C:\\Users\\hyeongck\\Desktop\\New folder\\Program_tool\\New\\New folder (2)\\AFEM8105_A2A_Rev0p6.xlsx";

            EXCEl_Interface.Open_Session1(Files2, "", false);

            RowCount = EXCEl_Interface.Get_Row_Count("Spec_Sheet_Band");
            ColumnCount = EXCEl_Interface.Get_Column_Count("Spec_Sheet_Band");

            Data_Interface.Band = new string[50];
            Data_Interface.Spec_Band = new Dictionary<string, Dictionary<string, string>>();

            for (int i = 1; i < RowCount + 1; i++)
            {
                object[,] ExcelData = EXCEl_Interface.Read("Spec_Sheet_Band", i);
                Data_Interface.TestPlanAddDic(ExcelData, i);
            }

            EXCEl_Interface.Close();

            string PW = textBox3.Text;
            string Files3 = textBox2.Text;

            Files3 = Files3.Replace('\r', ' ').Replace('\n', ' ').Trim();

            EXCEl_Interface.Open_Session2(Files3, PW, false);

            Data_Interface.Dic_Spec = new Dictionary<string, Data_Class.Data_Editing.Spec>();

            for (int i = 1; i < Data_Interface.Reference_Header.Length; i++)
            {
                Data_Class.Data_Editing.Spec Spec = new Data_Class.Data_Editing.Spec("-999", "999", "999", "", "", "", 0, 0);
                Data_Interface.Dic_Spec.Add(Data_Interface.Reference_Header[i], Spec);
            }

            foreach (KeyValuePair<string, Dictionary<string, string>> Spec_Num in Data_Interface.Spec_Band)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                RowCount = EXCEl_Interface.Get_Row_Count2(Spec_Num.Key.ToString());
                ColumnCount = EXCEl_Interface.Get_Column_Count2(Spec_Num.Key.ToString());

                Dictionary<string, string> Dummy = Data_Interface.Spec_Band[Spec_Num.Key.ToString()];
                EXCEl_Interface.SelectSheet2(Spec_Num.Key.ToString());

                bool ConvertFlag = false;
                int ConvertIndex = 0;
                object[,] Column = EXCEl_Interface.Read2_Range(Convert.ToInt16(Dummy["START_POSITION"]), 1, Convert.ToInt16(Dummy["START_POSITION"]), ColumnCount, Spec_Num.Key.ToString());

                for (int j = 1; j <= ColumnCount; j++)
                {
                    if (Column[1, j] != null)
                    {
                        if (Column[1, j].ToString().ToUpper() == "CONVERT")
                        {
                            ConvertFlag = true;
                            ConvertIndex = j;
                        }
                    }
                }
                object[,] Data = EXCEl_Interface.Read2_Range(Convert.ToInt16(Dummy["START_POSITION"]) + 1, 1, RowCount, ColumnCount, Spec_Num.Key.ToString());

                double test1 = TestTime1.Elapsed.TotalMilliseconds;


                for (int i = 1; i <= RowCount - 3; i++)
                {
                    Stopwatch TestTime2 = new Stopwatch();
                    TestTime2.Restart();
                    TestTime2.Start();

                    if (Data[i, 1] != null)
                    {
                        string Min = "";
                        string Max = "";
                        string Typical = "";
                        string GBT = "";


                        int Both = 0;

                        if (Data[i, Convert.ToInt16(Dummy["SPEC_MIN_POSITION"])] != null && Data[i, Convert.ToInt16(Dummy["SPEC_MAX_POSITION"])] == null)
                        {
                            Min = Data[i, Convert.ToInt16(Dummy["SPEC_MIN_POSITION"])].ToString();
                            Both = 1;
                        }
                        if (Data[i, Convert.ToInt16(Dummy["SPEC_MAX_POSITION"])] != null && Data[i, Convert.ToInt16(Dummy["SPEC_MIN_POSITION"])] == null)
                        {
                            Max = Data[i, Convert.ToInt16(Dummy["SPEC_MAX_POSITION"])].ToString();
                            Both = 2;
                        }
                        if (Data[i, Convert.ToInt16(Dummy["SPEC_MIN_POSITION"])] != null && Data[i, Convert.ToInt16(Dummy["SPEC_MAX_POSITION"])] != null)
                        {
                            Min = Data[i, Convert.ToInt16(Dummy["SPEC_MIN_POSITION"])].ToString();
                            Max = Data[i, Convert.ToInt16(Dummy["SPEC_MAX_POSITION"])].ToString();
                            Both = 3;
                        }

                        if (Data[i, Convert.ToInt16(Dummy["TYPICAL"])] != null)
                        {
                            Typical = Data[i, Convert.ToInt16(Dummy["TYPICAL"])].ToString();
                            Both = 4;
                        }



                        if (Data[i, Convert.ToInt16(Dummy["COMPLIANCE"])] != null)
                        {

                            GBT = Data[i, Convert.ToInt16(Dummy["COMPLIANCE"])].ToString().Replace("\n", "").Replace("/", "").TrimEnd();
                        }


                        if (ConvertFlag && Data[i, ConvertIndex] != null)
                        {
                            Data_Interface.Find_Para_by_Defined(Data[i, 1].ToString(), Min, Max, Typical, Data[i, ConvertIndex].ToString(), GBT, ConvertIndex, Both);
                        }
                        else
                        {
                            Data_Interface.Find_Para_by_Defined(Data[i, 1].ToString(), Min, Max, Typical, "", GBT, ConvertIndex, Both);
                        }


                    }
                    double test2 = TestTime2.Elapsed.TotalMilliseconds;
                }

                double test3 = TestTime1.Elapsed.TotalMilliseconds;
            }

            Dir.Dir_Directory Dir = new Dir.Dir_Directory("C:\\Automtion\\SpecLimit");

            Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();
            string NewFilename = Files1.Substring(Files1.LastIndexOf("\\") + 1);

            CSV_Interface.Write_Open("C:\\Automation\\SpecLimit\\" + NewFilename);

            int b = 0;
            foreach (KeyValuePair<string, Data_Class.Data_Editing.Spec> item in Data_Interface.Dic_Spec)
            {
                if (b == 0)
                {
                    CSV_Interface.Write("Parameter", "Min", "Max");
                }
                else
                {

                    if (item.Value.Convert != "")
                    {
                        // double Min = 0f;
                        //  double Max = 0f;

                        if (item.Value.Both == 1)
                        {
                            double Value = Convert_Data(item.Value.Convert, item.Value.Min);

                            CSV_Interface.Write(item.Key.ToString(), Convert.ToString(Value), item.Value.Max + "," + item.Value.Complience);
                        }
                        else if (item.Value.Both == 2)
                        {
                            double Value = Convert_Data(item.Value.Convert, item.Value.Max);

                            CSV_Interface.Write(item.Key.ToString(), item.Value.Min, Convert.ToString(Value) + "," + item.Value.Complience);
                        }
                        else if (item.Value.Both == 3)
                        {
                            double Value1 = Convert_Data(item.Value.Convert, item.Value.Min);
                            double Value2 = Convert_Data(item.Value.Convert, item.Value.Max);

                            CSV_Interface.Write(item.Key.ToString(), Convert.ToString(Value1), Convert.ToString(Value2) + "," + item.Value.Complience);
                        }
                        else if (item.Value.Both == 4)
                        {
                            double Value = Convert_Data(item.Value.Convert, item.Value.Typical);

                            CSV_Interface.Write(item.Key.ToString(), item.Value.Min, Convert.ToString(Value) + "," + item.Value.Complience + ",Typical");
                        }
                        else
                        {
                            CSV_Interface.Write(item.Key.ToString(), item.Value.Min, item.Value.Max + "," + item.Value.Complience);
                        }

                    }
                    else
                    {
                        CSV_Interface.Write(item.Key.ToString(), item.Value.Min, item.Value.Max + "," + item.Value.Complience);
                    }

                }

                b++;
            }


            CSV_Interface.Write_Close();
            EXCEl_Interface.Close2();


        }

        public double Convert_Data(string Value, string Data)
        {
            string First_Value = Value.Substring(0, 1);
            double Return_Value = 0f;
            string[] split;

            switch (First_Value)
            {
                case "/":
                    split = Value.Trim().Split('/');
                    Return_Value = Convert.ToDouble(split[1].Trim());

                    Return_Value = Convert.ToDouble(Data) / Return_Value;
                    break;

                case "*":
                    split = Value.Trim().Split('*');
                    Return_Value = Convert.ToDouble(split[1].Trim());

                    Return_Value = Convert.ToDouble(Data) / Return_Value;
                    break;
            }

            return Return_Value;
        }

        public void Convert_Data(string Value, string Data, string Data1)
        {
            string First_Value = Value.Substring(0, 1);
            double Return_Value = 0f;

            switch (First_Value)
            {
                case "*":
                    string[] split = Value.Split('*');
                    Return_Value = Convert.ToDouble(split[1].Trim());
                    break;
            }

        }

        #region

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] File = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in File)
                {
                    this.textBox1.Text += str + "\r" + "\n";
                }
            }
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }
        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] File = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in File)
                {
                    this.textBox2.Text += str + "\r" + "\n";
                }
            }
        }

        private void textBox2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }

        private void textBox4_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] File = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in File)
                {
                    this.textBox4.Text += str + "\r" + "\n";
                }
            }
        }

        private void textBox4_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }

        #endregion


    }
}
