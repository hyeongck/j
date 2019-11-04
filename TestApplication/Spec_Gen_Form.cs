using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;


namespace TestApplication
{
    public partial class Spec_Gen_Form : Form
    {
        int RowCount = 0;
        int ColumnCount = 0;
        string Description = "";
        string Complience = "";
        string Typical = "";

        CSV_Class.CSV CSV = new CSV_Class.CSV();
        CSV_Class.CSV.INT CSV_Interface;

        Data_Class.Data_Editing Data = new Data_Class.Data_Editing();
        Data_Class.Data_Editing.INT Data_Interface;

        DB_Class.DB_Editing DB = new DB_Class.DB_Editing();
        DB_Class.DB_Editing.INT DB_Interface;

        EXCEL_Class.Excel_Editing Excel = new EXCEL_Class.Excel_Editing();
        EXCEL_Class.Excel_Editing.INT EXCEL_Interface;

        JMP_Class.JMP_Editing.INT JMP_Interface;
        JMP_Class.JMP_Editing JMP = new JMP_Class.JMP_Editing();

        CSV_Class.For_Box Set_Data;
        Dictionary<string, CSV_Class.For_Box> SaveData;

        PPTX_Class.PPTX_Editing.INT PPTX_Interface;
        PPTX_Class.PPTX_Editing PPTX = new PPTX_Class.PPTX_Editing();

        JMP_Class.Script JMP_Script = new JMP_Class.Script();

        string Key = "FCM";
        string[] id;
        string JMP_File;
        string Report_BookName;

        int Row_Offset = 1;
        int Picture_Offset = 0;
        int Value__Count = 0;
        int Spec_Count = 0;
        int Convert_Index = 0;

        Dictionary<string, string[]> Spec_Dic;
        string[] Solted_Para;
        string[] Convert_Char;
        double Convert_Data;
        int Forobject_Row;
        int Forobject_Column;
        object[,] DummyData;
        string[] Spec;
        double[] sdata;
        Dictionary<string, string> dummy;
        string[] FindRow;
        int Loop_coint;

        Dictionary<string, List<string>> Lot_Information;
        Dictionary<string, Dictionary<string, List<string>>> Matching_Lots;
        Dictionary<string, Dictionary<string, List<string>>> information;
        Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>> Matching_Lots_Test;

        List<string> _Lot_Information_Dummy;

        OpenFileDialog Dialog = new OpenFileDialog();
        OpenFileDialog Dialog2 = new OpenFileDialog();
        Dictionary<int, Dictionary<int, string>> Box_Enum = new Dictionary<int, Dictionary<int, string>>();

        string[] Lot = new string[0];

        int PPTX_Count = 1;

        public Spec_Gen_Form()
        {
            InitializeComponent();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {

                EXCEL_Interface = Excel.Open(Key);
                Data_Interface = Data.Open(Key);

                string Files1 = textBox1.Text;
                Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

                EXCEL_Interface.Open_Session1(Files1, "", false);

                EXCEL_Interface.Clear_Data("Spec Number");
                EXCEL_Interface.MakeSheet("Spec_Sheet_Band");
                EXCEL_Interface.MakeSheet("Spec Number");

                RowCount = EXCEL_Interface.Get_Row_Count("Spec_Sheet_Band");
                ColumnCount = EXCEL_Interface.Get_Column_Count("Spec_Sheet_Band");

                Data_Interface.Band = new string[50];
                Data_Interface.Spec_Band = new Dictionary<string, Dictionary<string, string>>();

                for (int i = 1; i < RowCount + 1; i++)
                {
                    object[,] ExcelData = EXCEL_Interface.Read("Spec_Sheet_Band", i);
                    Data_Interface.TestPlanAddDic(ExcelData, i);
                }

                string PW = textBox3.Text;

                string Files2 = textBox2.Text;
                Files2 = Files2.Replace('\r', ' ').Replace('\n', ' ').Trim();

                EXCEL_Interface.Open_Session2(Files2, PW, false);

                Write_Spec_Number(Data_Interface, EXCEL_Interface);

                string New = NewFileName2(Files1);

                EXCEL_Interface.SaveAs(New);
                EXCEL_Interface.Close();

                EXCEL_Interface.Open_Session1(New, "", false);

                Get_PA_TestPlanSheet_And_AddValidation("Condition_PA", EXCEL_Interface);
                Get_PA_TestPlanSheet_And_AddValidation("Condition_FBAR", EXCEL_Interface);

                New = NewFileName(New, "_Add_Validation");

                EXCEL_Interface.SaveAs(New);
                EXCEL_Interface.Close();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "")
            //{

            Dialog = new OpenFileDialog();
            JMP_Interface = JMP.Open(Key);

            Dialog.Filter = "DB Files (*.db)| *.db";
            Dialog.InitialDirectory = "C:\\Automation\\DB\\";
            Dialog.Multiselect = true;
            Dialog.ShowDialog();


            string Files1 = "";

            string PW = "";
            string Files2 = "";

            string Files3 = "";


#if DEBUG

            Key = "FCM";
            Files1 = textBox1.Text;
            Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

            PW = textBox3.Text;
            Files2 = textBox2.Text;
            Files2 = Files2.Replace('\r', ' ').Replace('\n', ' ').Trim();

            Files3 = textBox4.Text;
            Files3 = Files3.Replace('\r', ' ').Replace('\n', ' ').Trim();

            //Key = "FCM";
            //Files1 = "C:\\Users\\hyeongck\\Desktop\\AFEM-8100-AP1_RF1_TCF_A7A_Rev04.XLSX";
            //Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

            //PW = "Bane19";
            //Files2 = "C:\\Users\\hyeongck\\Desktop\\Bane_2019_HB_SPAD_Specification_v3.2_released.XLSX";
            //Files2 = Files2.Replace('\r', ' ').Replace('\n', ' ').Trim();

            //Files3 = "C:\\Users\\hyeongck\\Desktop\\AFEM8100_A9A_TSF_Rev12.csv";
            //Files3 = Files3.Replace('\r', ' ').Replace('\n', ' ').Trim();

#else
            Key = "FCM";
            Files1 = textBox1.Text;
            Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

            PW = textBox3.Text;
            Files2 = textBox2.Text;
            Files2 = Files2.Replace('\r', ' ').Replace('\n', ' ').Trim();

            Files3 = textBox4.Text;
            Files3 = Files3.Replace('\r', ' ').Replace('\n', ' ').Trim();
#endif


            Define_Parameter();

            Dir.Dir_Directory Dir = new Dir.Dir_Directory("C:\\Automation\\DB\\FCM");




            CSV_Interface = CSV.Open(Key);
            Data_Interface = Data.Open(Key);
            DB_Interface = DB.Open(Key);
            EXCEL_Interface = Excel.Open(Key);

            CSV_Interface.Read_Open(Files3);


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

            CSV_Interface.Read_Close();

            Data_Interface.Reference_Header = Data_Interface.Ref_New_Header;

            DB_Interface.Open_DB(Dialog.FileNames, Data_Interface);

            //  Data_Interface.TheFirst_Trashes_Header_Count = Convert.ToInt16(DB_Interface.Get_Data_From_Table("INF", "FIRST"));
            //   Data_Interface.TheEnd_Trashes_Header_Count = Convert.ToInt16(DB_Interface.Get_Data_From_Table("INF", "END"));

            Data_Interface.TheFirst_Trashes_Header_Count = 0;
            Data_Interface.TheEnd_Trashes_Header_Count = 0;


            int Count2 = Data_Interface.Clotho_List.Count + Data_Interface.TheFirst_Trashes_Header_Count + Data_Interface.TheEnd_Trashes_Header_Count;
            Data_Interface.Define_DB_Count(Data_Interface.Reference_Header);

         



            Data_Interface.Make_New_header();

            EXCEL_Interface.Open_Session1(Files1, "", false);

            Matching_Lot_data();

            DB_Interface.Matching_Lots = Matching_Lots;

            string Query = "";

            //for (int k = 0; k < 10; k++)
            //{
            //    Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

            //    DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(Data_Interface, Query);
            //}

            //if (DB_Interface.Table_Count == 0) DB_Interface.Table_Count = 1;

            id = new string[0];

            for (int loop = 0; loop < Lot.Length; loop++)
            {
                string QueryTest1 = "Select id from " + Lot[loop] + " where fail not like '1'";
                string[] datas = DB_Interface.Get_Data_By_Query(QueryTest1);

                id = id.Concat(datas).ToArray();


            }


            RowCount = EXCEL_Interface.Get_Row_Count("Spec_Sheet_Band");
            ColumnCount = EXCEL_Interface.Get_Column_Count("Spec_Sheet_Band");

            Data_Interface.Band = new string[50];
            Data_Interface.Spec_Band = new Dictionary<string, Dictionary<string, string>>();

            for (int i = 1; i < RowCount + 1; i++)
            {
                object[,] ExcelData = EXCEL_Interface.Read("Spec_Sheet_Band", i);
                Data_Interface.TestPlanAddDic(ExcelData, i);
            }

            EXCEL_Interface.Close();

            EXCEL_Interface.Open_Session2(Files2, PW, true);

            string NewFilename = Files2.Substring(Files2.LastIndexOf("\\") + 1);

            int Dic_Count = Data_Interface.Spec_Band.Count();

            string Time = DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss");

            Dir = new Dir.Dir_Directory("C:\\Automation\\Result_Test_Spec\\" + NewFilename + "_" + Time);
            string FilePath = "C:\\Automation\\Result_Test_Spec\\" + NewFilename + "_" + Time;


            Report_BookName = EXCEL_Interface.MakeBook_For_Report(true);

            int Report = 0;
            int Sheet_Count = 0;


            PPTX_Interface = PPTX.Opened("FCM");

            PPTX_Interface.Open("C:\\Automation\\PPTX\\PPTX.pptx");
            PPTX_Interface.Slide(1);
            //  PPTX_Interface.Slide(PPTX_Count);

            foreach (KeyValuePair<string, Dictionary<string, string>> Spec_Band_key in Data_Interface.Spec_Band)
            {
                EXCEL_Interface.MakeSheet_For_Report(Spec_Band_key.Key.ToString(), Report);
                RowCount = EXCEL_Interface.Get_Row_Count2(Spec_Band_key.Key.ToString());
                ColumnCount = EXCEL_Interface.Get_Column_Count2(Spec_Band_key.Key.ToString());

                Dictionary<string, string> Dummy = Data_Interface.Spec_Band[Spec_Band_key.Key.ToString()];

                int Count = 1;
                int dummy_Count = 0;
                Report++;

                object[,] data1 = EXCEL_Interface.Read2_Range(Convert.ToInt16(Dummy["START_POSITION"]), 1, Convert.ToInt16(Dummy["START_POSITION"]), ColumnCount, Spec_Band_key.Key.ToString());

                for (int k = 0; k < data1.Length; k++)
                {
                    if (data1[1, k + 1] != null && data1[1, k + 1].ToString().ToUpper() == "CONVERT")
                    {
                        Convert_Index = k;
                        break;
                    }

                }


                if (Sheet_Count >= 0)
                {

                    data1 = EXCEL_Interface.Read2_Range(1, 1, RowCount, ColumnCount, Spec_Band_key.Key.ToString());
                    Row_Offset = 1;
                    Picture_Offset = 0;
                    Spec_Count = 0;

                    for (int i = Convert.ToInt16(Dummy["START_POSITION"]) + 1; i < RowCount; i++)
                    {
                        if (data1[i, 1] != null)
                        {

                            object[,] Dummy_Array = new object[1, ColumnCount];

                            data1[i, 1] = data1[i, 1].ToString().Replace('_', '-');

                      

                            for (int dummy = 0; dummy < ColumnCount; dummy++)
                            {
                                Dummy_Array[0, dummy] = data1[i, dummy + 1];
                            }


                            //   if (data1[i, 1].ToString() == "B1-RX-260")
                            //    {


                            DB_Interface.Values = new Dictionary<string, DB_Class.DB_Editing.Values>();

          

                           DB_Interface.Dic_Test_For_Spec_Gen = new Dictionary<string, CSV_Class.For_Box>();

                            Stopwatch TestTime1 = new Stopwatch();

                            TestTime1.Restart();
                            TestTime1.Start();

                            DB_Interface.Get_Defined_Para(Dummy_Array, Spec_Band_key.Key.ToString(), Data_Interface);


                            double test0 = TestTime1.Elapsed.TotalMilliseconds;

                            //CSV_Class.For_Box Box = new CSV_Class.For_Box(Data.Reference_Header[i], Data1, ID, WAFER_ID, SITE_ID, LOT_ID, 0f, 0f, "0", "0", "", Convert.ToString(Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k].Min[0]), Convert.ToString(Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k].Max[0]), Convert.ToString(Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Min[0]), Convert.ToString(Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Max[0]));
                            //Dic_Test.Add(Data_Interface.Reference_Header[i], Box);

                            if (DB_Interface.Dic_Test_For_Spec_Gen.Count != 0)
                            {
                                string KeyTest = Spec_Band_key.Key.ToString().Trim();

                               // PPTX_Interface.Title(Dummy_Array[0, 0].ToString(), Spec_Band_key.Key.ToString(), 40, 1);


                                if (Dummy_Array[0, Convert.ToInt16(Dummy["DESCRIPTION"]) - 1] != null)
                                {

                                    Description = Dummy_Array[0, Convert.ToInt16(Dummy["DESCRIPTION"]) - 1].ToString().Replace("\n", "").Replace("/", "").TrimEnd();
                                }

                                if (Dummy_Array[0, Convert.ToInt16(Dummy["COMPLIANCE"]) - 1] != null)
                                {

                                    Complience = Dummy_Array[0, Convert.ToInt16(Dummy["COMPLIANCE"]) - 1].ToString().Replace("\n", "").Replace("/", "").TrimEnd();
                                }


                                Dir = new Dir.Dir_Directory(FilePath + "\\" + KeyTest + "\\" + Description + "\\");


                                TestTime1 = new Stopwatch();

                                TestTime1.Restart();
                                TestTime1.Start();
                                Write_FCM_Data(Dummy_Array, DB_Interface.Values, Spec_Band_key.Key.ToString(), Count, dummy_Count, EXCEL_Interface, DB_Interface, Data_Interface, FilePath);

                                double test1 = TestTime1.Elapsed.TotalMilliseconds;
                        
                                //     PPTX_Interface.Slide(PPTX_Count);
                                //     PPTX_Count++;


                                dummy_Count++;
                            }
                        }
                        //   }
                        Count++;
                    }
                }
                Count = 0;
                Sheet_Count++;
            }

            EXCEL_Interface.Close2();
            DB_Interface.Close(Data_Interface);
            //     }

        }

        private void Write_Spec_Number(Data_Class.Data_Editing.INT Data_Interface, EXCEL_Class.Excel_Editing.INT EXCEl_Interface)
        {
            int Count_For_BandSpec_Column = 1;

            foreach (KeyValuePair<string, Dictionary<string, string>> Spec_Band_key in Data_Interface.Spec_Band)
            {
                int Count_For_BandSpec_Row = 1;

                RowCount = EXCEl_Interface.Get_Row_Count2(Spec_Band_key.Key.ToString());
                object[,] Excel_Data = EXCEl_Interface.Read2_ColumnbyColumn(Spec_Band_key.Key.ToString(), RowCount, Convert.ToInt16(Spec_Band_key.Value["START_POSITION"]) + 1, 1);
                object[,] DummyData = new object[RowCount - (Convert.ToInt16(Spec_Band_key.Value["START_POSITION"]) - 2), 1];
                DummyData[0, 0] = Spec_Band_key.Key.ToString();

                for (int i = 0; i < RowCount - Convert.ToInt16(Spec_Band_key.Value["START_POSITION"]); i++)
                {
                    if (Excel_Data[Count_For_BandSpec_Row, 1] != null && Excel_Data[Count_For_BandSpec_Row, 1].ToString().Contains("_"))
                    {
                        DummyData[Count_For_BandSpec_Row, 0] = Excel_Data[Count_For_BandSpec_Row, 1].ToString();
                        Count_For_BandSpec_Row++;
                    }
                    else
                    {
                        DummyData[Count_For_BandSpec_Row, 0] = "";
                        Count_For_BandSpec_Row++;
                    }
                }

                DummyData[Count_For_BandSpec_Row, 0] = "V";
                Count_For_BandSpec_Row++;

                DummyData = Resize(DummyData, Count_For_BandSpec_Row);
                EXCEl_Interface.SelectSheet("Spec Number");
                EXCEl_Interface.Write_Array(1, Count_For_BandSpec_Column, DummyData.Length, Count_For_BandSpec_Column, DummyData);
                Count_For_BandSpec_Column++;
            }
        }

        private void Write_FCM_Data(object[,] SpecSheetData, Dictionary<string, DB_Class.DB_Editing.Values> Values, string Spec_Num, int Count_For_SpecSheet, int Dummy_Count, EXCEL_Class.Excel_Editing.INT EXCEL_Interface, DB_Class.DB_Editing.INT DB_Interface, Data_Class.Data_Editing.INT Data_Interface, string FilePath)
        {
            Stopwatch TestTime = new Stopwatch();
            TestTime.Restart();
            TestTime.Start();

            int Data_lentgh = 0;
            int ParaandSpec_offset_Row = 5;

            EXCEL_Interface.MakeBook(true);


            double test0 = TestTime.Elapsed.TotalMilliseconds;

            Forobject_Row = 0;
            Forobject_Column = 0;

            DummyData = new object[1, 1];

            foreach (KeyValuePair<string, CSV_Class.For_Box> item in DB_Interface.Dic_Test_For_Spec_Gen)
            {
                DummyData = new object[5 + item.Value.ID.Length, DB_Interface.Dic_Test_For_Spec_Gen.Count + 1];
                DummyData[0, Forobject_Column] = "Parameter";
                DummyData[1, Forobject_Column] = "Broadcom_HighL";
                DummyData[2, Forobject_Column] = "Broadcom_LowL";
                DummyData[3, Forobject_Column] = "Apple_HighL";
                DummyData[4, Forobject_Column] = "Apple_LowL";

                int i = 0;
                for (i = 0; i < item.Value.ID.Length; i++)
                {
                    DummyData[i + 5, Forobject_Column] = Convert.ToString(i + 1);
                }

                break;
            }
            Forobject_Column = 1;

            this.dummy = new Dictionary<string, string>();

            Spec_Dic = new Dictionary<string, string[]>();
            Spec = new string[4];
            FindRow = new string[DB_Interface.Dic_Test_For_Spec_Gen.Count];
            Loop_coint = 0;

            foreach (KeyValuePair<string, CSV_Class.For_Box> item in DB_Interface.Dic_Test_For_Spec_Gen)
            {
                Spec = new string[4];

                object dummyTestData = null;

                if (SpecSheetData[0, Convert_Index] != null)
                {
                    Convert_Char = SpecSheetData[0, Convert_Index].ToString().Split(',');

                    //  DummyData[Forobject_Row, Forobject_Column] = item.Key.ToString(); Forobject_Row++;

                    sdata = new double[0];
                    Convert_Data = 0;

                    if (Convert_Char.Length == 1)
                    {
                        #region

                        MovetoSpecNone(item, Spec_Num, SpecSheetData);

                        for (int h = 0; h < Convert_Char.Length; h++)
                        {
                            if (Convert_Char[h].ToUpper() == "B-L")
                            {
                                MovetoSepecB_L(item, Spec_Num, SpecSheetData, false);
                            }
                            else if (Convert_Char[h].ToUpper() == "C-L")
                            {
                                MovetoSepecC_L(item, Spec_Num, SpecSheetData, false);
                            }
                            else if (Convert_Char[h].Contains("*"))
                            {
                                Mul(item, Spec_Num, SpecSheetData, Convert_Char[h]);
                            }
                            else if (Convert_Char[h].Contains("/"))
                            {
                                Divide(item, Spec_Num, SpecSheetData, Convert_Char[h]);
                            }
                        }

                       
                     //   sdata = Array.ConvertAll<object, double>(item.Value.data, Convert.ToDouble);
                        for (int n = 0; n < item.Value.data.Length; n++)
                        {
                            item.Value.data[n] = item.Value.data[n] * Convert_Data;
                        }

                     //   item.Value.data = Array.ConvertAll<double, string>(sdata, Convert.ToString);

                        #endregion
                    }
                    else if (Convert_Char.Length == 2)
                    {
                        #region
                        MovetoSpecNone(item, Spec_Num, SpecSheetData);
                        for (int h = 0; h < Convert_Char.Length; h++)
                        {
                            if (Convert_Char[h].ToUpper() == "B-L")
                            {
                                MovetoSepecB_L(item, Spec_Num, SpecSheetData, true);
                            }
                            else if (Convert_Char[h].ToUpper() == "B-H")
                            {
                                MovetoSepecB_H(item, Spec_Num, SpecSheetData, true);
                            }
                            else if (Convert_Char[h].ToUpper() == "C-L")
                            {
                                MovetoSepecC_L(item, Spec_Num, SpecSheetData, true);
                            }
                            else if (Convert_Char[h].Contains("*"))
                            {
                                Mul(item, Spec_Num, SpecSheetData, Convert_Char[h]);
                            }
                            else if (Convert_Char[h].Contains("/"))
                            {
                                Divide(item, Spec_Num, SpecSheetData, Convert_Char[h]);
                            }
                        }


                        //   sdata = Array.ConvertAll<object, double>(item.Value.data, Convert.ToDouble);
                        for (int n = 0; n < item.Value.data.Length; n++)
                        {
                            item.Value.data[n] = item.Value.data[n] * Convert_Data;
                        }

                        //   item.Value.data = Array.ConvertAll<double, string>(sdata, Convert.ToString);
                        #endregion
                    }
                    else if (Convert_Char.Length == 3)
                    {
                        #region
                        for (int h = 0; h < Convert_Char.Length; h++)
                        {
                            if (Convert_Char[h].ToUpper() == "B-L")
                            {
                                MovetoSepecB_L(item, Spec_Num, SpecSheetData, false);
                            }
                            else if (Convert_Char[h].ToUpper() == "C-L")
                            {
                                MovetoSepecC_L(item, Spec_Num, SpecSheetData, false);
                            }
                            else if (Convert_Char[h].Contains("*"))
                            {
                                Mul(item, Spec_Num, SpecSheetData, Convert_Char[h]);
                            }
                            else if (Convert_Char[h].Contains("/"))
                            {
                                Divide(item, Spec_Num, SpecSheetData, Convert_Char[h]);
                            }
                        }


                        //  sdata = Array.ConvertAll<object, double>(item.Value.data, Convert.ToDouble);
                        for (int n = 0; n < item.Value.data.Length; n++)
                        {
                            item.Value.data[n] = item.Value.data[n] * Convert_Data;
                        }

                        //    item.Value.data = Array.ConvertAll<double, string>(sdata, Convert.ToString);

                        #endregion

                    }
                }
                else
                {
                    #region




                    DummyData[Forobject_Row, Forobject_Column] = item.Key.ToString(); Forobject_Row++;
                    DummyData[Forobject_Row, Forobject_Column] = item.Value.Broadcom_Spec_Max.ToString(); Forobject_Row++;
                    DummyData[Forobject_Row, Forobject_Column] = item.Value.Broadcom_Spec_Min.ToString(); ; Forobject_Row++;


                    dummy = Data_Interface.Spec_Band[Spec_Num];

                    Spec[0] = item.Value.Broadcom_Spec_Min.ToString();
                    Spec[1] = item.Value.Broadcom_Spec_Max.ToString();
                    Spec[2] = "";
                    Spec[3] = "";

                    dummyTestData = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1];
                    if (dummyTestData != null)
                    {
                        DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString(); Forobject_Row++;
                        Spec[3] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString();
                    }
                    else
                    {
                        DummyData[Forobject_Row, Forobject_Column] = "999"; Forobject_Row++;
                        Spec[3] = "999";

                    }

                    dummyTestData = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1];

                    if (dummyTestData != null)
                    {
                        DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString(); Forobject_Row++;
                        Spec[2] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString();
                    }
                    else
                    {
                        DummyData[Forobject_Row, Forobject_Column] = "-999"; Forobject_Row++;
                        Spec[2] = "-999";
                    }

                    item.Value.Apple_Spec_Max = Spec[3];
                    item.Value.Apple_Spec_Min = Spec[2];
                    //dummyTestData = SpecSheetData[0, Convert.ToInt16(dummy["TYPICAL"]) - 1];

                    //if (dummyTestData != null)
                    //{
                    //    DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["TYPICAL"]) - 1].ToString(); Forobject_Row++;
                    //    Spec[3] = SpecSheetData[0, Convert.ToInt16(dummy["TYPICAL"]) - 1].ToString();
                    //}
                    //else
                    //{
                    //    DummyData[Forobject_Row, Forobject_Column] = "999"; Forobject_Row++;
                    //    Spec[3] = "-999";
                    //}
                    #endregion

                }


                Spec_Dic.Add(item.Key.ToString(), Spec);


                for (int i = 0; i < item.Value.data.Length; i++)
                {
                    DummyData[Forobject_Row, Forobject_Column] = item.Value.data[i].ToString(); Forobject_Row++;
                }

                Forobject_Column++;
                Forobject_Row = 0;
                Data_lentgh = item.Value.data.Length + ParaandSpec_offset_Row;
                Loop_coint++;
            }

            double test1 = TestTime.Elapsed.TotalMilliseconds;

            EXCEL_Interface.Write_Array(1, 1, Data_lentgh, DB_Interface.Dic_Test_For_Spec_Gen.Count + 1, DummyData);

            double test2 = TestTime.Elapsed.TotalMilliseconds;

            object[,] ReportData = new object[DB_Interface.Dic_Test_For_Spec_Gen.Count + 3, 14];

            ReportData[0, 0] = "SpecNumber";
            ReportData[0, 1] = SpecSheetData[0, 0];
            ReportData[0, 2] = "";
            ReportData[0, 3] = Complience;
            ReportData[0, 4] = "";
            ReportData[0, 5] = "";
            ReportData[0, 6] = "";
            ReportData[0, 7] = "";
            ReportData[0, 8] = "";
            ReportData[0, 9] = "";
            ReportData[0, 10] = "";
            ReportData[0, 11] = "";
            ReportData[0, 12] = "";
            ReportData[0, 13] = "";

            ReportData[1, 0] = "";
            ReportData[1, 1] = "";
            ReportData[1, 2] = "";
            ReportData[1, 3] = "";
            ReportData[1, 4] = "";
            ReportData[1, 5] = "";
            ReportData[1, 6] = "";
            ReportData[1, 7] = "";
            ReportData[1, 8] = "";
            ReportData[1, 9] = "";
            ReportData[1, 10] = "";
            ReportData[1, 11] = "";
            ReportData[1, 12] = "";
            ReportData[1, 13] = "";

            ReportData[2, 0] = "";
            ReportData[2, 1] = "Parameter";
            ReportData[2, 2] = "Min";
            ReportData[2, 3] = "Max";
            ReportData[2, 4] = "AVG";
            ReportData[2, 5] = "Std";
            ReportData[2, 6] = "C_CPK";
            ReportData[2, 7] = "B_CPK";
            ReportData[2, 8] = "";
            ReportData[2, 9] = "C_Spec_Min";
            ReportData[2, 10] = "C_Spec_Max";
            ReportData[2, 11] = "";
            ReportData[2, 12] = "B_Spec_Min";
            ReportData[2, 13] = "B_Spec_Max";

            int k = 3;
            for (int i = 0; i < DB_Interface.Dic_Test_For_Spec_Gen.Count; i++)
            {

                ReportData[k, 0] = (i + 1).ToString();
                ReportData[k, 1] = DummyData[0, i + 1];
                ReportData[k, 2] = "Min";
                ReportData[k, 3] = "Max";
                ReportData[k, 4] = "AVG";
                ReportData[k, 5] = "Std";
                ReportData[k, 6] = "C_CPK";
                ReportData[k, 7] = "B_CPK";
                ReportData[k, 8] = "";
                ReportData[k, 9] = DummyData[4, i + 1];
                ReportData[k, 10] = DummyData[3, i + 1];
                ReportData[k, 11] = "";
                ReportData[k, 12] = DummyData[2, i + 1];
                ReportData[k, 13] = DummyData[1, i + 1];
                k++;

            }

            double test3 = TestTime.Elapsed.TotalMilliseconds;

            Forobject_Row = 0;
            Forobject_Column = 0;
            DummyData = new object[6, DB_Interface.Dic_Test_For_Spec_Gen.Count + 1 + 2];

            DummyData[0, Forobject_Column] = "Max";
            DummyData[1, Forobject_Column] = "Min";
            DummyData[2, Forobject_Column] = "AVG";
            DummyData[3, Forobject_Column] = "StdDev";
            DummyData[4, Forobject_Column] = "CPK Apple Spec";
            DummyData[5, Forobject_Column] = "CPK Broadcom Spec";

            Forobject_Column++;
            Forobject_Column++;
            ParaandSpec_offset_Row++;

            for (int i = 0; i < DB_Interface.Dic_Test_For_Spec_Gen.Count; i++)
            {
                string columnLetter = ColumnIndexToColumnLetter(Forobject_Column); // returns CV
                DummyData[Forobject_Row, Forobject_Column - 1] = "=MAX(" + columnLetter + ParaandSpec_offset_Row + ":" + columnLetter + Data_lentgh + ")"; Forobject_Row++;
                DummyData[Forobject_Row, Forobject_Column - 1] = "=MIN(" + columnLetter + ParaandSpec_offset_Row + ":" + columnLetter + Data_lentgh + ")"; Forobject_Row++;
                DummyData[Forobject_Row, Forobject_Column - 1] = "=AVERAGE(" + columnLetter + ParaandSpec_offset_Row + ":" + columnLetter + Data_lentgh + ")"; Forobject_Row++;
                DummyData[Forobject_Row, Forobject_Column - 1] = "=STDEV(" + columnLetter + ParaandSpec_offset_Row + ":" + columnLetter + Data_lentgh + ")"; Forobject_Row++;
                DummyData[Forobject_Row, Forobject_Column - 1] = "=MIN((" + columnLetter + "4 -" + columnLetter + (Data_lentgh + 3) + ") / (3 * " + columnLetter + (Data_lentgh + 4) + "), (" + columnLetter + (Data_lentgh + 3) + " - " + columnLetter + "5)/ (3 * " + columnLetter + (Data_lentgh + 4) + "))"; Forobject_Row++;
                DummyData[Forobject_Row, Forobject_Column - 1] = "=MIN((" + columnLetter + "2 -" + columnLetter + (Data_lentgh + 3) + ") / (3 * " + columnLetter + (Data_lentgh + 4) + "), (" + columnLetter + (Data_lentgh + 3) + " - " + columnLetter + "3)/ (3 * " + columnLetter + (Data_lentgh + 4) + "))"; Forobject_Row++;


                Forobject_Column++;
                Forobject_Row = 0;
            }

            int Ref_Data_Length = Data_lentgh;
            int Data_lenght2 = Data_lentgh + 1;
            Data_lentgh = Data_lentgh + 1;


            string columnLetterA = ColumnIndexToColumnLetter(2); // returns CV
            string columnLetterB = ColumnIndexToColumnLetter(1 + Values.Count); // returns CV

            DummyData[Forobject_Row, Forobject_Column - 1] = "=MAX(" + columnLetterA + (Data_lentgh) + ":" + columnLetterB + (Data_lentgh) + ")"; Data_lentgh++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MIN(" + columnLetterA + (Data_lentgh) + ":" + columnLetterB + (Data_lentgh) + ")"; Data_lentgh++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=AVERAGE(" + columnLetterA + (Data_lentgh) + ":" + columnLetterB + (Data_lentgh) + ")"; Data_lentgh++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MAX(" + columnLetterA + (Data_lentgh) + ":" + columnLetterB + (Data_lentgh) + ")"; Data_lentgh++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MIN(" + columnLetterA + (Data_lentgh) + ":" + columnLetterB + (Data_lentgh) + ")"; Data_lentgh++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MIN(" + columnLetterA + (Data_lentgh) + ":" + columnLetterB + (Data_lentgh) + ")"; Data_lentgh++; Forobject_Row++;

            Forobject_Column++;
            Forobject_Row = 0;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MATCH(MAX(" + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + "), " + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + ",0)"; Data_lenght2++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MATCH(MIN(" + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + "), " + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + ",0)"; Data_lenght2++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MATCH(AVERAGE(" + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + "), " + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + ",0)"; Data_lenght2++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "";
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MATCH(MAX(" + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + "), " + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + ",0)"; Data_lenght2++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MATCH(MIN(" + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + "), " + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + ",0)"; Data_lenght2++; Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column - 1] = "=MATCH(MIN(" + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + "), " + columnLetterA + (Data_lenght2) + ":" + columnLetterB + (Data_lenght2) + ",0)"; Data_lenght2++; Forobject_Row++;

            double test4 = TestTime.Elapsed.TotalMilliseconds;

            int RowCount1 = EXCEL_Interface.Get_Row_Count("Sheet1");

            double test5 = TestTime.Elapsed.TotalMilliseconds;

            EXCEL_Interface.Write_Array_Formula(RowCount1 + 1, 1, Ref_Data_Length + 6, DB_Interface.Dic_Test_For_Spec_Gen.Count + 3, DummyData);

            double test6 = TestTime.Elapsed.TotalMilliseconds;

            for (int i = 0; i < 6; i++)
            {
                object[,] Data = EXCEL_Interface.Read("Sheet1", RowCount1 + i + 1);

                if (i == 0)
                {
                    for (int j = 0; j < DB_Interface.Dic_Test_For_Spec_Gen.Count; j++)
                    {
                        ReportData[j + 3, 3] = Data[1, j + 2];
                    }
                }
                else if (i == 1)
                {
                    for (int j = 0; j < DB_Interface.Dic_Test_For_Spec_Gen.Count; j++)
                    {
                        ReportData[j + 3, 2] = Data[1, j + 2];
                    }
                }
                else if (i == 2)
                {
                    for (int j = 0; j < DB_Interface.Dic_Test_For_Spec_Gen.Count; j++)
                    {
                        ReportData[j + 3, 4] = Data[1, j + 2];
                    }
                }
                else if (i == 3)
                {
                    for (int j = 0; j < DB_Interface.Dic_Test_For_Spec_Gen.Count; j++)
                    {
                        ReportData[j + 3, 5] = Data[1, j + 2];
                    }
                }
                else if (i == 4)
                {
                    for (int j = 0; j < DB_Interface.Dic_Test_For_Spec_Gen.Count; j++)
                    {
                        ReportData[j + 3, 6] = Data[1, j + 2];
                    }
                }
                else if (i == 5)
                {
                    for (int j = 0; j < DB_Interface.Dic_Test_For_Spec_Gen.Count; j++)
                    {
                        ReportData[j + 3, 7] = Data[1, j + 2];
                    }
                }

            }
            double test7 = TestTime.Elapsed.TotalMilliseconds;

            EXCEL_Interface.SelectSheet_For_Report(Spec_Num);
            EXCEL_Interface.Write_Array_For_Report(Row_Offset + 1, 1, DB_Interface.Dic_Test_For_Spec_Gen.Count + 3 + Row_Offset, 14, ReportData);

            double test8 = TestTime.Elapsed.TotalMilliseconds;



            object dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 1, DB_Interface.Dic_Test_For_Spec_Gen.Count + 3);
            try
            {
                EXCEL_Interface.Interior("Sheet1", Ref_Data_Length + 1, Convert.ToInt16(dummyData.ToString()) + 1);
            }
            catch
            {

            }


            dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 2, DB_Interface.Dic_Test_For_Spec_Gen.Count + 3);

            try
            {
                EXCEL_Interface.Interior("Sheet1", Ref_Data_Length + 2, Convert.ToInt16(dummyData.ToString()) + 1);
            }
            catch
            {

            }
            dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 4, DB_Interface.Dic_Test_For_Spec_Gen.Count + 3);
            try
            {
                EXCEL_Interface.Interior("Sheet1", Ref_Data_Length + 4, Convert.ToInt16(dummyData.ToString()) + 1);
            }
            catch
            {

            }




            object Worst_Apple_Cpk;
            dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 5, DB_Interface.Dic_Test_For_Spec_Gen.Count + 3);
            Worst_Apple_Cpk = dummyData;

            if (Convert.ToDouble(Worst_Apple_Cpk.ToString()) < 0)
            {
            }
            else
            {
                EXCEL_Interface.Interior("Sheet1", Ref_Data_Length + 5, Convert.ToInt16(dummyData.ToString()) + 1);
            }

            object Worst_Broadcom_Cpk;
            dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 6, DB_Interface.Dic_Test_For_Spec_Gen.Count + 3);
            Worst_Broadcom_Cpk = dummyData;
            if (Convert.ToDouble(Worst_Apple_Cpk.ToString()) < 0)
            {
            }
            else
            {
                EXCEL_Interface.Interior("Sheet1", Ref_Data_Length + 6, Convert.ToInt16(dummyData.ToString()) + 1);
            }


            object Apple_Dummy_Hight;
            object Apple_Dummy_Low;

            Apple_Dummy_Hight = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", 4, 2);
            Apple_Dummy_Low = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", 5, 2);

            object[,] dummyData_for_SpecSheet = new object[1, 9];

            if (Apple_Dummy_Hight.ToString() == "999" && Apple_Dummy_Low.ToString() == "-999")
            {
                dummyData_for_SpecSheet[0, 0] = dummyData = "Y";
                try
                {

                    dummyData_for_SpecSheet[0, 1] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", 3, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                }
                catch { }

                try
                {
                    dummyData_for_SpecSheet[0, 2] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", 2, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                }
                catch { }

                try
                {
                    dummyData_for_SpecSheet[0, 3] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 1, Values.Count + 2);
                }
                catch { }
                try
                {
                    dummyData_for_SpecSheet[0, 4] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 2, Values.Count + 2);
                }
                catch { }

                try
                {
                    dummyData_for_SpecSheet[0, 5] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 3, Values.Count + 2);
                }
                catch { }

                try
                {

                    dummyData_for_SpecSheet[0, 6] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 4, Convert.ToInt16(Worst_Broadcom_Cpk.ToString()) + 1);
                }
                catch { }
                try
                {
                    dummyData_for_SpecSheet[0, 7] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 5, Convert.ToInt16(Worst_Broadcom_Cpk.ToString()) + 1);
                }
                catch { }
                try
                {
                    dummyData_for_SpecSheet[0, 8] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 6, Convert.ToInt16(Worst_Broadcom_Cpk.ToString()) + 1);
                }
                catch { }

            }
            else
            {
                dummyData_for_SpecSheet[0, 0] = dummyData = "Y";

                if (Convert.ToDouble(Worst_Apple_Cpk.ToString()) < 0)
                {
                    dummyData_for_SpecSheet[0, 1] = "";
                    dummyData_for_SpecSheet[0, 2] = "";
                }
                else
                {
                    dummyData_for_SpecSheet[0, 1] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", 3, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                    dummyData_for_SpecSheet[0, 2] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", 2, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);

                }

                dummyData_for_SpecSheet[0, 3] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 1, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                dummyData_for_SpecSheet[0, 4] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 2, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                dummyData_for_SpecSheet[0, 5] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 3, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);

                if (Convert.ToDouble(Worst_Apple_Cpk.ToString()) < 0)
                {
                    dummyData_for_SpecSheet[0, 6] = "";
                    dummyData_for_SpecSheet[0, 7] = "";
                    dummyData_for_SpecSheet[0, 8] = "";
                }
                else
                {
                    dummyData_for_SpecSheet[0, 6] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 4, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                    dummyData_for_SpecSheet[0, 7] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 5, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                    dummyData_for_SpecSheet[0, 8] = dummyData = EXCEL_Interface.Selected_RowandColumn_Read("Sheet1", Ref_Data_Length + 6, Convert.ToInt16(Worst_Apple_Cpk.ToString()) + 1);
                }

            }

            double test9 = TestTime.Elapsed.TotalMilliseconds;

          //  EXCEL_Interface.SelectSheet("Sheet2");

            double test10 = TestTime.Elapsed.TotalMilliseconds;
            // Transpose(Values, Spec);


            //try
            //{
            string[] split = new string[0];

            Dictionary<int, Dictionary<int, string>> OrderbySequence = new Dictionary<int, Dictionary<int, string>>();

            int Paralen = 0;
            bool Falg = true;
            string TextName = "";
            foreach (KeyValuePair<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test_For_Spec_Gen)
            {
                split = test.Key.Split('_');

                int kk = 0;
                foreach (KeyValuePair<int, Dictionary<int, string>> D in Box_Enum)
                {
                    foreach (KeyValuePair<int, string> S in D.Value)
                    {
                        string[] dummy = new string[0];

                        if (S.Value == null)
                        {
                            //   OrderbySequence.Add(Convert.ToInt16(D.Key), D.Value);
                        }
                        else
                        {
                            dummy = S.Value.Split('_');

                            if (kk == 0)
                            {
                                if(Falg)
                                {
                                    Paralen = dummy.Length;

                                    TextName = "";
                                    if (Paralen != 1)
                                    {
                                        TextName = S.Value;
                                    }
                                    Falg = false;
                                }
                          
                                Text = S.Value;
                            }
                            if (dummy.Length == 1)
                            {
                                if (S.Value.ToString().ToUpper() == split[1].ToUpper())
                                {
                                    OrderbySequence.Add(Convert.ToInt16(D.Key), D.Value);
                                }
                            }
                            else
                            {

                                if (S.Value.ToString().ToUpper() == split[1].ToUpper() + "_" + split[2].ToUpper())
                                {
                                    OrderbySequence.Add(Convert.ToInt16(D.Key), D.Value);
                                }
                            }
                        }
                        if (kk == 0)
                            break;
                        kk++;
                    }
                }
                break;
            }

            bool Save_Falg = true;

            if(split[1].Contains("ACLR"))
            {

            }

            EXCEL_Interface.SaveAs(FilePath + "\\" + Spec_Num + "\\" + Description + "\\" + SpecSheetData[0, 0] + "_" + split[1] + ".xlsx");
            EXCEL_Interface.Close();

            double test11 = TestTime.Elapsed.TotalMilliseconds;

            CSV_Interface.Write_Open(FilePath + "\\" + Spec_Num + "\\" + Description + "\\" + SpecSheetData[0, 0] + "_" + split[1] + ".csv");

 
            CSV_Interface.ForBoxplotWrite(SpecSheetData[0, 0].ToString() + "_" + split[1] + ".csv", DB_Interface.ID, DB_Interface.Dic_Test_For_Spec_Gen , split[1]);
            CSV_Interface.Write_Close();


            double test12 = TestTime.Elapsed.TotalMilliseconds;


          

            JMP_Draw_For_Boxplot(FilePath + "\\" + Spec_Num + "\\" + Description + "\\" + SpecSheetData[0, 0] + "_" + split[1] + ".csv", OrderbySequence, Save_Falg, PPTX_Count);
            //}
            double test13 = TestTime.Elapsed.TotalMilliseconds;
            //catch
            //{
            //   // MessageBox.Show("");
            //}

            double Height = (Values.Count + 15) * 17;
            // double Test = (a * (Value__Count + 1)) + 262 + Picture_Offset - 310; 

            if (Spec_Count == 0)
            {
                Height = 0;
            }
            else
            {
                Height = (Row_Offset) * 17;
            }


            float Size_x = 500 + (Values.Count * 30);
            float Size_y = Convert.ToSingle((Values.Count + 15) * 17);

            double test14 = TestTime.Elapsed.TotalMilliseconds;

            //float Left, float Top, float Width = -1F, float Height = -1F);

            try

            {
              

              //  EXCEL_Interface.Insert_Image(FilePath + "\\" + Spec_Num + "\\" + Description + "\\" + SpecSheetData[0, 0] + ".jpg", 800, Convert.ToSingle(Height), Size_x, Size_y);
            }
            catch
            {

            }
            Value__Count = DB_Interface.Dic_Test_For_Spec_Gen.Count;
            //   Picture_Offset += Convert.ToInt16((a * (Value__Count)) + 262);
            Spec_Count++;
            Row_Offset += DB_Interface.Dic_Test_For_Spec_Gen.Count + 5;
            double test15 = TestTime.Elapsed.TotalMilliseconds;

            WriteToSpecSheetData(dummy["DEFINE_SPEC"], Convert.ToInt16(dummy["COMPLIANCE"]) + 3, Count_For_SpecSheet, Dummy_Count, dummyData_for_SpecSheet, EXCEL_Interface);

            double test16 = TestTime.Elapsed.TotalMilliseconds;
        }

        private void WriteToSpecSheetData(string Spec_SheetName, int Start_Position, int RawCount, int Dummy_Count, object[,] dummyData_for_SpecSheet, EXCEL_Class.Excel_Editing.INT EXCEL_Interface)
        {
            int NewCount = RawCount + Start_Position;

            int SheetColumnCount_For_Write_toSpecSheet = 0;

            if (Dummy_Count == 0)
            {
                int SheetColumnCount = 0;
                int ColumnCount = EXCEL_Interface.Get_Column_Count2(Spec_SheetName);

                SheetColumnCount_For_Write_toSpecSheet = SheetColumnCount;

                object[,] dummyData_for_SpecSheet_Header = new object[1, 9];

                dummyData_for_SpecSheet_Header[0, 0] = "ATE";
                dummyData_for_SpecSheet_Header[0, 1] = "ATE Limit(Min)";
                dummyData_for_SpecSheet_Header[0, 2] = "ATE Limit(Max)";
                dummyData_for_SpecSheet_Header[0, 3] = "Max";
                dummyData_for_SpecSheet_Header[0, 4] = "Min";
                dummyData_for_SpecSheet_Header[0, 5] = "Avg";
                dummyData_for_SpecSheet_Header[0, 6] = "StdDev";
                dummyData_for_SpecSheet_Header[0, 7] = "CPK Apple Spec";
                dummyData_for_SpecSheet_Header[0, 8] = "CPK Broadcom Spec";

                //  EXCEL_Interface.Write_Array2(Start_Position, SheetColumnCount_For_Write_toSpecSheet + 2, Start_Position, SheetColumnCount_For_Write_toSpecSheet + 2 + 8, dummyData_for_SpecSheet_Header);
            }

            //   EXCEL_Interface.Write_Array2(Start_Position + RawCount - 1, SheetColumnCount_For_Write_toSpecSheet + 2, Start_Position + RawCount - 1, SheetColumnCount_For_Write_toSpecSheet + 2 + 8, dummyData_for_SpecSheet);

            EXCEL_Interface.Write_Array2(RawCount + 3, Start_Position, RawCount + 3, Start_Position + 8, dummyData_for_SpecSheet);

        }

        private void Transpose(Dictionary<string, DB_Class.DB_Editing.Values> Values, string[] Spec)
        {

            string key = "";

            SaveData = new Dictionary<string, CSV_Class.For_Box>();



            foreach (KeyValuePair<string, DB_Class.DB_Editing.Values> item in Values)
            {
                key = item.Key.ToString();
                break;
            }

            string[] split = key.Split('_');


            int Count = Values[key].Data.Count();
            object[,] DummyData = null;
            int Coulumn = 0;


            if (split[0] == "PT")
            {

                DummyData = new object[Count * Values.Count + 1, 12];
                Coulumn = 12;
                int k = 0;
                int j = 1;


                DummyData[0, 0] = "Label";
                DummyData[0, 1] = "Parameter";
                DummyData[0, 2] = "Band";
                DummyData[0, 3] = "Note";
                DummyData[0, 4] = "Tx";
                DummyData[0, 5] = "Ant";
                DummyData[0, 6] = "Rx";
                DummyData[0, 7] = "Pout";
                DummyData[0, 8] = "Vcc";
                DummyData[0, 9] = "Modulation";
                DummyData[0, 10] = "Freq";
                DummyData[0, 11] = split[1];

                Dictionary<string, int> ForSolt = new Dictionary<string, int>();

                ForSolt.Add("Parameter", 1);
                ForSolt.Add("Band", 3);
                ForSolt.Add("Note", 2);
                ForSolt.Add("Tx", 14);
                ForSolt.Add("Ant", 15);
                ForSolt.Add("Rx", 16);
                ForSolt.Add("Pout", 8);
                ForSolt.Add("Vcc", 10);
                ForSolt.Add("Modulation", 6);
                ForSolt.Add("Freq", 9);


                Solt(Values, ForSolt);


                int for_Count = 0;

                for (for_Count = 0; for_Count < Values.Count; for_Count++)
                // foreach (KeyValuePair<string, DB_Class.DB_Editing.Values> item in Values)
                {
                    DB_Class.DB_Editing.Values dummy_data = Values[Solted_Para[for_Count]];

                    split = Solted_Para[for_Count].Split('_');

                    for (k = 0; k < dummy_data.Data.Length; k++)
                    {
                        DummyData[j, 0] = id[k];
                        DummyData[j, 1] = split[1];
                        DummyData[j, 2] = split[3];
                        DummyData[j, 3] = split[2];
                        DummyData[j, 4] = split[14];
                        DummyData[j, 5] = split[15];
                        DummyData[j, 6] = split[16];
                        DummyData[j, 7] = split[8];
                        DummyData[j, 8] = split[10];
                        DummyData[j, 9] = split[6];
                        DummyData[j, 10] = split[9];
                        DummyData[j, 11] = dummy_data.Data[k];
                        j++;
                    }

                    string[] Specs = Spec_Dic[Solted_Para[for_Count]];

                    for (int w = 0; w < Specs.Length; w++)
                    {
                        if (w == 0 && Specs[0].ToUpper() == "TBD")
                        {
                            Specs[0] = "-999";
                        }
                        else if (w == 1 && Specs[1].ToUpper() == "TBD")
                        {
                            Specs[1] = "999";
                        }
                        else if (w == 2 && Specs[2].ToUpper() == "TBD")
                        {
                            Specs[2] = "-999";
                        }
                        else if (w == 3 && Specs[3].ToUpper() == "TBD")
                        {
                            Specs[3] = "999";
                        }


                    }
                    double[] Dats = Array.ConvertAll<object, double>(dummy_data.Data, Convert.ToDouble);

             //       Set_Data = new CSV_Class.For_Box(Solted_Para[for_Count], null, Dats.Min(), Dats.Max(), "", "", "", Specs[2], Specs[3], Specs[0], Specs[1]);

                    SaveData.Add(Solted_Para[for_Count], Set_Data);


                }
            }
            else if (split[0] == "PR")
            {
                DummyData = new object[Count * Values.Count + 1, 11];
                Coulumn = 11;
                int k = 0;
                int j = 1;

                DummyData[0, 0] = "Label";
                DummyData[0, 1] = "Parameter";
                DummyData[0, 2] = "Band";
                DummyData[0, 3] = "Note";
                DummyData[0, 4] = "Ant";
                DummyData[0, 5] = "Rx";
                DummyData[0, 6] = "Mode";
                DummyData[0, 7] = "Vdd";
                DummyData[0, 8] = "Bias";
                DummyData[0, 9] = "Freq";
                DummyData[0, 10] = split[1];

                Dictionary<string, int> ForSolt = new Dictionary<string, int>();

                ForSolt.Add("Parameter", 1);
                ForSolt.Add("Band", 3);
                ForSolt.Add("Note", 2);
                ForSolt.Add("Ant", 15);
                ForSolt.Add("Rx", 16);
                ForSolt.Add("Mode", 4);
                ForSolt.Add("Bias", 12);
                ForSolt.Add("Idd", 11);
                ForSolt.Add("Freq", 9);


                Solt(Values, ForSolt);

                int for_Count = 0;

                for (for_Count = 0; for_Count < Values.Count; for_Count++)
                {
                    DB_Class.DB_Editing.Values dummy_data = Values[Solted_Para[for_Count]];

                    split = Solted_Para[for_Count].Split('_');

                    if (split[1].ToUpper().Contains("IIP3"))
                    {
                        for (k = 0; k < dummy_data.Data.Length; k++)
                        {
                            DummyData[j, 0] = id[k];
                            DummyData[j, 1] = split[1] + "-" + split[6];

                            DummyData[j, 2] = split[3];
                            DummyData[j, 3] = split[2];
                            DummyData[j, 4] = split[15];
                            DummyData[j, 5] = split[16];
                            DummyData[j, 6] = split[4];
                            DummyData[j, 7] = split[11];
                            DummyData[j, 8] = split[12];
                            DummyData[j, 9] = split[9];
                            DummyData[j, 10] = dummy_data.Data[k];
                            j++;
                        }
                    }
                    else
                    {
                        for (k = 0; k < dummy_data.Data.Length; k++)
                        {
                            DummyData[j, 0] = id[k];
                            DummyData[j, 1] = split[1];
                            DummyData[j, 2] = split[3];
                            DummyData[j, 3] = split[2];
                            DummyData[j, 4] = split[15];
                            DummyData[j, 5] = split[16];
                            DummyData[j, 6] = split[4];
                            DummyData[j, 7] = split[11];
                            DummyData[j, 8] = split[12];
                            DummyData[j, 9] = split[9];
                            DummyData[j, 10] = dummy_data.Data[k];
                            j++;
                        }
                    }


                    string[] Specs = Spec_Dic[Solted_Para[for_Count]];


                    for (int w = 0; w < Specs.Length; w++)
                    {
                        if (w == 0 && Specs[0].ToUpper() == "TBD")
                        {
                            Specs[0] = "-999";
                        }
                        else if (w == 1 && Specs[1].ToUpper() == "TBD")
                        {
                            Specs[1] = "999";
                        }
                        else if (w == 2 && Specs[2].ToUpper() == "TBD")
                        {
                            Specs[2] = "-999";
                        }
                        else if (w == 3 && Specs[3].ToUpper() == "TBD")
                        {
                            Specs[3] = "999";
                        }


                    }

                    double[] Dats = Array.ConvertAll<object, double>(dummy_data.Data, Convert.ToDouble);

                  //  Set_Data = new CSV_Class.For_Box(Solted_Para[for_Count], null, Dats.Min(), Dats.Max(), "", "", "", Specs[2], Specs[3], Specs[0], Specs[1]);
                    SaveData.Add(Solted_Para[for_Count], Set_Data);


                }
            }
            else if (split[0] == "F")
            {
                DummyData = new object[Count * Values.Count + 1, 11];
                Coulumn = 11;
                int k = 0;
                int j = 1;

                DummyData[0, 0] = "Label";
                DummyData[0, 1] = "Parameter";
                DummyData[0, 2] = "Band";
                DummyData[0, 3] = "Start_Freq";
                DummyData[0, 4] = "Stop_Freq";
                DummyData[0, 5] = "Vcc";
                DummyData[0, 6] = "Mode";
                DummyData[0, 7] = "Tx";
                DummyData[0, 8] = "Ant";
                DummyData[0, 9] = "Rx";
                DummyData[0, 10] = split[1];

                Dictionary<string, int> ForSolt = new Dictionary<string, int>();

                ForSolt.Add("Parameter", 1);
                ForSolt.Add("Band", 2);
                ForSolt.Add("Start_Freq", 12);
                ForSolt.Add("Stop_Freq", 13);
                ForSolt.Add("Vcc", 8);
                ForSolt.Add("Mode", 8);
                ForSolt.Add("Tx", 3);
                ForSolt.Add("Ant", 4);
                ForSolt.Add("Rx", 5);


                Solt(Values, ForSolt);

                int for_Count = 0;

                for (for_Count = 0; for_Count < Values.Count; for_Count++)
                {
                    DB_Class.DB_Editing.Values dummy_data = Values[Solted_Para[for_Count]];

                    split = Solted_Para[for_Count].Split('_');

                    for (k = 0; k < dummy_data.Data.Length; k++)
                    {
                        DummyData[j, 0] = id[k];
                        DummyData[j, 1] = split[1];
                        DummyData[j, 2] = split[2];
                        DummyData[j, 3] = split[12];
                        DummyData[j, 4] = split[13];
                        DummyData[j, 5] = split[8];
                        DummyData[j, 6] = split[7];
                        DummyData[j, 7] = split[3];
                        DummyData[j, 8] = split[4];
                        DummyData[j, 9] = split[5];
                        DummyData[j, 10] = dummy_data.Data[k];
                        j++;
                    }
                    string[] Specs = Spec_Dic[Solted_Para[for_Count]];


                    for (int w = 0; w < Specs.Length; w++)
                    {
                        if (w == 0 && Specs[0].ToUpper() == "TBD")
                        {
                            Specs[0] = "-999";
                        }
                        else if (w == 1 && Specs[1].ToUpper() == "TBD")
                        {
                            Specs[1] = "999";
                        }
                        else if (w == 2 && Specs[2].ToUpper() == "TBD" || w == 2 && Specs[2].Contains("G"))
                        {
                            Specs[2] = "-999";
                        }
                        else if (w == 3 && Specs[3].ToUpper() == "TBD" || w == 3 && Specs[3].Contains("G"))
                        {
                            Specs[3] = "999";
                        }

                    }

                    double[] Dats = Array.ConvertAll<object, double>(dummy_data.Data, Convert.ToDouble);

               //     Set_Data = new CSV_Class.For_Box(Solted_Para[for_Count], null, Dats.Min(), Dats.Max(), "", "", "", Specs[2], Specs[3], Specs[0], Specs[1]);

                    SaveData.Add(Solted_Para[for_Count], Set_Data);
                }
            }

            double Count_Forindex = Math.Truncate(Convert.ToDouble(DummyData.Length / (Coulumn * 90000)));



            if (Count_Forindex == 0)
            {
                Count_Forindex = 1;
                EXCEL_Interface.Write_Array(1, 1, Count * Values.Count + 1, Coulumn, DummyData);
            }
            else
            {
                Count_Forindex = Count_Forindex;

                long Total = 0;
                for (int ii = 0; ii < Count_Forindex; ii++)
                {
                    object[,] ArrayCopy = null;

                    if (ii == 0)
                    {
                        ArrayCopy = new object[90000, Coulumn];
                        Array.Copy(DummyData, 0, ArrayCopy, 0, Coulumn * 90000);

                        EXCEL_Interface.Write_Array(1, 1, 90000, Coulumn, ArrayCopy);
                    }
                    else if (ii == Count_Forindex - 1)
                    {
                        double dasd = Convert.ToDouble((DummyData.Length - (90000 * (ii * Coulumn))));

                        ArrayCopy = new object[Convert.ToInt64(dasd) / Coulumn, Coulumn];
                        Array.Copy(DummyData, 90000 * ii * Coulumn, ArrayCopy, 0, Convert.ToInt64(dasd));

                        EXCEL_Interface.Write_Array((90000 * ii) + 1, 1, DummyData.Length / Coulumn, Coulumn, ArrayCopy);

                        Total = (90000 * ii);
                    }
                    else if (ii < Count_Forindex)
                    {
                        ArrayCopy = new object[90000, Coulumn];
                        Array.Copy(DummyData, 90000 * ii * Coulumn, ArrayCopy, 0, Coulumn * 90000);

                        EXCEL_Interface.Write_Array((90000 * ii) + 1, 1, (90000 * (ii + 1)), Coulumn, ArrayCopy);

                        Total = (90000 * ii);
                    }




                }
            }







        }

        public static IEnumerable<string> NaturalSort(IEnumerable<string> list)
        {
            int maxLen = list.Select(s => s.Length).Max();
            Func<string, char> PaddingChar = s => char.IsDigit(s[0]) ? ' ' : char.MaxValue;

            return list
                    .Select(s =>
                        new
                        {
                            OrgStr = s,
                            SortStr = Regex.Replace(s, @"(\d+)|(\D+)", m => m.Value.PadLeft(maxLen, PaddingChar(m.Value)))
                        })
                    .OrderBy(x => x.SortStr)
                    .Select(x => x.OrgStr);
        }


        private void Get_PA_TestPlanSheet_And_AddValidation(string Spec_Sheet_Band, EXCEL_Class.Excel_Editing.INT EXCEl_Interface)
        {
            try
            {
                Dictionary<string, List<string>> GetTCFDefineSpecNum1 = new Dictionary<string, List<string>>();
                List<string> GetTCFDefineSpecNum = new List<string>();
                Dictionary<string, string> Excel_Combobox = new Dictionary<string, string>();

                ColumnCount = EXCEl_Interface.Get_Column_Count("Spec Number");
                RowCount = EXCEl_Interface.Get_Row_Count("Spec Number");

                string columnLetter = "";

                for (int i = 1; i < ColumnCount + 1; i++)
                {
                    columnLetter = ColumnIndexToColumnLetter(i); // returns CV

                    object[,] data = EXCEl_Interface.Read_ColumnbyColumn("Spec Number", 1, RowCount, i);

                    for (int j = 1; j <= data.Length; j++)
                    {
                        if (data[j, 1] != null)
                        {
                            if (j != 1)
                            {
                                GetTCFDefineSpecNum.Add(data[j, 1].ToString());
                            }
                        }
                    }
                    GetTCFDefineSpecNum1.Add(data[1, 1].ToString(), GetTCFDefineSpecNum);
                    Excel_Combobox.Add(data[1, 1].ToString(), "='Spec Number'!$" + columnLetter + "$" + 2 + ":" + "$" + columnLetter + "$" + (GetTCFDefineSpecNum.Count() + 2));
                    GetTCFDefineSpecNum = new List<string>();
                }


                ColumnCount = EXCEl_Interface.Get_Column_Count(Spec_Sheet_Band);
                RowCount = EXCEl_Interface.Get_Row_Count(Spec_Sheet_Band);

                columnLetter = ColumnIndexToColumnLetter(ColumnCount); // returns CV


                int Band_Sel = 0;
                int k = 0;
                int Para_Count = 0;

                bool Break_flag = false;
                bool[] flag = new bool[ColumnCount];

                for (int i = 1; i < RowCount + 1; i++)
                {
                    object[,] data = EXCEl_Interface.Read(Spec_Sheet_Band, i);

                    if (data[1, i] != null)
                    {
                        for (int j = 1; j < data.Length; j++)
                        {
                            if (data[1, j] != null)
                            {
                                if (data[1, j].ToString().Contains("BAND_SPEC"))
                                {
                                    Band_Sel = j;
                                    k = i + 1;

                                    for (int jj = 0; jj < data.Length; jj++)
                                    {
                                        if (data[1, jj + 1] != null)
                                        {
                                            if (data[1, jj + 1].ToString().ToUpper().Contains("PARA."))
                                            {
                                                flag[jj + 1] = true;
                                                Para_Count++;
                                            }
                                        }
                                    }
                                    Break_flag = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (Break_flag)
                        break;
                }

                ColumnCount = EXCEl_Interface.Get_Column_Count(Spec_Sheet_Band);
                RowCount = EXCEl_Interface.Get_Row_Count(Spec_Sheet_Band);

                object[,] data1 = EXCEl_Interface.Read_Range(1, 1, RowCount, ColumnCount, Spec_Sheet_Band);

                for (int i = k; i < RowCount + 1; i++)
                {
                    //object[,] data = Excel_Dll.Excel_Control.Read(i, Spec_Sheet_Band);

                    if (data1[i, Band_Sel] != null && Excel_Combobox.ContainsKey(data1[i, Band_Sel].ToString()))
                    {
                        for (int t = 1; t < ColumnCount; t++)
                        {
                            if (flag[t])
                            {
                                string columnLetter2 = ColumnIndexToColumnLetter(t); // returns CV
                                string columnLetter3 = ColumnIndexToColumnLetter(Para_Count); // returns CV

                                EXCEl_Interface.AddValidation(i, t, i, t + Para_Count - 1, Excel_Combobox[data1[i, Band_Sel].ToString()], Spec_Sheet_Band);
                                break;

                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        private string NewFileName(string File, string AddString)
        {
            string[] Filepath = File.Split('\\');
            string NewFilepath = "";

            for (int i = 0; i < Filepath.Length - 1; i++)
            {
                NewFilepath += Filepath[i] + "\\";
            }

            string NewFilename = File.Substring(File.LastIndexOf("\\") + 1);
            string[] Split_Filename = NewFilename.Split('.');

            for (int i = 0; i < Split_Filename.Length; i++)
            {
                if (i == Split_Filename.Length - 1)
                {
                    if (System.IO.File.Exists(File))
                    {
                        NewFilepath += "- Copy." + Split_Filename[i];
                    }
                    else
                    {
                        NewFilepath += AddString + "." + Split_Filename[i];
                    }

                    break;
                }
                NewFilepath += Split_Filename[i];
            }
            return NewFilepath;
        }

        private string NewFileName2(string File)
        {
            string[] Filepath = File.Split('\\');
            string NewFilepath = "";
            for (int i = 0; i < Filepath.Length - 1; i++)
            {
                NewFilepath += Filepath[i] + "\\";
            }

            string NewFilename = File.Substring(File.LastIndexOf("\\") + 1);
            string[] Split_Filename = NewFilename.Split('.');

            for (int i = 0; i < Split_Filename.Length; i++)
            {
                if (i == Split_Filename.Length - 1)
                {
                    NewFilepath += "_Edit." + Split_Filename[i];
                    break;
                }
                NewFilepath += Split_Filename[i];
            }
            return NewFilepath;
        }

        private object[,] Resize(object[,] original, int Rows)
        {
            var newArray = new object[Rows, 1];
            int minRows = Math.Min(Rows, original.GetLength(0));
            int minCols = Math.Min(1, original.GetLength(1));
            int doummy = 0;
            for (int i = 0; i < minRows; i++)
                for (int j = 0; j < minCols; j++)
                {
                    if (original[i, j].ToString() == "")
                    {
                        doummy--;
                    }
                    else
                    {
                        newArray[doummy, j] = original[i, j];
                    }
                    doummy++;
                }

            newArray = new object[doummy, 1];
            var Neworiginal = newArray;
            doummy = 0;
            for (int i = 0; i < original.Length; i++)
                for (int j = 0; j < minCols; j++)
                {
                    if (original[i, j].ToString() == "")
                    {
                        doummy--;
                    }
                    else
                    {
                        Neworiginal[doummy, j] = original[i, j];
                    }

                    doummy++;
                }

            return newArray;
        }

        private void JMP_Draw_For_Boxplot(string FilePaht, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, int PPTX_Count)
        {
            JMP_Interface.Open_Session(true);

            JMP_File = FilePaht;

          //  JMP_Interface.Open_Document(FilePaht);
         //   JMP_Interface.GetDataTable();

            JMP_Class.Script Distribution_Script;

            Distribution_Script = null;
            List<string>[] Para_Test = new List<string>[OrderbySequence.Count];

            Distribution_Script = JMP_Interface.Make_Script("FCM_VARIABLILITY", DB_Interface.Dic_Test_For_Spec_Gen, null, FilePaht, OrderbySequence, Save_Falg, ref Para_Test,false,false,false,0f);

            string Script = Distribution_Script.Scrip_Data;
            string[] Split = Script.Split('#');



            int Count = 0;
            int index_Test = 0;
            string Naming = "";
            string By = "";
            foreach (KeyValuePair<int, Dictionary<int, string>> _S in OrderbySequence)
            {

                if (_S.Value.ContainsKey(777))
                {
                    string Filename = FilePaht.Substring(FilePaht.LastIndexOf("\\") + 1);
                    int length = FilePaht.Length;
                    Filename = FilePaht.Substring(0, length - Filename.Length);

                    string Name = FilePaht.Substring(0, FilePaht.Length - 4);

                    foreach (KeyValuePair<int, string> _SS in _S.Value)
                    {
                        if (Count == _S.Value.Count - 3)
                        {
                            Naming = _SS.Value;
                        }
                        else if (Count == _S.Value.Count - 1)
                        {
                            Array arr = Enum.GetValues(typeof(BoxPlot));

                            string[] split1 = _SS.Value.Split('>');

                            for (int k = 0; k < split1.Length; k++)
                            {
                                var test = (BoxPlot)Enum.Parse(typeof(BoxPlot), split1[k]);

                                if (k == split1.Length - 1)
                                {
                                    By += test.ToString();
                                }
                                else
                                {
                                    By += test.ToString() + "_";
                                }

                            }

                        }
                        Count++;
                    }


                    JMP_Interface.Path = Name + "_" + Naming + "_By_" + By.ToString();

                    CSV_Interface.Write_Open("C:\\temp\\dummy\\BOXPLOT.jsl");
                    CSV_Interface.WriteScript(Split[index_Test]);
                    CSV_Interface.Write_Close();

                    JMP_Interface.Run_Script("C:\\temp\\dummy\\BOXPLOT.jsl");

                    CSV_Interface.Write_Open(Name + "_" + Naming + "_By_" + By + ".jsl");

                    string[] split = Split[0].Split('\n');
                    string Remake = "";
                    if (index_Test == 0)
                    {
                        Remake = split[0].Substring(0, split[0].Length - 7) + "jrp\");\n";
                        Remake += "dt = current data table();\n";
                        for (int l = 1; l < split.Length; l++)
                        {
                            Remake += split[l];
                        }
                    }
                    else
                    {
                        for (int l = 0; l < split.Length; l++)
                        {
                            Remake += split[l];
                        }
                    }

                    Split[index_Test] = Remake;
                    CSV_Interface.WriteScript(Split[index_Test]);


                    CSV_Interface.Write_Close();


                    index_Test++;
                    Count = 0;
                }
                else
                {

                    Count = 0;
                    string Filename = FilePaht.Substring(FilePaht.LastIndexOf("\\") + 1);
                    int length = FilePaht.Length;
                    Filename = FilePaht.Substring(0, length - Filename.Length);

                    string Name = FilePaht.Substring(0, FilePaht.Length - 4);

                    foreach (KeyValuePair<int, string> _SS in _S.Value)
                    {
                        if (Count == _S.Value.Count - 2)
                        {
                            Naming = _SS.Value;
                        }
                        Count++;
                    }


                    JMP_Interface.Path = Name + "_" + Naming + "_By_" + By;

                    CSV_Interface.Write_Open("C:\\temp\\dummy\\BOXPLOT.jsl");
                    CSV_Interface.WriteScript(Split[index_Test]);
                    CSV_Interface.Write_Close();

                    JMP_Interface.Run_Script("C:\\temp\\dummy\\BOXPLOT.jsl");


                    CSV_Interface.Write_Open(Name + "_" + Naming + ".jsl");


                    string[] split = Split[0].Split('\n');
                    string Remake = "";
                    if (index_Test == 0)
                    {
                        Remake = split[0].Substring(0, split[0].Length - 7) + "jrp\");\n";
                        Remake += "dt = current data table();\n";
                        for (int l = 1; l < split.Length; l++)
                        {
                            Remake += split[l];
                        }
                    }
                    else
                    {
                        for (int l = 0; l < split.Length; l++)
                        {
                            Remake += split[l];
                        }
                    }

                    Split[index_Test] = Remake;
                    CSV_Interface.WriteScript(Split[index_Test]);


                    CSV_Interface.Write_Close();


                    index_Test++;
                    Count = 0;
                }
              
            }


            string CloseDT = "dt = current data table();\n Close(dt, nosave)";
            CSV_Interface.Write_Open("C:\\temp\\dummy\\CloseDT.jsl");
            CSV_Interface.WriteScript(CloseDT);
            CSV_Interface.Write_Close();

            JMP_Interface.Run_Script("C:\\temp\\dummy\\CloseDT.jsl");


            string Rename = FilePaht.Substring(0, FilePaht.Length -4);

         
            float Left = 20;
            float Top = 80;

            float Width = 700;
            float Height1 = 300;

            int index = 0;
            string by = "";

            int PPTX_Slide_Count = 1;


            for (int k = 0; k < Split.Length; k++)
            {
                if (Split.Length == 1)
                {
                    bool flag = false;
                    string Value = "";
                    int dummy = 0;
                    int dummy1 = 0;

                    #region
                    foreach (KeyValuePair<int, Dictionary<int, string>> Data in OrderbySequence)
                    {
                        if (Data.Value.Keys.Contains(777))
                        {
                            #region

                            if (index == dummy)
                            {
                                int Dic_Count = 0;
                                if (Data.Value.Keys.Contains(777))
                                {
                                    Dic_Count = Data.Value.Count - 3;
                                }
                                else
                                {
                                    Dic_Count = Data.Value.Count - 2;
                                }

                                foreach (KeyValuePair<int, string> s in Data.Value)
                                {

                                    if (dummy1 == Dic_Count)
                                    {
                                        if (!flag)
                                        {
                                            Value = s.Value.ToString();
                                            flag = true;

                                        }

                                        // break;
                                    }
                                    else if (dummy1 == Data.Value.Count - 1)
                                    {

                                        for (int kk = 0; kk < Para_Test[dummy].Count; kk++)
                                        {
                                            by = "";

                                            string[] sp = Para_Test[dummy][kk].Split(',');

                                            for (int j = 0; j < sp.Length; j++)
                                            {
                                                if (j == sp.Length - 1)
                                                {
                                                    by += sp[j];
                                                }
                                                else
                                                {
                                                    by += sp[j] + "_";
                                                }
                                            }


                                            string NewFilename = Rename.Substring(Rename.LastIndexOf("\\") + 1);
                                            string path_dummy = Rename + "\\" + NewFilename + "_" + by + ".jpg";

                                            Left = 20;
                                            Top = 80;

                                            Width = 900;
                                            Height1 = 400;


                                            PPTX_Interface.Title(NewFilename.ToUpper() + "_" + Value.ToUpper() + "_" + by, "33", 40, PPTX_Count);
                                            PPTX_Interface.AddPicture(path_dummy, Left, Top, Width, Height1);

                                            PPTX_Count++;


                                        }

                                    }
                                    dummy1++;
                                }

                                if (flag)
                                {
                                    flag = false;
                                    break;
                                }


                            }
                            dummy++;
                            #endregion
                        }
                        else
                        {
                            if (index == dummy)
                            {
                                int Dic_Count = 0;
                                if (Data.Value.Keys.Contains(777))
                                {
                                    Dic_Count = Data.Value.Count - 3;
                                }
                                else
                                {
                                    Dic_Count = Data.Value.Count - 2;
                                }

                                foreach (KeyValuePair<int, string> s in Data.Value)
                                {

                                    if (dummy1 == Dic_Count)
                                    {
                                        if (!flag)
                                        {
                                            Value = s.Value.ToString();
                                            flag = true;

                                        }

                                    }

                                    dummy1++;
                                }


                                string NewFilename = Rename.Substring(Rename.LastIndexOf("\\") + 1);
                                string path_dummy = Rename  + "\\" + NewFilename + "_" + Value + ".jpg";

                                Left = 20;
                                Top = 80;

                                Width = 900;
                                Height1 = 400;

                                //    PPTX_Interface.Slide(1);
                                PPTX_Interface.Title(NewFilename.ToUpper() + "_" + Value, "33", 40, PPTX_Count);
                                PPTX_Interface.AddPicture(path_dummy, Left, Top, Width, Height1);

                                //    PPTX_Interface.Title(Dummy_Array[0, 0].ToString(), Spec_Band_key.Key.ToString(), 40, 1);

                                PPTX_Count++;
                                if (flag)
                                {
                                    flag = false;
                                    break;
                                }
                            }
                            dummy++;
                        }

                    }
                    index++;


                    #endregion
                }
                else
                {
                    bool flag = false;
                    string Value = "";
                    int dummy = 0;
                    int dummy1 = 0;

                    #region
                    foreach (KeyValuePair<int, Dictionary<int, string>> Data in OrderbySequence)
                    {
                        if (Data.Value.Keys.Contains(777))
                        {
                            #region

                            if (index == dummy)
                            {
                                int Dic_Count = 0;
                                if (Data.Value.Keys.Contains(777))
                                {
                                    Dic_Count = Data.Value.Count - 3;
                                }
                                else
                                {
                                    Dic_Count = Data.Value.Count - 2;
                                }

                                foreach (KeyValuePair<int, string> s in Data.Value)
                                {

                                    if (dummy1 == Dic_Count)
                                    {
                                        if (!flag)
                                        {
                                            Value = s.Value.ToString();
                                            flag = true;

                                        }

                                        // break;
                                    }
                                    else if (dummy1 == Data.Value.Count - 1)
                                    {

                                        for (int kk = 0; kk < Para_Test[dummy].Count; kk++)
                                        {
                                            by = "";

                                            string[] sp = Para_Test[dummy][kk].Split(',');

                                            for (int j = 0; j < sp.Length; j++)
                                            {
                                                if (j == sp.Length - 1)
                                                {
                                                    by += sp[j];
                                                }
                                                else
                                                {
                                                    by += sp[j] + "_";
                                                }
                                            }


                                            string NewFilename = Rename.Substring(Rename.LastIndexOf("\\") + 1);
                                            string path_dummy = Rename + "\\" + NewFilename + "_" + by + ".jpg";

                                            Left = 20;
                                            Top = 80;

                                            Width = 900;
                                            Height1 = 400;


                                            PPTX_Interface.Title(NewFilename.ToUpper() + "_" + Value.ToUpper() + "_" + by, "33", 40, PPTX_Count);
                                            PPTX_Interface.AddPicture(path_dummy, Left, Top, Width, Height1);

                                            PPTX_Count++;


                                        }

                                    }
                                    dummy1++;
                                }

                                if (flag)
                                {
                                    flag = false;
                                    break;
                                }


                            }
                            dummy++;
                            #endregion
                        }
                        else
                        {
                            if (index == dummy)
                            {
                                int Dic_Count = 0;
                                if (Data.Value.Keys.Contains(777))
                                {
                                    Dic_Count = Data.Value.Count - 3;
                                }
                                else
                                {
                                    Dic_Count = Data.Value.Count - 2;
                                }

                                foreach (KeyValuePair<int, string> s in Data.Value)
                                {

                                    if (dummy1 == Dic_Count)
                                    {
                                        if (!flag)
                                        {
                                            Value = s.Value.ToString();
                                            flag = true;

                                        }

                                    }

                                    dummy1++;
                                }


                                string NewFilename = Rename.Substring(Rename.LastIndexOf("\\") + 1);
                                string path_dummy = Rename + "_" + Value + "\\" + NewFilename + "_" + Value + ".jpg";

                                Left = 20;
                                Top = 80;

                                Width = 900;
                                Height1 = 400;

                                //    PPTX_Interface.Slide(1);
                                PPTX_Interface.Title(NewFilename.ToUpper() + "_" + Value, "33", 40, PPTX_Count);
                                PPTX_Interface.AddPicture(path_dummy, Left, Top, Width, Height1);

                                //    PPTX_Interface.Title(Dummy_Array[0, 0].ToString(), Spec_Band_key.Key.ToString(), 40, 1);

                                PPTX_Count++;
                                if (flag)
                                {
                                    flag = false;
                                    break;
                                }
                            }
                            dummy++;
                        }

                    }
                    index++;


                    #endregion


                }


                Left += Left + Width;
            }



            //  PPTX_Interface.AddPicture(FilePath + "\\" + Spec_Num + "\\" + Description + "\\" + SpecSheetData[0, 0] + "_Transpose.jpg" , Left, Top, Width, Height1);

        }

        private string[] Solt(Dictionary<string, DB_Class.DB_Editing.Values> Values, Dictionary<string, int> Solt)
        {
            int ww = 0;
            string[] w = Solt.Keys.ToArray();
            Dictionary<long, string> dummy_Test1 = new Dictionary<long, string>();

            foreach (KeyValuePair<string, DB_Class.DB_Editing.Values> item in Values)
            {
                string[] split = item.ToString().Split('_');

                string strTmp = "";
                string st = "";
                for (int sd = 0; sd < Solt.Count; sd++)
                {
                    strTmp = Regex.Replace(split[Solt[w[sd]]], @"\D", "");
                    strTmp = strTmp.PadRight(10, '0');

                    st += split[Solt[w[sd]]] + strTmp + "_";
                }

                dummy_Test1.Add(ww, st);
                ww++;
            }


            var sorted = dummy_Test1.OrderBy(num => num.Value);

            Solted_Para = new string[Values.Count];
            ww = 0;

            string[] dummy_para = Values.Keys.ToArray();
            foreach (KeyValuePair<long, string> item in sorted)
            {
                Solted_Para[ww] = dummy_para[item.Key];
                ww++;
            }

            return null;
        }

        private void MovetoSepecB_L(KeyValuePair<string, CSV_Class.For_Box> item, string Spec_Num, object[,] SpecSheetData, bool Falg)
        {
            //if (!Falg)
            //{
            //    DummyData[Forobject_Row, Forobject_Column] = item.Key.ToString(); Forobject_Row++;
            //    DummyData[Forobject_Row, Forobject_Column] = item.Value.Low.ToString(); FindRow[Loop_coint] = Convert.ToString(Forobject_Row) + "," + Convert.ToString(Forobject_Column); Forobject_Row++;
            //    DummyData[Forobject_Row, Forobject_Column] = item.Value.High.ToString(); Forobject_Row++;
            //}
            //else
            //{

            //    string[] Find_Row_Split = FindRow[Loop_coint].Split(',');

            //    DummyData[Convert.ToInt16(Find_Row_Split[0]) - 1, Convert.ToInt16(Find_Row_Split[1])] = item.Key.ToString();
            //    DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])] = item.Value.Low.ToString();
            //    DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])] = item.Value.High.ToString();
            //}
            //Spec[0] = Convert.ToString(item.Value.High);
            //Spec[1] = Convert.ToString(item.Value.Low);


        }

        private void MovetoSepecB_H(KeyValuePair<string, CSV_Class.For_Box> item, string Spec_Num, object[,] SpecSheetData, bool Falg)
        {
            //if (!Falg)
            //{

            //    DummyData[Forobject_Row, Forobject_Column] = item.Key.ToString(); Forobject_Row++;
            //    DummyData[Forobject_Row, Forobject_Column] = item.Value.High.ToString(); FindRow[Loop_coint] = Convert.ToString(Forobject_Row) + "," + Convert.ToString(Forobject_Column); Forobject_Row++; Forobject_Row++;
            //    DummyData[Forobject_Row, Forobject_Column] = item.Value.Low.ToString();
            //}
            //else
            //{

            //    string[] Find_Row_Split = FindRow[Loop_coint].Split(',');

            //    DummyData[Convert.ToInt16(Find_Row_Split[0]) - 1, Convert.ToInt16(Find_Row_Split[1])] = item.Key.ToString();
            //    DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])] = item.Value.Low.ToString();
            //    DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])] = item.Value.High.ToString();
            //}
            //Spec[0] = Convert.ToString(item.Value.Low);
            //Spec[1] = Convert.ToString(item.Value.High);


        }

        private void MovetoSepecC_L(KeyValuePair<string, CSV_Class.For_Box> item, string Spec_Num, object[,] SpecSheetData, bool Falg)
        {
            dummy = new Dictionary<string, string>();

            dummy = Data_Interface.Spec_Band[Spec_Num];

            Spec[2] = "";
            Spec[3] = "";


            if (!Falg)
            {

                object dummyTestData = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1];
                if (dummyTestData != null)
                {
                    DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString(); Forobject_Row++;
                    Spec[3] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString();
                }
                else
                {
                    DummyData[Forobject_Row, Forobject_Column] = "999"; Forobject_Row++;
                    Spec[3] = "999";

                }

                dummyTestData = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1];

                if (dummyTestData != null)
                {
                    DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString(); Forobject_Row++;
                    Spec[2] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString();
                }
                else
                {
                    DummyData[Forobject_Row, Forobject_Column] = "-999"; Forobject_Row++;
                    Spec[2] = "-999";
                }
            }
            else
            {
                string[] Find_Row_Split = FindRow[Loop_coint].Split(',');

                object dummyTestData = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1];
                if (dummyTestData != null)
                {
                    DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString();
                    Spec[3] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString();
                }
                else
                {
                    DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])] = "999";
                    Spec[3] = "999";

                }

                dummyTestData = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1];

                if (dummyTestData != null)
                {
                    DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString();
                    Spec[2] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString();
                }
                else
                {
                    DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])] = "-999";
                    Spec[2] = "-999";
                }
            }




        }

        private void MovetoSpecNone(KeyValuePair<string, CSV_Class.For_Box> item, string Spec_Num, object[,] SpecSheetData)
        {
            DummyData[Forobject_Row, Forobject_Column] = item.Key.ToString(); Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column] = item.Value.Broadcom_Spec_Max.ToString(); FindRow[Loop_coint] = Convert.ToString(Forobject_Row) + "," + Convert.ToString(Forobject_Column); Forobject_Row++;
            DummyData[Forobject_Row, Forobject_Column] = item.Value.Broadcom_Spec_Min.ToString(); ; Forobject_Row++;


            dummy = Data_Interface.Spec_Band[Spec_Num];

            Spec[0] = item.Value.Broadcom_Spec_Max.ToString();
            Spec[1] = item.Value.Broadcom_Spec_Min.ToString();
            Spec[2] = "";
            Spec[3] = "";

            object dummyTestData_Max = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1];
            object dummyTestData_Min = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1];
            object dummyTestData_Typical = SpecSheetData[0, Convert.ToInt16(dummy["TYPICAL"]) - 1];
            bool flag = true;

            if (dummyTestData_Max == null && dummyTestData_Min == null && dummyTestData_Typical != null)
            {
                DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["TYPICAL"]) - 1].ToString(); Forobject_Row++;
                Spec[3] = SpecSheetData[0, Convert.ToInt16(dummy["TYPICAL"]) - 1].ToString();

                Spec[2] = "-999"; Forobject_Row++;
                flag = false;
            }

            if (flag)
            {

                if (dummyTestData_Max != null)
                {
                    DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString(); Forobject_Row++;
                    Spec[3] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MAX_POSITION"]) - 1].ToString();
                }
                else
                {
                    DummyData[Forobject_Row, Forobject_Column] = "999"; Forobject_Row++;
                    Spec[3] = "999";

                }



                if (dummyTestData_Min != null)
                {
                    DummyData[Forobject_Row, Forobject_Column] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString(); Forobject_Row++;
                    Spec[2] = SpecSheetData[0, Convert.ToInt16(dummy["SPEC_MIN_POSITION"]) - 1].ToString();
                }
                else
                {
                    DummyData[Forobject_Row, Forobject_Column] = "-999"; Forobject_Row++;
                    Spec[2] = "-999";
                }

            }


            item.Value.Apple_Spec_Max = Spec[3];
            item.Value.Apple_Spec_Min = Spec[2];

        }

        private void Mul(KeyValuePair<string, CSV_Class.For_Box> item, string Spec_Num, object[,] SpecSheetData, string Value)
        {
            string Replace = Value.Replace("*", "");
            string[] Find_Row_Split = FindRow[Loop_coint].Split(',');

            string D1 = Convert.ToString(DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])]);
            double D1_1 = Convert.ToDouble(D1) * Convert.ToDouble(Replace);

            string D2 = Convert.ToString(DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])]);
            double D2_1 = Convert.ToDouble(D2) * Convert.ToDouble(Replace);

            Convert_Data = Convert.ToDouble(Replace);

            DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])] = D1_1;
            DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])] = D2_1;

            Spec[0] = Convert.ToString(Convert.ToDouble(Spec[0]) * Convert.ToDouble(Replace));
            Spec[1] = Convert.ToString(Convert.ToDouble(Spec[1]) * Convert.ToDouble(Replace));


        }

        private void Divide(KeyValuePair<string, CSV_Class.For_Box> item, string Spec_Num, object[,] SpecSheetData, string Value)
        {
            string Replace = Value.Replace("*", "");
            string[] Find_Row_Split = FindRow[Loop_coint].Split(',');

            string D1 = Convert.ToString(DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])]);
            double D1_1 = Convert.ToDouble(D1) * Convert.ToDouble(Replace);

            string D2 = Convert.ToString(DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])]);
            double D2_1 = Convert.ToDouble(D2) * Convert.ToDouble(Replace);

            Convert_Data = Convert.ToDouble(Replace);

            DummyData[Convert.ToInt16(Find_Row_Split[0]), Convert.ToInt16(Find_Row_Split[1])] = D1_1;
            DummyData[Convert.ToInt16(Find_Row_Split[0]) + 1, Convert.ToInt16(Find_Row_Split[1])] = D2_1;

            Spec[0] = Convert.ToString(Convert.ToDouble(Spec[0]) / Convert.ToDouble(Replace));
            Spec[1] = Convert.ToString(Convert.ToDouble(Spec[1]) / Convert.ToDouble(Replace));


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
                    //      this.textBox4.Text += str + "\r" + "\n";
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

        private void textBox5_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] File = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in File)
                {
                    //       this.textBox5.Text += str + "\r" + "\n";
                }
            }
        }

        private void textBox5_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }

        private void textBox6_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] File = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in File)
                {
                    //       this.textBox6.Text += str + "\r" + "\n";
                }
            }
        }

        private void textBox6_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }


        #endregion

        private void textBox4_DragEnter_1(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }

        private void textBox4_DragDrop_1(object sender, DragEventArgs e)
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

        public void Matching_Lot_data()
        {
            Lot = new string[0];

            for (int k = 0; k < 1; k++)
            {
                // string Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

                string Query = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY Name";

                Lot = DB_Interface.Get_Data_By_Query(Query);

            }

            Remove_table();

            Matching_Lots = new Dictionary<string, Dictionary<string, List<string>>>();

            Matching_Lots_Test = new Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>();

            for (int Lot_index = 0; Lot_index < Lot.Length; Lot_index++)
            {
                string Lot_string = Lot[Lot_index];
                string[] Sub_Lot = new string[0];

                // string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
                string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
                string[] Lot_data = DB_Interface.Get_Data_By_Query(Query);
                Lot_Information = new Dictionary<string, List<string>>();
                information = new Dictionary<string, Dictionary<string, List<string>>>();


                Query = "Select count(*) from " + Lot[Lot_index];
                string[] dummy = DB_Interface.Get_Data_By_Query(Query);

                if (Convert.ToInt16(dummy[0]) == 0)
                {
                    DB_Interface.DropTable(Data_Interface, Lot[Lot_index]);
                    break;
                }


                if (Lot_data.Length != 0 && Lot_data.Length == 1)
                {
                    Query = "Select DISTINCT SUBLOT from " + Lot[Lot_index] + " where LotID = '" + Lot_data[0] + "'";
                    string[] data = DB_Interface.Get_Data_By_Query(Query);

                    Sub_Lot = Sub_Lot.Concat(data).ToArray();

                    Sub_Lot = Sub_Lot.Distinct().ToArray();
                    Array.Sort(Sub_Lot);

                    _Lot_Information_Dummy = new List<string>();

                    for (int k = 0; k < Sub_Lot.Length; k++)
                    {
                        _Lot_Information_Dummy.Add(Sub_Lot[k]);
                    }

                    Lot_Information.Add(Lot_data[0], _Lot_Information_Dummy);
                    information.Add("LOTID", Lot_Information);


                    Matching_Lots.Add(Lot[Lot_index], Lot_Information);

                    string[] Wafer = new string[0];

                    Query = "Select DISTINCT WAFER_ID from " + Lot[Lot_index];
                    data = DB_Interface.Get_Data_By_Query(Query);

                    Wafer = Wafer.Concat(data).ToArray();

                    Wafer = Wafer.Distinct().ToArray();
                    Array.Sort(Wafer);

                    _Lot_Information_Dummy = new List<string>();
                    Lot_Information = new Dictionary<string, List<string>>();

                    for (int k = 0; k < Wafer.Length; k++)
                    {
                        _Lot_Information_Dummy.Add(Wafer[k]);
                    }

                    Lot_Information.Add(Lot_data[0], _Lot_Information_Dummy);
                    information.Add("WAFERID", Lot_Information);

                    Matching_Lots_Test.Add(Lot[Lot_index], information);
                }
                else if (Lot_data.Length > 1)
                {
                    for (int i = 0; i < Lot_data.Length; i++)
                    {
                        Query = "Select DISTINCT SUBLOT from " + Lot[Lot_index] + " where LotID = '" + Lot_data[i] + "'";
                        string[] data = DB_Interface.Get_Data_By_Query(Query);

                        Sub_Lot = Sub_Lot.Concat(data).ToArray();

                        Sub_Lot = Sub_Lot.Distinct().ToArray();
                        Array.Sort(Sub_Lot);

                        _Lot_Information_Dummy = new List<string>();

                        for (int k = 0; k < Sub_Lot.Length; k++)
                        {
                            _Lot_Information_Dummy.Add(Sub_Lot[k]);
                        }

                        Lot_Information.Add(Lot_data[i], _Lot_Information_Dummy);

                    }
                    information.Add("LOTID", Lot_Information);
                    Matching_Lots.Add(Lot[Lot_index], Lot_Information);


                    Lot_Information = new Dictionary<string, List<string>>();

                    for (int i = 0; i < Lot_data.Length; i++)
                    {
                        string[] Wafer = new string[0];

                        Query = "Select DISTINCT WAFER_ID from " + Lot[Lot_index];
                        string[] data = DB_Interface.Get_Data_By_Query(Query);

                        Wafer = Wafer.Concat(data).ToArray();

                        Wafer = Wafer.Distinct().ToArray();
                        Array.Sort(Wafer);

                        _Lot_Information_Dummy = new List<string>();


                        for (int k = 0; k < Wafer.Length; k++)
                        {
                            _Lot_Information_Dummy.Add(Wafer[k]);
                        }

                        Lot_Information.Add(Lot_data[i], _Lot_Information_Dummy);


                    }

                    information.Add("WAFERID", Lot_Information);


                    Matching_Lots_Test.Add(Lot[Lot_index], information);
                }
            }
        }

        public void Remove_table()
        {

            int k = 0;
            for (int i = 0; i < Lot.Length; i++)
            {
                if (Lot[i] == "Clotho_Spec")
                {

                    Lot[i] = ""; k++;
                }
                else if (Lot[i] == "Files")
                {
                    Lot[i] = ""; k++;
                }
                else if (Lot[i] == "INF")
                {
                    Lot[i] = ""; k++;

                }
                else if (Lot[i] == "REFHEADER")
                {
                    Lot[i] = ""; k++;
                }
                else if (Lot[i] == "Customer_Spec")
                {
                    Lot[i] = ""; k++;
                }

            }

            Lot = Lot.Where(x => !string.IsNullOrEmpty(x)).ToArray();
        }

        public void Define_Parameter()
        {
            Dialog2.Reset();
            Dialog2 = new OpenFileDialog();
            Box_Enum = new Dictionary<int, Dictionary<int, string>>();
            //  Dialog.Filter = ".csv";
            Dialog2.InitialDirectory = "C:\\Automation\\box_plot\\";
            Dialog2.Multiselect = false;
            Dialog2.ShowDialog();
            string[] Ignore_Spec = new string[2];

            bool flag = false;
            int Row = 0;
            if (Dialog2.FileNames.Length > 0)
            {
                CSV_Class.CSV CSV = new CSV_Class.CSV();

                CSV_Interface = CSV.Open(Key);

                CSV_Interface.Read_Open(Dialog2.FileNames[0]);
                while (!CSV_Interface.StreamReader.EndOfStream)
                {
                    string[] data = CSV_Interface.Read();

                    if (data[0].ToUpper() == "LABEL")
                    {
                      

                    }
                    else if (data[0] != "" && flag)
                    {
                     

                        string[] split = data[2].Split('>');


                        Dictionary<int, string> Test = new Dictionary<int, string>();
                        int Key = 0;
                        bool Flag_Test = true;

                        for (int kk = 0; kk < split.Length; kk++)
                        {
                            foreach (BoxPlot i in Enum.GetValues(typeof(BoxPlot)))
                            {
                                string info = i.ToString();
                                int NB = (int)i;

                                if (Flag_Test)
                                {
                                    Test.Add(999, data[1]);
                                    Flag_Test = false;
                                }
                                if (split[kk].Trim() == Convert.ToString(NB).Trim())
                                {
                                    Test.Add(NB, info.Trim());
                                    break;

                                }
                            }
                        }


                        if (!Box_Enum.ContainsKey(Convert.ToInt16(data[0])))
                        {
                            Test.Add(888, data[4]);
                            if(data[5].ToUpper().Contains("BY"))
                            {
                                string[] split1 = data[5].Split(':');

                                if(split1.Length > 1)
                                {
                                    string[] By = split1[1].Split('>');

                                    if (By.Length > 1)
                                    {
                                        Test.Add(777, split1[1]);
                                    }
                                    else
                                    {
                                        Test.Add(777, split1[1]);
                                    }
                      
                                }
                            }
                            Box_Enum.Add(Convert.ToInt16(data[0]), Test);
                            //  Box_Enum.Add(Ignore_Spec[1]);
                        }

                        //   dataGridView3.Rows[Row].Cells[i].Value = data[i].ToString();

                        Row++;
                    }
                    else if (data[0].ToUpper() == "START")
                    {
                        flag = true;
                      
                        Row++;
                    }
                }

            }
            CSV_Interface.Read_Close();
        }

        public enum BoxPlot
        {
            Identifier = 0,
            Parameter,
            Measuer,
            Band,
            Pmode,
            Modulation,
            Waveform,
            Power_Identifier,
            Pout,
            Frequency,
            Vcc,
            Vdd,
            DAC1,
            DAC2,
            TX,
            ANT,
            RX,
            Extra,
            Note1,
            SpecNumber,
            Site,
            Lot,
            Wafer

        }
    }
}
