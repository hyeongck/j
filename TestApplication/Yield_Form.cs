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
using System.Threading;


namespace TestApplication
{
    public partial class Yield_Form : Form
    {
        Data_Class.Data_Editing.INT Data_Interface;
        Data_Class.Data_Editing Data_Edit = new Data_Class.Data_Editing();

        DB_Class.DB_Editing DB = new DB_Class.DB_Editing();
        DB_Class.DB_Editing.INT DB_Interface;

        CSV_Class.CSV.INT Csv_Interface;
        CSV_Class.CSV CSV = new CSV_Class.CSV();

        string[] Lot;


        Dictionary<string, List<string>> Lot_Information;
        Dictionary<string, string> Matching_Lots;

        public Yield_Form()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (textBox1.Text != "")
            //{
            //    int workerT, PortT;
            //    int MinRequire  = 25;

            //    ThreadPool.GetMinThreads(out workerT, out PortT);

            //    if (workerT < MinRequire)
            //    {
            //        int inc = MinRequire - workerT;
            //        workerT += inc;
            //    }

            //    ThreadPool.SetMinThreads(workerT, 20);

            //    Data_Class.Data_Editing.Yield Data = new Data_Class.Data_Editing.Yield();

            //    string Key = "YIELD";

            //    Csv_Interface = CSV.Open(Key);
            //    Data_Interface = Data_Edit.Open(Key);
            //    DB_Interface = DB.Open(Key);

            //    string Files1 = textBox1.Text;
            //    Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

            //    Yield_Cal_Form.CSV_File_Path = Files1;

            //    Csv_Interface.Read_Open(Files1);

            //    #region Find_First_and_Spec_Row

            //    // Find First Row

            //    while (!Csv_Interface.StreamReader.EndOfStream)
            //    {
            //        Csv_Interface.Read();
            //        bool Flag = Data_Interface.Find_First_Row(Csv_Interface.Get_String);
            //        if (Flag) break;
            //    }

            //    for (int l = 0; l < Csv_Interface.Get_String.Length; l++)
            //    {
            //        if (Csv_Interface.Get_String[l].ToUpper() == "SBIN")
            //        {
            //            DB_Interface.Bin_place = 1;
            //        }

            //    }
            //    // Find Spec High

            //while (!Csv_Interface.StreamReader.EndOfStream)
            //{
            //    Csv_Interface.Read();
            //    bool Flag = Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);
            //    if (Flag) break;
            //}

            //    // Find Spec Low

            //    while (!Csv_Interface.StreamReader.EndOfStream)
            //    {
            //        Csv_Interface.Read();
            //        bool Flag = Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);
            //        if (Flag) break;
            //    }

            //    #endregion


            //    Data_Interface.Define_DB_Count(Csv_Interface.Get_String);
            //    Data_Interface.Make_New_header();

            //    DB_Interface.Open_DB(Files1, Data_Interface);
            //    DB_Interface.DropTable(Data_Interface,"");

            //    DB_Interface.trans(Data_Interface);
            //    DB_Interface.Insert_Header(Data_Interface);
            //    DB_Interface.Insert_Spec_Header(Data_Interface);
            //    DB_Interface.Insert_New_Spec_Header(Data_Interface);
            //    DB_Interface.Insert_Spec_Data("spec");
            //    DB_Interface.Insert_Spec_Data("newspec");

            //    Stopwatch TestTime = new Stopwatch();
            //    TestTime.Restart();
            //    TestTime.Start();

            //    int Data_Count = 0;

            //    Insert_Count_Form IForm = new Insert_Count_Form();
            //    IForm.Show();

            //    string split_File = Files1.Substring(Files1.LastIndexOf("\\") + 1);
            //    string[] _split_File = split_File.Split('_');

            //    if (Csv_Interface.Get_String[8] == "" || Csv_Interface.Get_String[9] == "")
            //    {
            //        DB_Interface.Lot_ID = _split_File[1];
            //        DB_Interface.SubLot_ID = _split_File[2];
            //    }
            //    else
            //    {
            //        DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
            //        DB_Interface.SubLot_ID = Csv_Interface.Get_String[9];
            //    }



            //    DB_Interface.Insert_ThreadFlags = new ManualResetEvent[2];
            //    DB_Interface.Insert_Thread_Wait = new bool[2];

            //    for (int thread_i = 0; thread_i < 2; thread_i++)
            //    {
            //        DB_Interface.Insert_ThreadFlags[thread_i] = new ManualResetEvent(false);
            //    }

            //    DB_Interface.TheFirst_Trashes_Header_Count = Data_Interface.TheFirst_Trashes_Header_Count;
            //    DB_Interface.TheEnd_Trashes_Header_Count = Data_Interface.TheEnd_Trashes_Header_Count;

            //    string[] GetData = Csv_Interface.Read();
            //    Data_Interface.Getstring = GetData;
            //    DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];
            //    Data_Count++;

            //    for (int thread_i = 0; thread_i < 2; thread_i++)
            //    {
            //        DB_Interface.Insert_ThreadFlags[thread_i].Reset();
            //    }
            //    ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

            //    DB_Interface.Insert_ThreadFlags[1].Set();

            //    DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
            //    DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

            //    GetData = Csv_Interface.Read();
            //    Data_Interface.Getstring = GetData;

            //    if (Csv_Interface.Get_String[8] == "" || Csv_Interface.Get_String[9] == "")
            //    {
            //        DB_Interface.Lot_ID = _split_File[1];
            //        DB_Interface.SubLot_ID = _split_File[2];
            //    }
            //    else
            //    {
            //        DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
            //        DB_Interface.SubLot_ID = Csv_Interface.Get_String[9];
            //    }

            //    DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

            //    Data_Count++;


            //    while (!Csv_Interface.StreamReader.EndOfStream)
            //    {
            //        Stopwatch TestTime1 = new Stopwatch();
            //        TestTime1.Restart();
            //        TestTime1.Start();

            //        for (int thread_i = 0; thread_i < 2; thread_i++)
            //        {
            //            DB_Interface.Insert_ThreadFlags[thread_i].Reset();
            //        }
            //        ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

            //        GetData = Csv_Interface.Read_Test();


            //        DB_Interface.Insert_ThreadFlags[1].Set();

            //        DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
            //        DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

            //        Data_Interface.Getstring = GetData;

            //        if (Csv_Interface.Get_String[8] == "" || Csv_Interface.Get_String[9] == "")
            //        {
            //            DB_Interface.Lot_ID = _split_File[1];
            //            DB_Interface.SubLot_ID = _split_File[2];
            //        }
            //        else
            //        {
            //            DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
            //            DB_Interface.SubLot_ID = Csv_Interface.Get_String[9];
            //        }
            //        DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

            //        IForm.Print_Count(Data_Count);

            //        Data_Count++;
            //    }

            //    for (int thread_i = 0; thread_i < 2; thread_i++)
            //    {
            //        DB_Interface.Insert_ThreadFlags[thread_i].Reset();
            //    }
            //    ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

            //    DB_Interface.Insert_ThreadFlags[1].Set();

            //    DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
            //    DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

            //    DB_Interface.Commit(Data_Interface);
            //    double Testime = TestTime.Elapsed.TotalMilliseconds;

            //   // MessageBox.Show(Convert.ToString(Testime));

            //    IForm.Close();

            //    List<int>[] TestResult = new List<int>[Data_Interface.New_Header.Length];
            //    Dictionary<string, List<int>> TestResult_Dic = new Dictionary<string, List<int>>();
            //    DB_Interface.Cal_Value_by_rowsdata = new Dictionary<string, DB_Class.DB_Editing.Data_Calculation>();

            //    int i = 0;
            //    foreach (var item in TestResult)
            //    {
            //        TestResult[i] = new List<int>();
            //        TestResult[i].Add(0);
            //        i++;
            //    }

            //    foreach (List<int>[] item in DB_Interface.ForCampare_Yield_List1)
            //    {
            //        for (int k = 0; k < item.Length; k++)
            //        {
            //            for (int j = 0; j < item[k].Count; j++)
            //            {
            //                if (item[k][j] == 1)
            //                {
            //                    TestResult[(k * Data_Interface.DB_Column_Limit) + j][0] += 1;
            //                }
            //            }
            //        }
            //    }
            //    double[] dummy = new double[11];

            //    for (int j = 0; j < Data_Interface.New_Header.Length; j++)
            //    {
            //        TestResult_Dic.Add(Data_Interface.Reference_Header[j], TestResult[j]);
            //        DB_Interface.Cal_Value_by_rowsdata.Add(Data_Interface.Reference_Header[j], new DB_Class.DB_Editing.Data_Calculation(dummy));
            //    }

            //    //Csv_Interface.Connect_Write("C:\\Automation\\Yield\\1.csv");

            //    //foreach (var item in TestResult_Dic)
            //    //{
            //    //    CSV.Write(Csv_Interface, Key, item.Key.ToString(), item.Value[0].ToString());
            //    //}

            //    //CSV.Write_Close(Csv_Interface, Key);
            //    Csv_Interface.Read_Close();

            //    Yield_Cal_Form Yield_Cal = new Yield_Cal_Form(DB_Interface.ForCampare_Yield_List1, DB_Interface.ForCampare_Yield_Fro_DB_List, TestResult_Dic, Key, Data_Edit, Data_Interface, DB, DB_Interface, Data_Count ,0);
            //    Yield_Cal.Show();

            //}

        }

        private void button2_Click(object sender, EventArgs e)
        {

            OpenFileDialog Dialog = new OpenFileDialog();

            Dialog.Filter = "DB Files (*.db)| *.db";
            Dialog.InitialDirectory = "C:\\Automation\\DB\\yield";
            Dialog.Multiselect = true;
            Dialog.ShowDialog();


            if (Dialog.FileNames.Length > 0)
            {

                Data_Class.Data_Editing.Yield Data = new Data_Class.Data_Editing.Yield();

                string Key = "YIELD";
                Data_Interface = Data_Edit.Open(Key);
                Csv_Interface = CSV.Open(Key);
                DB_Interface = DB.Open(Key);


                string Files1 = textBox2.Text;
                Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

                Yield_Cal_Form.CSV_File_Path = Files1;

                DB_Interface.Open_DB(Dialog.FileNames, Data_Interface);

                DB_Interface.Filename = Dialog.FileNames[0];
                #region NPI Spec

                if (textBox2.Text == "")
                {
                    DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                    //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                    Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                    Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                    Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                    Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                    Data_Interface.Make_New_header();

                    Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                    //  DB_Interface.Insert_Spec_Header(Data_Interface);

                    Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                    string[] BIn_Count = new string[5];
                    //for (int k = 0; k < 1; k++)
                    //{
                    //    string Query1 = "select count(id) from Clotho_Spec";

                    //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                    //}

                    //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                    for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                    {
                        Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                        Data_Interface.Clotho_Spcc_List.Add(Specs);
                    }

                    // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                    Data_Interface.Data_Table = "Clotho_Spec";
                    DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);

                    Data_Interface.SWBIN_Dic = new Dictionary<string, Data_Class.Data_Editing.SWBIN>();

                    for (int k = 1; k < 6; k++)
                    {
                        Data_Class.Data_Editing.SWBIN SS = new Data_Class.Data_Editing.SWBIN(Convert.ToString(k), Convert.ToString(k), false);
                        Data_Interface.SWBIN_Dic.Add(Convert.ToString(k), SS);
                    }

                }
                else
                {
                    if (MessageBox.Show("Do you want to Replace Clotho Spec?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Data_Interface.Data_Table = "Clotho_Spec";
                        DB_Interface.DropTable(Data_Interface, "drop table Clotho_Spec");

                        Csv_Interface.Read_Open(Files1);

                        int m = 0;
                        Data_Interface.Clotho_Spec_Data = new string[40000];
                        while (!Csv_Interface.StreamReader.EndOfStream)
                        {
                            Data_Interface.Clotho_Spec_Data[m] = Csv_Interface.Read_Cloth_Spec();
                            m++;
                        }
                        var Var = Data_Interface.Clotho_Spec_Data;
                        Array.Resize(ref Var, m);
                        Data_Interface.Clotho_Spec_Data = Var;
                        Var = null;

                        Data_Interface.Find_Cloth_DataFile(Data_Interface.Clotho_Spec_Data);

                        Csv_Interface.Read_Close();


                        DB_Interface.DIC_IQR = new Dictionary<string, DB_Class.DB_Editing.IQR>();


                        for (int i = 0; i < Data_Interface.Ref_New_Header.Length; i++)
                        {
                            DB_Class.DB_Editing.IQR Dummy = new DB_Class.DB_Editing.IQR(1.5, 1.5, null);

                            DB_Interface.DIC_IQR.Add(Data_Interface.Ref_New_Header[i], Dummy);
                        }

                        Data_Interface.Reference_Header = Data_Interface.Ref_New_Header;
                        Csv_Interface.Read_Close();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);
                        Data_Interface.Make_New_header();

                        Data_Interface.Data_Table = "Clotho_Spec";
                        DB_Interface.Insert_Spec_Header(Data_Interface);


                        DB_Interface.Insert_Spec_Data("Clotho_Spec");


                    }
                    else
                    {

                

                        DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                        //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                        Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                        Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                        Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                        Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                        Data_Interface.Make_New_header();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                      //  DB_Interface.Insert_Spec_Header(Data_Interface);

                        Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                        string[] BIn_Count = new string[5];
                        //for (int k = 0; k < 1; k++)
                        //{
                        //    string Query1 = "select count(id) from Clotho_Spec";

                        //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                        //}

                        //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                        for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                        {
                            Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                            Data_Interface.Clotho_Spcc_List.Add(Specs);
                        }

                       // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                        Data_Interface.Data_Table = "Clotho_Spec";
                        DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);
                    }

                }

                #endregion

                #region Customer Spec

                if (textBox3.Text == "")
                {
                    DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                    //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                    Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                    Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                    Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                    Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                    Data_Interface.Make_New_header();

                    Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                    //  DB_Interface.Insert_Spec_Header(Data_Interface);

                    Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                    string[] BIn_Count = new string[5];
                    //for (int k = 0; k < 1; k++)
                    //{
                    //    string Query1 = "select count(id) from Clotho_Spec";

                    //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                    //}

                    //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                    for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                    {
                        Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                        Data_Interface.Customor_Clotho_List.Add(Specs);
                    }

                    // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                    Data_Interface.Data_Table = "Customer_Spec";
                    DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);

                    Data_Interface.SWBIN_Dic = new Dictionary<string, Data_Class.Data_Editing.SWBIN>();

                    for(int k = 1; k < 6; k ++)
                    {
                        Data_Class.Data_Editing.SWBIN SS = new Data_Class.Data_Editing.SWBIN(Convert.ToString(k), Convert.ToString(k), false);
                        Data_Interface.SWBIN_Dic.Add(Convert.ToString(k) , SS);
                    }

                }
                else
                {
                    if (MessageBox.Show("Do you want to Replace Customer Spec?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        //for (int k = 0; k < 10; k++)
                        //{
                        //    Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

                        //    DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(Data_Interface, Query);
                        //}

                        string Files2 = textBox3.Text;
                        Files2 = Files2.Replace('\r', ' ').Replace('\n', ' ').Trim();


                        Data_Interface.Data_Table = "Customer_Spec";
                        DB_Interface.DropTable(Data_Interface, "drop table Customer_Spec");

                        Csv_Interface.Read_Open(Files2);

                        int m = 0;
                        Data_Interface.Customer_Clotho_Spec_Data = new string[40000];
                        while (!Csv_Interface.StreamReader.EndOfStream)
                        {
                            Data_Interface.Customer_Clotho_Spec_Data[m] = Csv_Interface.Read_Cloth_Spec();
                            m++;
                        }
                        var Var = Data_Interface.Customer_Clotho_Spec_Data;
                        Array.Resize(ref Var, m);
                        Data_Interface.Customer_Clotho_Spec_Data = Var;
                        Var = null;

                        Data_Interface.Find_Cloth_DataFile_For_New_Spec(Data_Interface.Customer_Clotho_Spec_Data);

                        Csv_Interface.Read_Close();


                        DB_Interface.DIC_IQR = new Dictionary<string, DB_Class.DB_Editing.IQR>();


                        for (int i = 0; i < Data_Interface.Ref_New_Header.Length; i++)
                        {
                            DB_Class.DB_Editing.IQR Dummy = new DB_Class.DB_Editing.IQR(1.5, 1.5, null);

                            DB_Interface.DIC_IQR.Add(Data_Interface.Ref_New_Header[i], Dummy);
                        }

                        Data_Interface.Reference_Header = Data_Interface.Ref_New_Header;
                        Csv_Interface.Read_Close();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);
                        Data_Interface.Make_New_header();

                        Data_Interface.Data_Table = "Customer_Spec";
                        DB_Interface.Insert_Spec_Header(Data_Interface);


                        DB_Interface.Insert_Spec_Data("Customer_Spec");


                    }
                    else
                    {



                        DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                        //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                        Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                        Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                        Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                        Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                        Data_Interface.Make_New_header();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                        //  DB_Interface.Insert_Spec_Header(Data_Interface);

                        Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                        string[] BIn_Count = new string[5];
                        //for (int k = 0; k < 1; k++)
                        //{
                        //    string Query1 = "select count(id) from Clotho_Spec";

                        //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                        //}

                        //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                        for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                        {
                            Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                            Data_Interface.Customor_Clotho_List.Add(Specs);
                        }

                        // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                        Data_Interface.Data_Table = "Customer_Spec";
                        DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);
                    }

                }

                #endregion

                Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;
               


                DB_Interface.Table_Count = 0;
           
                List<int>[] TestResult = new List<int>[Data_Interface.New_Header.Length];
                Dictionary<string, List<int>> TestResult_Dic = new Dictionary<string, List<int>>();
                DB_Interface.Cal_Value_by_rowsdata = new Dictionary<string, DB_Class.DB_Editing.Data_Calculation>();

                TestResult = new List<int>[1];
                TestResult[0] = new List<int>();
                TestResult[0].Add(0);
                double[] dummy = new double[14];

                for (int j = 0; j < Data_Interface.New_Header.Length; j++)
                {
                    TestResult_Dic.Add(Data_Interface.Ref_New_Header[j], TestResult[0]);
                    DB_Interface.Cal_Value_by_rowsdata.Add(Data_Interface.Ref_New_Header[j], new DB_Class.DB_Editing.Data_Calculation(Data_Interface.Clotho_Spcc_List[0].Max.Length));
                }

                string Query = "";

                //for (int k = 0; k < 10; k++)
                //{
                //    Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

                //    DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(Data_Interface, Query);
                //}



                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                Matching_Lot_data();

                int Sample_Count = 0;

                foreach(KeyValuePair<string,string>t in Matching_Lots)
                {
                    Query = "select count(SITEID) from " + t.Value;
                    Sample_Count += DB_Interface.Get_Sample_Count(0, Query);
                }

                double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                
                int Hidden_Sample_Count = 0;


                foreach (KeyValuePair<string, string> t in Matching_Lots)
                {
                    Query = "select count(parameter) from " + t.Value + "  where FAIL = 1";

                    Hidden_Sample_Count += DB_Interface.Get_Sample_Count(0, Query);
                }

                double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;



                //DB_Interface.Read_Dispose(Data_Interface);
                //double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

                //DB_Interface.Set_Conn(Data_Interface);
                //double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;

                DB_Interface._From_Db = true;
                Yield_Cal_Form Yield_Cal = new Yield_Cal_Form(DB_Interface.ForCampare_Yield_List1, DB_Interface.ForCampare_Yield_Fro_DB_List, TestResult_Dic, Key, Data_Edit, Data_Interface, DB, DB_Interface, Sample_Count, Hidden_Sample_Count);




            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            OpenFileDialog Dialog = new OpenFileDialog();

            Dialog.Filter = "DB Files (*.db)| *.db";
            Dialog.InitialDirectory = "C:\\Automation\\DB\\yield";
            Dialog.Multiselect = true;
            Dialog.ShowDialog();


            if (Dialog.FileNames.Length > 0)
            {

                Data_Class.Data_Editing.Yield Data = new Data_Class.Data_Editing.Yield();

                string Key = "YIELD";
                Data_Interface = Data_Edit.Open(Key);
                Csv_Interface = CSV.Open(Key);
                DB_Interface = DB.Open(Key);


                string Files1 = textBox4.Text;
                Files1 = Files1.Replace('\r', ' ').Replace('\n', ' ').Trim();

                Yield_Cal_Form.CSV_File_Path = Files1;

                DB_Interface.Open_DB(Dialog.FileNames, Data_Interface);

                DB_Interface.Filename = Dialog.FileNames[0];
                #region NPI Spec

                if (textBox4.Text == "")
                {
                    DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                    //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                    Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                    Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                    Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                    Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                    Data_Interface.Make_New_header();

                    Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                    //  DB_Interface.Insert_Spec_Header(Data_Interface);

                    Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                    string[] BIn_Count = new string[5];
                    //for (int k = 0; k < 1; k++)
                    //{
                    //    string Query1 = "select count(id) from Clotho_Spec";

                    //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                    //}

                    //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                    for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                    {
                        Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                        Data_Interface.Clotho_Spcc_List.Add(Specs);
                    }

                    // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                    Data_Interface.Data_Table = "Clotho_Spec";
                    DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);

                    Data_Interface.SWBIN_Dic = new Dictionary<string, Data_Class.Data_Editing.SWBIN>();

                    for (int k = 1; k < 6; k++)
                    {
                        Data_Class.Data_Editing.SWBIN SS = new Data_Class.Data_Editing.SWBIN(Convert.ToString(k), Convert.ToString(k), false);
                        Data_Interface.SWBIN_Dic.Add(Convert.ToString(k), SS);
                    }

                }
                else
                {
                    if (MessageBox.Show("Do you want to Replace Clotho Spec?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Data_Interface.Data_Table = "Clotho_Spec";
                        DB_Interface.DropTable(Data_Interface, "drop table Clotho_Spec");

                        Csv_Interface.Read_Open(Files1);

                        int m = 0;
                        Data_Interface.Clotho_Spec_Data = new string[40000];
                        while (!Csv_Interface.StreamReader.EndOfStream)
                        {
                            Data_Interface.Clotho_Spec_Data[m] = Csv_Interface.Read_Cloth_Spec();
                            m++;
                        }
                        var Var = Data_Interface.Clotho_Spec_Data;
                        Array.Resize(ref Var, m);
                        Data_Interface.Clotho_Spec_Data = Var;
                        Var = null;

                        Data_Interface.Find_Cloth_DataFile(Data_Interface.Clotho_Spec_Data);

                        Csv_Interface.Read_Close();


                        DB_Interface.DIC_IQR = new Dictionary<string, DB_Class.DB_Editing.IQR>();


                        for (int i = 0; i < Data_Interface.Ref_New_Header.Length; i++)
                        {
                            DB_Class.DB_Editing.IQR Dummy = new DB_Class.DB_Editing.IQR(1.5, 1.5, null);

                            DB_Interface.DIC_IQR.Add(Data_Interface.Ref_New_Header[i], Dummy);
                        }

                        Data_Interface.Reference_Header = Data_Interface.Ref_New_Header;
                        Csv_Interface.Read_Close();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);
                        Data_Interface.Make_New_header();

                        Data_Interface.Data_Table = "Clotho_Spec";
                        DB_Interface.Insert_Spec_Header(Data_Interface);


                        DB_Interface.Insert_Spec_Data("Clotho_Spec");


                    }
                    else
                    {



                        DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                        //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                        Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                        Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                        Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                        Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                        Data_Interface.Make_New_header();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                        //  DB_Interface.Insert_Spec_Header(Data_Interface);

                        Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                        string[] BIn_Count = new string[5];
                        //for (int k = 0; k < 1; k++)
                        //{
                        //    string Query1 = "select count(id) from Clotho_Spec";

                        //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                        //}

                        //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                        for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                        {
                            Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                            Data_Interface.Clotho_Spcc_List.Add(Specs);
                        }

                        // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                        Data_Interface.Data_Table = "Clotho_Spec";
                        DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);
                    }

                }

                #endregion

                if (textBox1.Text == "")
                {
                    DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                    //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                    Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                    Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                    Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                    Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                    Data_Interface.Make_New_header();

                    Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                    //  DB_Interface.Insert_Spec_Header(Data_Interface);

                    Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                    string[] BIn_Count = new string[5];
                    //for (int k = 0; k < 1; k++)
                    //{
                    //    string Query1 = "select count(id) from Clotho_Spec";

                    //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                    //}

                    //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                    for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                    {
                        Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                        Data_Interface.Customor_Clotho_List.Add(Specs);
                    }

                    // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                    Data_Interface.Data_Table = "Customer_Spec";
                    DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);

                    Data_Interface.SWBIN_Dic = new Dictionary<string, Data_Class.Data_Editing.SWBIN>();

                    for (int k = 1; k < 6; k++)
                    {
                        Data_Class.Data_Editing.SWBIN SS = new Data_Class.Data_Editing.SWBIN(Convert.ToString(k), Convert.ToString(k), false);
                        Data_Interface.SWBIN_Dic.Add(Convert.ToString(k), SS);
                    }

                }
                else
                {
                    if (MessageBox.Show("Do you want to Replace Customer Spec?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        //for (int k = 0; k < 10; k++)
                        //{
                        //    Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

                        //    DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(Data_Interface, Query);
                        //}

                        string Files2 = textBox1.Text;
                        Files2 = Files2.Replace('\r', ' ').Replace('\n', ' ').Trim();


                        Data_Interface.Data_Table = "Customer_Spec";
                        DB_Interface.DropTable(Data_Interface, "drop table Customer_Spec");

                        Csv_Interface.Read_Open(Files2);

                        int m = 0;
                        Data_Interface.Customer_Clotho_Spec_Data = new string[40000];
                        while (!Csv_Interface.StreamReader.EndOfStream)
                        {
                            Data_Interface.Customer_Clotho_Spec_Data[m] = Csv_Interface.Read_Cloth_Spec();
                            m++;
                        }
                        var Var = Data_Interface.Customer_Clotho_Spec_Data;
                        Array.Resize(ref Var, m);
                        Data_Interface.Customer_Clotho_Spec_Data = Var;
                        Var = null;

                        Data_Interface.Find_Cloth_DataFile_For_New_Spec(Data_Interface.Customer_Clotho_Spec_Data);

                        Csv_Interface.Read_Close();


                        DB_Interface.DIC_IQR = new Dictionary<string, DB_Class.DB_Editing.IQR>();


                        for (int i = 0; i < Data_Interface.Ref_New_Header.Length; i++)
                        {
                            DB_Class.DB_Editing.IQR Dummy = new DB_Class.DB_Editing.IQR(1.5, 1.5, null);

                            DB_Interface.DIC_IQR.Add(Data_Interface.Ref_New_Header[i], Dummy);
                        }

                        Data_Interface.Reference_Header = Data_Interface.Ref_New_Header;
                        Csv_Interface.Read_Close();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);
                        Data_Interface.Make_New_header();

                        Data_Interface.Data_Table = "Customer_Spec";
                        DB_Interface.Insert_Spec_Header(Data_Interface);


                        DB_Interface.Insert_Spec_Data("Customer_Spec");


                    }
                    else
                    {



                        DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                        //   Data_Interface.Clotho_Spcc_List = new List<Data_Class.Data_Editing.Clotho_Spec>[Bin];
                        Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();
                        Data_Interface.Reference_Header = Data_Interface.Reference_Header_List.ToArray();

                        Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                        Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;

                        Data_Interface.Make_New_header();

                        Data_Interface.Define_DB_Count(Data_Interface.Ref_New_Header);

                        //  DB_Interface.Insert_Spec_Header(Data_Interface);

                        Data_Interface.Customor_Clotho_List = new List<Data_Class.Data_Editing.Clotho_Spec>();

                        string[] BIn_Count = new string[5];
                        //for (int k = 0; k < 1; k++)
                        //{
                        //    string Query1 = "select count(id) from Clotho_Spec";

                        //    BIn_Count = DB_Interface.Get_Data_By_Query(Query1);

                        //}

                        //int Bin = Convert.ToInt16(BIn_Count[0]) / 2;

                        for (int k = 0; k < Data_Interface.Reference_Header_List.Count; k++)
                        {
                            Data_Class.Data_Editing.Clotho_Spec Specs = new Data_Class.Data_Editing.Clotho_Spec(new double[5], new double[5]);
                            Data_Interface.Customor_Clotho_List.Add(Specs);
                        }

                        // DB_Interface.Insert_Spec_Data("Clotho_Spec");

                        Data_Interface.Data_Table = "Customer_Spec";
                        DB_Interface.Road_Save_Customer_Spec_table(Data_Interface);
                    }

                }


                Data_Interface.Ref_New_Header = new string[Data_Interface.Reference_Header.Length];

                Data_Interface.Ref_New_Header = Data_Interface.Reference_Header;



                DB_Interface.Table_Count = 0;

                List<int>[] TestResult = new List<int>[Data_Interface.New_Header.Length];
                Dictionary<string, List<int>> TestResult_Dic = new Dictionary<string, List<int>>();
                DB_Interface.Cal_Value_by_rowsdata = new Dictionary<string, DB_Class.DB_Editing.Data_Calculation>();

                TestResult = new List<int>[1];
                TestResult[0] = new List<int>();
                TestResult[0].Add(0);
                double[] dummy = new double[14];

                for (int j = 0; j < Data_Interface.New_Header.Length; j++)
                {
                    TestResult_Dic.Add(Data_Interface.Ref_New_Header[j], TestResult[0]);
                    DB_Interface.Cal_Value_by_rowsdata.Add(Data_Interface.Ref_New_Header[j], new DB_Class.DB_Editing.Data_Calculation(Data_Interface.Clotho_Spcc_List[0].Max.Length));
                }

                string Query = "";

                //for (int k = 0; k < 10; k++)
                //{
                //    Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

                //    DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(Data_Interface, Query);
                //}



                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                Matching_Lot_data();

                int Sample_Count = 0;

                foreach (KeyValuePair<string, string> t in Matching_Lots)
                {
                    Query = "select count(SITEID) from " + t.Value;
                    Sample_Count += DB_Interface.Get_Sample_Count(0, Query);
                }

                double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;


                int Hidden_Sample_Count = 0;


                foreach (KeyValuePair<string, string> t in Matching_Lots)
                {
                    Query = "select count(parameter) from " + t.Value + "  where FAIL = 1";

                    Hidden_Sample_Count += DB_Interface.Get_Sample_Count(0, Query);
                }

                double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;



                //DB_Interface.Read_Dispose(Data_Interface);
                //double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

                //DB_Interface.Set_Conn(Data_Interface);
                //double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;

                DB_Interface._From_Db = true;
              //  Yield_Cal_Form_Second Yield_Cal = new Yield_Cal_Form_Second(DB_Interface.ForCampare_Yield_List1, DB_Interface.ForCampare_Yield_Fro_DB_List, TestResult_Dic, Key, Data_Edit, Data_Interface, DB, DB_Interface, Sample_Count, Hidden_Sample_Count);





            }
        }
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] File = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in File)
                {
                    //     this.textBox1.Text += str + "\r" + "\n";
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

        private void textBox3_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] File = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string str in File)
                {
                    this.textBox3.Text += str + "\r" + "\n";
                }
            }
        }

        private void textBox3_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
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
                else if (Lot[i] == "Customer_Spec")
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
                else if (Lot[i] == "spec")
                {
                    Lot[i] = ""; k++;
                }
                else if (Lot[i] == "NPI_Spec")
                {
                    Lot[i] = ""; k++;
                }
                else if (Lot[i] == "Current_Setting")
                {
                    Lot[i] = ""; k++;
                }
                else if (Lot[i].Contains("CHAN"))
                {
                    Lot[i] = ""; k++;
                }
                else if (Lot[i].Contains("Trace_Info"))
                {
                    Lot[i] = ""; k++;
                }

            }

            Lot = Lot.Where(x => !string.IsNullOrEmpty(x)).ToArray();
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


            Lot_Information = new Dictionary<string, List<string>>();
            Matching_Lots = new Dictionary<string, string>();

            for (int Lot_index = 0; Lot_index < Lot.Length; Lot_index++)
            {
                string Lot_string = Lot[Lot_index];
                string[] Sub_Lot = new string[0];

                string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
                string[] Lot_data = DB_Interface.Get_Data_By_Query(Query);

                if (Lot_data.Length != 0)
                {

                    Query = "Select DISTINCT SUBLOT from " + Lot[Lot_index] + " where LotID = '" + Lot_data[0] + "'";
                    string[] data = DB_Interface.Get_Data_By_Query(Query);

                    Sub_Lot = Sub_Lot.Concat(data).ToArray();


                    Sub_Lot = Sub_Lot.Distinct().ToArray();
                    Array.Sort(Sub_Lot);

                    List<string> _Lot_Information_Dummy = new List<string>();

                    for (int k = 0; k < Sub_Lot.Length; k++)
                    {
                        _Lot_Information_Dummy.Add(Sub_Lot[k]);
                    }

                    Lot_Information.Add(Lot_data[0], _Lot_Information_Dummy);
                    Matching_Lots.Add(Lot_data[0], Lot[Lot_index]);
                }
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

        private void textBox1_DragDrop_1(object sender, DragEventArgs e)
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

        private void textBox1_DragEnter_1(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
         //   ATE.SPARA_Form s = new ATE.SPARA_Form();
        }
    }
}
