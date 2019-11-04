using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using System.Threading;
using System.Diagnostics;
using System.Data.Common;
using System.Collections;
using System.Reflection;
using CSV_Class;

namespace DB_Class
{
    public class DB_Editing
    {


        public class FCM_Automation_EXCEL : INT
        {

            public Data_Class.Data_Editing.INT Data { get; set; }
            public ReaderWriterLockSlim[] sqlitelock { get; set; }
            public string[] strConn { get; set; }
            public SQLiteConnection[] conn { get; set; }
            public SQLiteCommand[] cmd { get; set; }

            public SQLiteDataAdapter[] sqlAdapter { get; set; }
            public SQLiteCommandBuilder[] sqlcmdbuilder { get; set; }
            public SQLiteDataReader[] SqReader { get; set; }

            public DbDataReader[] DbReader { get; set; }
            public DataSet[] ds { get; set; }
            public DataTable dt_test { get; set; }
            public DataTable[] dt { get; set; }
            public SQLiteTransaction[] tran { get; set; }

            public ManualResetEvent[] ThreadFlags { get; set; }

            public ManualResetEvent[] Insert_ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public int Limit { get; set; }
            public int Limit_Count { get; set; }
            public int Table_Count { get; set; }
            public bool[] Insert_Thread_Wait { get; set; }
            public double[] Testtime { get; set; }
            public string Filename { get; set; }
            public double[][] test { get; set; }
            public string[][] Teststring { get; set; }
            public double[][] Testdouble { get; set; }

            public object[] ID { get; set; }
            public object[] Value { get; set; }
            public object[] WAFER_ID { get; set; }
            public object[] LOT_ID { get; set; }
            public object[] SITE_ID { get; set; }
            public object[] Variation { get; set; }


            public Dictionary<string, double[]> Selected_Parameter_Distribution { get; set; }

            public List<List<RowAndPass>[]>[] Yield_Test { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test_New_Spec { get; set; }

            public List<List<int>[]>[] For_Any_Yield_Percent { get; set; }
            public List<List<int>>[] For_Any_Yield { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_Lot { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_SITE { get; set; }
            public List<int[]>[] ForCampare_Yield_Fro_DB { get; set; }
            public List<List<int>>[] For_New_Spec_ForCampare_Yield2 { get; set; }

            public List<List<int[]>>[] ForCampare_Yield_Fro_DB_List { get; set; }
            public List<List<List<List<int>[]>>>[] ForCampare_Yield_DB_LotVariation { get; set; }
            public List<List<List<int[]>>>[] ForCampare_Yield_Fro_DB_List_LotVariation { get; set; }

            public Dictionary<string, int> Refer_Site_And_Num { get; set; }
            public Dictionary<string, int> Refer_Lot_And_Num { get; set; }
            public List<int>[] ForCampare_Yield_List { get; set; }
            public List<List<int>[]> ForCampare_Yield_List1 { get; set; }
            public List<List<int>[]>[] ForCampare_Yield_List2 { get; set; }
            public Dictionary<string, Values> Values { get; set; }

            public Dictionary<string, Data_Calculation> Cal_Value_by_rowsdata { get; set; }
            public Dictionary<string, Data_Calculation> For_New_Spec_Cal_Value_by_rowsdata { get; set; }

            public List<double[]>[] DB_DataSet_Values { get; set; }

            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }
            public Stopwatch[] TestTime1 { get; set; }
            public Stopwatch[] TestTime2 { get; set; }
            public Stopwatch[] TestTime3 { get; set; }
            public Stopwatch[] TestTime4 { get; set; }
            public Stopwatch[] TestTime5 { get; set; }

            public Dictionary<string, int> Lot_Dic { get; set; }
            public Dictionary<string, int> Site_Dic { get; set; }
            public Dictionary<string, int> Bin_Dic { get; set; }
            public Dictionary<string, Dictionary<string, List<string>>> Matching_Lots { get; set; }
            public Dictionary<string, List<string>> Matching_Lot { get; set; }
            public long SampleCount { get; set; }
            public object Update_Data_ID { get; set; }
            public string[] Update_Datas_ID { get; set; }
            public string Get_Gross_Para { get; set; }
            public Dictionary<string, IQR> DIC_IQR { get; set; }
            public double Get_Gross_Persent { get; set; }
            public string Get_Gross_Selector { get; set; }
            public List<List<int>[]>[] ForCampare_Yield { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent_For_New_Spec { get; set; }
            public List<List<int>>[] For_Any_Yield_For_New_Spec { get; set; }
            public List<List<int>[]>[] For_New_Spec_ForCampare_Yield { get; set; }
            public List<Dictionary<string, Gross>[]> List_Gross_Values { get; set; }
            public List<int>[] Check { get; set; }
            public List<List<int>[]> Test { get; set; }
            public Dictionary<string, Gross>[] Gross_Values1 { get; set; }
            public Dictionary<string, CSV_Class.For_Box> Dic_Test_For_Spec_Gen { get; set; }
            public Dictionary<string, CSV_Class.For_Box>[] Dic_Test { get; set; }
            public string Table { get; set; }

            public double[] Make_New_Spec_For_Yield_Min { get; set; }
            public double[] Make_New_Spec_For_Yield_Max { get; set; }

            public List<string> Gross { get; set; }

            public List<string[]>[] DataSet_Value { get; set; }
            public long NB { get; set; }
            public List<double[]>[] DataSet_Double_Value { get; set; }

            public int[] Each_Thread_Count { get; set; }


            public string Lot_ID { get; set; }
            public string SubLot_ID { get; set; }
            public string Tester_ID { get; set; }
            public string Site { get; set; }
            public string Bin { get; set; }
            public int Bin_place { get; set; }
            public string ID_Unit { get; set; }
            public string Query { get; set; }
            public bool _From_Db { get; set; }
            public object[] Std_Value { get; set; }
            public double[] Std_Value_Convert { get; set; }
            public int Spec_Table_Count { get; set; }
            public bool _Flag { get; set; }
            public bool _SUBLOT_Flag { get; set; }
            public bool Clotho_Spec_Flag { get; set; }
            public string Before_Lot_ID { get; set; }
            public string Changed_Lot_ID { get; set; }
            
            public string[] No_Index { get; set; }
            public string[] Paraname { get; set; }
            public string[] SpecMin { get; set; }
            public string[] SpecMax { get; set; }
            public string[] DataMin { get; set; }
            public  string[] DataMedian { get; set; }
            public  string[] DataMax { get; set; }
            public  string[] CPK { get; set; }
            public  string[] STD { get; set; }
            public  string[] Percent { get; set; }
            public  string[] Fail { get; set; }

            public string[] Line { get; set; }

            public int Count_Current_Setting { get; set; }

            public void Open_DB(string FileName, Data_Class.Data_Editing.INT Data_Edit)
            {
                string Filename = FileName.Substring(FileName.LastIndexOf("\\") + 1);
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\FCM\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA SCHEMA.SYNCHRONOUS = OFF; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; DEBUG = 1;Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\FCM\\" + Filename + "_" + i + ".db;  PRAGMA LOCKING_MODE = EXCLUSIVE; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    conn[i].Open();
                }




            }

            public void Open_DB(string[] FileName, Data_Class.Data_Editing.INT Data_Edit)
            {

                Data_Edit.DB_Count = FileName.Length;
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    string Filename = FileName[i].Substring(FileName[i].LastIndexOf("\\") + 1);

                    int length = Filename.Length;
                    Filename = Filename.Substring(0, length - 5);

                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "\\" + Filename + i + ".db";
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + ".csv\\" + Filename.Substring(0, Filename.Length) + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    //cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    //cmd[i].ExecuteNonQuery();

                }



            }
            public void DropTable(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {
                try
                {
                    for (int i = 0; i < Data_Edit.DB_Count; i++)
                    {
                        cmd[i].CommandText = "";
                        cmd[i].CommandText = "drop TABLE data";
                        cmd[i].ExecuteNonQuery();
                    }
                }
                catch { }

            }

            public void Insert_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                //Teststring = new string[7][];

                //Teststring[0] = new string[Data.DB_Column_Limit];
                //Teststring[1] = new string[Data.DB_Column_Limit];
                //Teststring[2] = new string[Data.DB_Column_Limit];
                //Teststring[3] = new string[Data.DB_Column_Limit];
                //Teststring[4] = new string[Data.DB_Column_Limit];
                //Teststring[5] = new string[Data.DB_Column_Limit];
                //Teststring[6] = new string[Data.Per_DB_Column_Count[6]];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    // cmd[i].CommandText = "";
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(MakecolumnsThread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void MakecolumnsThread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " VARCAHR2(100)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR2(100)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR2(100)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(");");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    // cmd[i].CommandText = "";
                    stringA[i] = new StringBuilder();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }
            public void Insert_Current_Setting(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void Insert_Current_Setting_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {
                Data = Data_Edit;
                this.Table = Table;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  cmd[i].Reset();
                    //    ThreadFlags[i] = new ManualResetEvent(false);
                    Insert_Current_Setting_Data_Thread(i);
                    //  ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //       Wait[i] = ThreadFlags[i].WaitOne();
                }

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Clotho_Spec_Max_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}

            }

            public void Insert_Current_Setting_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();
                stringA[i].Clear();
                SampleCount = 1;

                cmd[i] = new SQLiteCommand(conn[i]);

                int Count = Data.Per_DB_Column_Count[i];


                int k = 0;


                if (Table.ToUpper() == "CLOTHO_SPEC")
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Min[Spec_Count] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Min[Spec_Count] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                    }




                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }
                else
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[0] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Min[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;

                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }
                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }




                //   ThreadFlags[i].Set();
            }
            public void Insert_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE spec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE spec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY );");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_New_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    // cmd[i].CommandText = "";
                    stringA[i] = new StringBuilder();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_New_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void Insert_New_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE newspec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE newspec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY );");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Data(Data_Class.Data_Editing.INT Data_Edit)
            {
                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                // sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                tran = new SQLiteTransaction[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];


                //Testdouble = new double[7][];

                //Testdouble[0] = new double[Data.DB_Column_Limit];
                //Testdouble[1] = new double[Data.DB_Column_Limit];
                //Testdouble[2] = new double[Data.DB_Column_Limit];
                //Testdouble[3] = new double[Data.DB_Column_Limit];
                //Testdouble[4] = new double[Data.DB_Column_Limit];
                //Testdouble[5] = new double[Data.DB_Column_Limit];
                //Testdouble[6] = new double[Data.Per_DB_Column_Count[6]];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();
                }

            }

            public void Insert_Ref_Header_Data(Data_Class.Data_Editing.INT Data_Edit)
            {


            }
            public void Insert_Data(long Sample)
            {

            }
            public void Insert_Data_Get_From_DB(int Sample)
            {

            }
            public void Insert_Spec_Get_From_DB(Data_Class.Data_Editing.INT Data_Edit)
            {


                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Get_From_DB_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }

            public void Insert_Spec_Get_From_DB_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Data(string Tablename)
            {

            }

            public void Insert_Spec_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {

            }
            /// <summary>
            /// 
            /// </summary>
            /// <param name="Tablename"></param>
            public void Insert_Files_Name(string Tablename)
            {

                //Table = Tablename;
                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}


            }
            public void Make_table(string Tablename)
            {

            }
            public void Make_table2(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {

            }
            public void Make_table_For_Filename(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {
                Data = Data_Edit;
                Table = Tablename;

                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(_Make_Table_For_Filename), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void _Make_Table_For_Filename(Object threadContext)
            {
                int i = (int)threadContext;

                stringA[i].Append("CREATE TABLE " + Table + "(FIle VARCAHR(20))");


                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].CommandText = "";

                stringA[i].Append(",");

                ThreadFlags[i].Set();
            }

            public void Make_table_For_Trace(string Tablename, string Chan, bool Flag)
            {
                stringA[0].Clear();
                stringA[0].Append("CREATE TABLE " + Tablename + "( FIRST VARCAHR(5), END VARCAHR(5), DBCOUNT VARCHAR(5), COLUMNCOUNT VARCHAR(5) );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";

                stringA[0].Clear();
                stringA[0].Append("INSERT INTO INF VALUES ('" + TheFirst_Trashes_Header_Count + "' , '" + TheEnd_Trashes_Header_Count + "' , '" + Data.Per_DB_Column_Count.Length + "' , '" + Data.Per_DB_Column_Count[Data.Per_DB_Column_Count.Length - 1] + "' );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";
            }
            public void Delete_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Delete_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Delete_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;
                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[0] + "',");

                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_LowSpec[Data.DB_Column_Limit * i + k] + "', '" + SampleCount + "');");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                Thread.Sleep(100);
                stringA[i].Clear();
                cmd[i].Reset();
                k = 0;
                SampleCount = 2;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[0] + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_HighSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_HighSpec[Data.DB_Column_Limit * i + k] + "', '" + SampleCount + "');");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }
            public void Delete_Lot_Data(string Query)
            {

            }
            public void Insert_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                int k = 0;
                for (k = 0; k < Count - 1; k++)
                {
                    if (k == 0)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO data VALUES ('" + Data.New_Data[0] + "',");
                            // Testdouble[i][0] = Data.New_Data[0];
                            //  ForCampare_Yield_List[0][0] = 0;
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO data VALUES ('" + Data.New_Data[Data.DB_Column_Limit * i] + "',");
                            // Testdouble[i][0] = Data.New_Data[Data.DB_Column_Limit * i];
                            //if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Data.New_Data[Data.DB_Column_Limit * i] || Data.New_LowSpec[Data.DB_Column_Limit * i] > Data.New_Data[Data.DB_Column_Limit * i])
                            //{
                            //    ForCampare_Yield_List[i][0] = 1;
                            //}
                        }

                    }
                    else
                    {
                        stringA[i].Append("'" + Data.New_Data[Data.DB_Column_Limit * i + k] + "',");
                        // Testdouble[i][j] = Data.New_Data[Data.DB_Column_Limit * i + j];
                        //if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Data.New_Data[Data.DB_Column_Limit * i + k] || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Data.New_Data[Data.DB_Column_Limit * i + k])
                        //{
                        //    ForCampare_Yield_List[i][k] = 1;
                        //}
                    }
                }

                stringA[i].Append("'" + Data.New_Data[Data.DB_Column_Limit * i + k] + "');");
                //  Testdouble[i][j] = Data.New_Data[Data.DB_Column_Limit * i + j];

                //if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Data.New_Data[Data.DB_Column_Limit * i + k] || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Data.New_Data[Data.DB_Column_Limit * i + k])
                //{
                //    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                //}

                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].Reset();

                ThreadFlags[i].Set();
            }

            public void Save_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                //Update_Data_ID = data;

                //if (data != null)
                //{
                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        ThreadFlags[i] = new ManualResetEvent(false);
                //        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                //    }

                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        Wait[i] = ThreadFlags[i].WaitOne();
                //    }
                //}

            }

            public void Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                //Update_Data_ID = data;

                //if (data != null)
                //{
                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        ThreadFlags[i] = new ManualResetEvent(false);
                //        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                //    }

                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        Wait[i] = ThreadFlags[i].WaitOne();
                //    }
                //}

            }

            public void Road_Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                //  SampleCount = Sample;

                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Road_Save_Customer_Spec_table_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Road_Save_Customer_Spec_table_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }


            public void LOTID_Update(string Query, string Query2, string CellID)
            {

            }

            public void Gross_Update_Data(object data)
            {

            }

            public void Gross_Update_Datas(List<string> data)
            {

            }

            public void Chnaged_Spec_Update_Data(int DB, int Index, string Parameter, double Spec, int GetId)
            {

            }
            public Dictionary<string, double[]> Chnaged_Spec_Anl_Yield(int DB, int Index, string Parameter)
            {
                Dictionary<string, double[]> Dic_Change_Spec = new Dictionary<string, double[]>();

                return Dic_Change_Spec;
            }
            public void Get_Ave_Data(Data_Class.Data_Editing.INT Data_Edit)
            {
                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                //sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                ds = new DataSet[Data.DB_Count];

                ds[0] = new DataSet();
                // sqlAdapter[i] = new SQLiteDataAdapter();
                stringA[0] = new StringBuilder();
                ThreadFlags[0] = new ManualResetEvent(false);

                // stringA[0].Append("Select " + Select_Para + " from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }
            }
            public void Get_Ave_Data_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {
                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                //sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                ds = new DataSet[Data.DB_Count];

                ds[0] = new DataSet();
                // sqlAdapter[i] = new SQLiteDataAdapter();
                stringA[0] = new StringBuilder();
                ThreadFlags[0] = new ManualResetEvent(false);

                // stringA[0].Append("Select " + Select_Para + " from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }
            }
            public void Set_Refer_for_Anlyzer(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_Ave_Data2(Data_Class.Data_Editing.INT Data_Edit)
            {
                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                //sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                ds = new DataSet[Data.DB_Count];

                ds[0] = new DataSet();
                // sqlAdapter[i] = new SQLiteDataAdapter();
                stringA[0] = new StringBuilder();
                ThreadFlags[0] = new ManualResetEvent(false);

                // stringA[0].Append("Select " + Select_Para + " from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }
            }

            public void Get_Saved_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {




            }

            public List<double[]> Get_Saved_Spec_Thread(Object threadContext)
            {
                return null;
            }
            public void Get_Rows_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface)
            {


            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface, DataTable dt)
            {
                //stringA[DB].Clear();
                //stringA[DB].Append("Select id, " + Select_Para + " from data");

                //cmd[DB].CommandText = stringA[DB].ToString();
                //ds[DB] = new DataSet();

                //sqlAdapter[DB].SelectCommand = cmd[DB];
                //sqlAdapter[DB].Fill(ds[DB]);

                //ID = new object[ds[DB].Tables[0].Rows.Count];
                //Value = new object[ds[DB].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //{
                //    ID[count] = dr.ItemArray[0];
                //    Value[count] = dr.ItemArray[1];

                //    count++;
                //}

                //double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);


                //stringA[DB].Clear();
            }
            public void Get_Selected_Para(int DB, string Select_Para, bool Flag, string Selector)
            {


            }
            public double[] Get_Find_Bin(string Query)
            {
                return null;
            }
            public List<object[]> Get_Data_By_Querys(string Query)
            {
                return null;
            }
            public string[] Get_Data_By_Query(string Query)
            {

                stringA[0].Clear();
                stringA[0].Append(Query);
     

                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[500000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    Value1[count] = stringD[0];

                    count++;

                }

                Array.Resize(ref Value1, count);

                cmd[0].Dispose();
                SqReader[0].Close();

                string[] _string = Array.ConvertAll<object, string>(Value1, Convert.ToString);


                stringA[0].Clear();

                return _string;
            }
            public Dictionary<string, object[]> Get_Data_By_Query_S4PD(string Query, string Chan)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                // cmd[0].CommandText = stringA[0].ToString();
                // ds[0] = new DataSet();

                // sqlAdapter[0].SelectCommand = cmd[0];
                // sqlAdapter[0].Fill(ds[0]);

                // Value = new object[ds[0].Tables[0].Rows.Count];

                //// int count = 0;
                // foreach (DataRow dr in ds[0].Tables[0].Rows)
                // {
                //     Value[count] = dr.ItemArray[0];
                //     count++;
                // }

                //  string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);
                // SqReader[0] = cmd[0].ExecuteReader();
                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[500000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    Value1[count] = stringD[0];

                    count++;

                }

                Array.Resize(ref Value1, count);

                cmd[0].Dispose();
                SqReader[0].Close();

                string[] _string = Array.ConvertAll<object, string>(Value1, Convert.ToString);


                stringA[0].Clear();
                return null;
            }
            public string[] Get_Data_By_Query(string Query, int DB)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);

                stringA[0].Clear();
                return _string;
            }

            public void Get_Defined_Para(object[,] DummyData, string key, Data_Class.Data_Editing.INT Data_InterFace)
            {

                Data = Data_InterFace;
                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                tran = new SQLiteTransaction[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];
                ds = new DataSet[Data.DB_Count];


                ds[0] = new DataSet();
                sqlAdapter[0] = new SQLiteDataAdapter();
                stringA[0] = new StringBuilder();

                //SampleCount = 0;
                //for (int k_NB = 0; k_NB < Table_Count; k_NB++)
                //{
                //    Query = "Select count(id) from data" + k_NB + " where Fail = '0'";

                //    SampleCount += Get_Sample_Count(null, Query);
                //}


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ds[i] = new DataSet();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder();
                    ThreadFlags[i] = new ManualResetEvent(false);

                    Data_Information s = new Data_Information(DummyData, key, i, Data_InterFace.Reference_Header, Data_InterFace.New_HighSpec, Data_InterFace.New_LowSpec);
                    Find_Para_Thread(s);

                 //   ThreadPool.QueueUserWorkItem(new WaitCallback(Find_Para_Thread), new Data_Information(DummyData, key, i, Data_InterFace.Reference_Header, Data_InterFace.New_HighSpec, Data_InterFace.New_LowSpec));
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                 //   Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();
                    //cmd[i].CommandText = "";
                }

            }
            public void Get_Current_Setting(Data_Class.Data_Editing.INT Data_Edit, int NB)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                 //   Get_From_Db_Data(i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
               //     Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Find_Para_Thread(Object threadContext)
            {
                Data_Information argArray = (Data_Information)threadContext;

                int i = (int)argArray.DB_NB;
                int k = 0;
                object[,] data = (object[,])argArray.DummyData;
                string Key = (string)argArray.Key;
                string[] Ref_Header = (string[])argArray.Ref_Header;
                double[] High_Spec = (double[])argArray.High_Spec;
                double[] Low_Spec = (double[])argArray.Low_Spec;


       

    

                int Count = 0;

                if( i == 0)
                {
                    Count = 1993;
                }
                else if( i == Data.DB_Count - 1)
                {
                    Count = Data.Per_DB_Column_Count[i] - 9;
                }
                else
                {
                    Count = 1993;
                }

                for (k = 0; k < Count; k++)
                {
                    string[] Split = Ref_Header[Data.DB_Column_Limit * i + k].Split('_');
                    if (Split[Split.Length - 1].ToUpper() == data[0, 0].ToString().ToUpper())
                    {
                        int Find_DB = 0;
                        if (k + (Data.Per_DB_Column_Count[0] * i) > Data.DB_Column_Limit - 10)
                        {
                            for (int kk = 0; kk < Data.DB_Count; kk++)
                            {
                                if (k + (Data.Per_DB_Column_Count[0] * i) <= Data.Per_DB_Column_Count_End[kk] - 9)
                                {
                                    Find_DB = kk;
                                    break;
                                }
                            }
                        }
                        object[] rawdata = new object[0];

                        this.ID = new object[0];
                        this.WAFER_ID = new object[0];
                        this.LOT_ID = new object[0];
                        this.SITE_ID = new object[0];
                        this.Value = new object[0];

                        object[] ID_Dummy = new object[0];
                        object[] WAFERID_Dummy = new object[0];
                        object[] LOTID_Dummy = new object[0];
                        object[] SITEID_Dummy = new object[0];
                        object[] Value_Dummy = new object[0];

                        foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                        {

                            int count = 0;
                            Dictionary<string, List<string>> tests = key.Value;

                            object[] rawdatas = new object[0];
                            foreach (KeyValuePair<string, List<string>> ts in tests)
                            {

                                if (Find_DB == 0)
                                {
                                    cmd[Find_DB] = new SQLiteCommand(conn[Find_DB]);
                                    sqlAdapter[Find_DB] = new SQLiteDataAdapter();

                                    stringA[Find_DB] = new StringBuilder();
                                    ds[Find_DB] = new DataSet();

                                    stringA[Find_DB].Append("Select id, WAFER_ID, LOTID, SITEID, " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from " + key.Key + " where Fail not like '1'");
                                    //  stringA[Find_DB].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from " + key.Key + "  where Fail not like '1'");

                                    cmd[Find_DB].CommandText = stringA[Find_DB].ToString();
                                    sqlAdapter[Find_DB].SelectCommand = cmd[Find_DB];
                                    sqlAdapter[Find_DB].Fill(ds[Find_DB]);

                                    ID_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];
                                    WAFERID_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];
                                    LOTID_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];
                                    SITEID_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];
                                    Value_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];

                                    foreach (DataRow dr in ds[Find_DB].Tables[0].Rows)
                                    {
                                        ID_Dummy[count] = dr.ItemArray[0];
                                        WAFERID_Dummy[count] = dr.ItemArray[1];
                                        LOTID_Dummy[count] = dr.ItemArray[2];
                                        SITEID_Dummy[count] = dr.ItemArray[3];
                                        Value_Dummy[count] = dr.ItemArray[4];

                                        count++;
                                    }


                                    stringA[Find_DB] = new StringBuilder();
                                    sqlAdapter[Find_DB].Dispose();
                                    cmd[Find_DB].Dispose();
                                    ds[Find_DB] = new DataSet();


                                }
                                else
                                {
                                    cmd[Find_DB] = new SQLiteCommand(conn[Find_DB]);
                                    sqlAdapter[Find_DB] = new SQLiteDataAdapter();

                                    stringA[Find_DB] = new StringBuilder();
                                    ds[Find_DB] = new DataSet();

                                    stringA[Find_DB].Append("Select id, LOTID, SITEID, " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from " + key.Key + " where Fail not like '1'");
                                    //  stringA[Find_DB].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from " + key.Key + "  where Fail not like '1'");

                                    cmd[Find_DB].CommandText = stringA[Find_DB].ToString();
                                    sqlAdapter[Find_DB].SelectCommand = cmd[Find_DB];
                                    sqlAdapter[Find_DB].Fill(ds[Find_DB]);

                                    ID_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];
                                    LOTID_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];
                                    SITEID_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];
                                    Value_Dummy = new object[ds[Find_DB].Tables[0].Rows.Count];

                                    foreach (DataRow dr in ds[Find_DB].Tables[0].Rows)
                                    {
                                        ID_Dummy[count] = dr.ItemArray[0];
                                        LOTID_Dummy[count] = dr.ItemArray[1];
                                        SITEID_Dummy[count] = dr.ItemArray[2];
                                        Value_Dummy[count] = dr.ItemArray[3];

                                        count++;
                                    }

                                    count = 0;

                                    stringA[0].Clear();
                                    stringA[0].Append("Select WAFER_ID from " + key.Key + " where Fail not like '1'");

                                    cmd[0] = new SQLiteCommand(conn[0]);
                                    sqlAdapter[0] = new SQLiteDataAdapter();


                                    cmd[0].CommandText = stringA[0].ToString();
                                    ds[0] = new DataSet();

                                    sqlAdapter[0].SelectCommand = cmd[0];
                                    sqlAdapter[0].Fill(ds[0]);

                                    WAFERID_Dummy = new object[ds[0].Tables[0].Rows.Count];


                                    foreach (DataRow dr in ds[0].Tables[0].Rows)
                                    {
                                        WAFERID_Dummy[count] = dr.ItemArray[0];
                                        count++;
                                    }


                                    stringA[Find_DB] = new StringBuilder();
                                    sqlAdapter[Find_DB].Dispose();
                                    cmd[Find_DB].Dispose();
                                    ds[Find_DB] = new DataSet();


                                }
                            }

                            this.ID = this.ID.Concat(ID_Dummy).ToArray();
                            this.WAFER_ID = this.WAFER_ID.Concat(WAFERID_Dummy).ToArray();
                            this.LOT_ID = this.LOT_ID.Concat(LOTID_Dummy).ToArray();
                            this.SITE_ID = this.SITE_ID.Concat(SITEID_Dummy).ToArray();
                            this.Value = this.Value.Concat(Value_Dummy).ToArray();
                        }

                        if (Find_DB != 0)
                        {
                            cmd[0].Dispose();
                            sqlAdapter[0].Dispose();
                        }
                        cmd[Find_DB].Dispose();
                        sqlAdapter[Find_DB].Dispose();

                        double[] Data1 = Array.ConvertAll<object, double>(this.Value, Convert.ToDouble);
                        string[] ID = Array.ConvertAll<object, string> (this.ID, Convert.ToString);
                        string[] WAFER_ID = Array.ConvertAll<object, string>(this.WAFER_ID, Convert.ToString);
                        string[] LOT_ID = Array.ConvertAll<object, string>(this.LOT_ID, Convert.ToString);
                        string[] SITE_ID = Array.ConvertAll<object, string>(this.SITE_ID, Convert.ToString);

                        string Apple_Min = "";
                        string Apple_Max = ""; ;

                        string B_Min = Convert.ToString(Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Min[0]);
                        string B_Max = Convert.ToString(Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Max[0]);

                        
                        double Min = Data1.Min();
                        double Max = Data1.Max();

                        CSV_Class.For_Box Box = new CSV_Class.For_Box(Data.Reference_Header[Data.DB_Column_Limit * i + k], Data1, ID, WAFER_ID, SITE_ID, LOT_ID, Min, Max, "0", "0", "", Apple_Min, Apple_Max, B_Min, B_Max);
                        Dic_Test[0].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], Box);

                        //CSV_Class.For_Box Box = new CSV_Class.For_Box(Data.Reference_Header[i], Data1, ID, WAFER_ID, SITE_ID, LOT_ID, 0f, 0f, "0", "0", "", Convert.ToString(Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k].Min[0]), Convert.ToString(Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k].Max[0]), Convert.ToString(Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Min[0]), Convert.ToString(Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Max[0]));


                        //Values Inf = new Values(rawdata, Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Min[0], Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Max[0], Key);

                        //Values.Add(Ref_Header[Data.DB_Column_Limit * i + k], Inf);



                    }
                }
                ThreadFlags[i].Set();
            }

            public void Get_Gross_Check_Para(Data_Class.Data_Editing.INT Data_Edit, string Select_Para, double Persent, string Selector, int SelectedBin)
            {

            }
            public void Get_From_Db_Data_for_Anly(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_From_Db_Data_for_Anly_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_From_Db_Ref_Header(Data_Class.Data_Editing.INT Data_Edit)
            {



            }
            public void Get_From_Db_Ref_Header_Thread(Object threadContext)
            {



            }
            public int Get_Sample_Count(int DB, string Query)
            {

                stringA[0].Clear();
                stringA[0].Append(Query);



                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                }

                //   sqlAdapter[0].Dispose();
                //   cmd[0].Dispose();

                //   conn[0].Dispose();

                //sqlAdapter[0].Dispose();
                //stringA[0].Clear();

                //   cmd[0].Dispose();
                // conn[0].Close();


                int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Data_Count[0];

            }
            public int Get_Column_Count(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                }

                sqlAdapter[0].Dispose();
                cmd[0].Dispose();

                int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Data_Count[0];
            }
            public void Close(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    conn[i].Close();
                }
            }
            public void Read_Dispose(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }

            public void Set_Conn(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }
            public void trans(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait = new bool[Data.DB_Count];
                    stringA = new StringBuilder[Data.DB_Count];
                    tran = new SQLiteTransaction[Data.DB_Count];
                    tran[i] = conn[i].BeginTransaction();
                    cmd[i].Transaction = tran[i];
                }
            }

            public void Commit(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    tran[i].Commit();
                }
            }
            public string Get_Data_From_Table(string Table, string header)
            {

                stringA[0].Clear();
                stringA[0].Append("select " + header + " from " + Table);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }
                stringA[0].Clear();

                return Convert.ToString(Value[0]);
            }

        }

        public class Yield_DB : INT
        {
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ReaderWriterLockSlim[] sqlitelock { get; set; }
            public string[] strConn { get; set; }
            public SQLiteConnection[] conn { get; set; }
            public SQLiteCommand[] cmd { get; set; }

            public SQLiteDataAdapter[] sqlAdapter { get; set; }
            public SQLiteCommandBuilder[] sqlcmdbuilder { get; set; }
            public SQLiteDataReader[] SqReader { get; set; }

            public DbDataReader[] DbReader { get; set; }
            public DataSet[] ds { get; set; }
            public DataTable dt_test { get; set; }
            public DataTable[] dt { get; set; }
            public SQLiteTransaction[] tran { get; set; }

            public ManualResetEvent[] ThreadFlags { get; set; }
            public ManualResetEvent[] Insert_ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }
            public string Filename { get; set; }
            public int Limit { get; set; }
            public int Limit_Count { get; set; }
            public int Table_Count { get; set; }
            public bool[] Insert_Thread_Wait { get; set; }
            public double[] Testtime { get; set; }

            public double[][] test { get; set; }

            double[] Testtime1 { get; set; }
            double[] Testtime2 { get; set; }
            double[] Testtime3 { get; set; }
            public string[][] Teststring { get; set; }
            public double[][] Testdouble { get; set; }

            public object[] ID { get; set; }
            public object[] Value { get; set; }
            public object[] WAFER_ID { get; set; }
            public object[] LOT_ID { get; set; }
            public object[] SITE_ID { get; set; }
            public Dictionary<string, double[]> Selected_Parameter_Distribution { get; set; }

            public object[] Variation { get; set; }
            public Dictionary<string, IQR> DIC_IQR { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test_New_Spec { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent { get; set; }
            public List<List<int>>[] For_Any_Yield { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_Lot { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_SITE { get; set; }
            public List<List<int>[]>[] ForCampare_Yield { get; set; }

            public List<List<int>[]>[] For_Any_Yield_Percent_For_New_Spec { get; set; }
            public List<List<int>>[] For_Any_Yield_For_New_Spec { get; set; }
            public List<List<int>[]>[] For_New_Spec_ForCampare_Yield { get; set; }
            public List<List<int>>[] For_New_Spec_ForCampare_Yield2 { get; set; }

            public List<List<List<List<int>[]>>>[] ForCampare_Yield_DB_LotVariation { get; set; }

            public List<int[]>[] ForCampare_Yield_Fro_DB { get; set; }
            public List<List<int[]>>[] ForCampare_Yield_Fro_DB_List { get; set; }

            public List<List<List<int[]>>>[] ForCampare_Yield_Fro_DB_List_LotVariation { get; set; }
            public Dictionary<string, int> Refer_Site_And_Num { get; set; }
            public Dictionary<string, int> Refer_Lot_And_Num { get; set; }
            public List<int>[] ForCampare_Yield_List { get; set; }
            public List<List<int>[]> ForCampare_Yield_List1 { get; set; }

            public List<List<int>[]>[] ForCampare_Yield_List2 { get; set; }
            public Dictionary<string, Values> Values { get; set; }
            public Dictionary<string, Data_Calculation> Cal_Value_by_rowsdata { get; set; }
            public Dictionary<string, Data_Calculation> For_New_Spec_Cal_Value_by_rowsdata { get; set; }

            public List<double[]>[] DB_DataSet_Values { get; set; }

            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public Dictionary<string, CSV_Class.For_Box>[] Dic_Test { get; set; }
            public Dictionary<string, int> Lot_Dic { get; set; }
            public Dictionary<string, int> Site_Dic { get; set; }
            public Dictionary<string, int> Bin_Dic { get; set; }
            public Dictionary<string, Dictionary<string, List<string>>> Matching_Lots { get; set; }
            public Dictionary<string, List<string>> Matching_Lot { get; set; }
            public Stopwatch[] TestTime1 { get; set; }
            public Stopwatch[] TestTime2 { get; set; }
            public Stopwatch[] TestTime3 { get; set; }
            public Stopwatch[] TestTime4 { get; set; }
            public Stopwatch[] TestTime5 { get; set; }
            public long SampleCount { get; set; }
            public object Update_Data_ID { get; set; }
            public string[] Update_Datas_ID { get; set; }
            public string Get_Gross_Para { get; set; }
            public double Get_Gross_Persent { get; set; }
            public string Get_Gross_Selector { get; set; }
            public int Get_Gross_Selectedbin { get; set; }
            public Dictionary<string, CSV_Class.For_Box> Dic_Test_For_Spec_Gen { get; set; }
            public List<Dictionary<string, Gross>[]> List_Gross_Values { get; set; }
            public Dictionary<string, Gross>[] Gross_Values1 { get; set; }
            public long NB { get; set; }
            public List<int>[] Check { get; set; }
            public List<List<int>[]> Test { get; set; }
            public string Table { get; set; }
            public object[] Std_Value { get; set; }
            public double[] Std_Value_Convert { get; set; }
            public double[] Make_New_Spec_For_Yield_Min { get; set; }
            public double[] Make_New_Spec_For_Yield_Max { get; set; }
            public List<string> Gross { get; set; }
            public List<string[]>[] DataSet_Value { get; set; }
            public List<double[]>[] DataSet_Double_Value { get; set; }

            public int[] Each_Thread_Count { get; set; }

            public string Lot_ID { get; set; }
            public string SubLot_ID { get; set; }
            public string Tester_ID { get; set; }
            public string Site { get; set; }
            public string Bin { get; set; }
            public int Bin_place { get; set; }
            public string ID_Unit { get; set; }
            public string Query { get; set; }
            public bool _From_Db { get; set; }
            public int Spec_Table_Count { get; set; }
            public bool _Flag { get; set; }
            public bool _SUBLOT_Flag { get; set; }
            public bool Clotho_Spec_Flag { get; set; }
            public string Before_Lot_ID { get; set; }
            public string Changed_Lot_ID { get; set; }

            public string[] No_Index { get; set; }
            public string[] Paraname { get; set; }
            public string[] SpecMin { get; set; }
            public string[] SpecMax { get; set; }
            public string[] DataMin { get; set; }
            public string[] DataMedian { get; set; }
            public string[] DataMax { get; set; }
            public string[] CPK { get; set; }
            public string[] STD { get; set; }
            public string[] Percent { get; set; }
            public string[] Fail { get; set; }

            public string[] Line { get; set; }

            public int Count_Current_Setting { get; set; }

            public void Open_DB(string FileName, Data_Class.Data_Editing.INT Data_Edit)
            {
                string Filename = FileName.Substring(FileName.LastIndexOf("\\") + 1);
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    cmd[i].ExecuteNonQuery();

                }


                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                dt = new DataTable[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    dt[i] = new DataTable();
                    cmd[i].CommandText = "PRAGMA synchronous";
                    SqReader[i] = cmd[i].ExecuteReader();
                    dt[i].Load(SqReader[i]);
                }

                ForCampare_Yield_List = new List<int>[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data_Edit.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }
                ForCampare_Yield_List1 = new List<List<int>[]>();

            }
            public void Open_DB(string[] FileName, Data_Class.Data_Editing.INT Data_Edit)
            {

                Data_Edit.DB_Count = FileName.Length;
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    string Filename = FileName[i].Substring(FileName[i].LastIndexOf("\\") + 1);

                    int length = Filename.Length;
                    Filename = Filename.Substring(0, length - 5);

                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "\\" + Filename + i + ".db";
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + ".csv\\" + Filename.Substring(0, Filename.Length) + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    //cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    //cmd[i].ExecuteNonQuery();

                }



            }
            public void DropTable(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {

                try
                {
                    for (int i = 0; i < Data_Edit.DB_Count; i++)
                    {
                        cmd[i] = new SQLiteCommand(conn[i]);
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        cmd[i].Dispose();
                    }
                }
                catch
                {

                }


            }
            public void Insert_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(MakecolumnsThread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void MakecolumnsThread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("CREATE TABLE IF Not Exists " + Data.Data_Table + "(" + Data.New_Header[0] + " VARCAHR(5)");
                        }
                        else
                        {
                            stringA[i].Append("CREATE TABLE IF Not Exists " + Data.Data_Table + "(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                        }

                    }
                    else
                    {
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY, Fail VARCHAR(5) , LOT_ID VARCHAR(5), SUBLOT_ID VARCHAR(5), BIN VARCHAR(5) , SITE VARCHAR(5), TESTER VARCHAR(5));");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }
            public void Insert_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                //  Insert_Spec_Header_Thread(0);
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                  
                    //  ThreadFlags[i] = new ManualResetEvent(false);
                    Insert_Spec_Header_Thread(i);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //    stringA[i] = new StringBuilder();

                }

            }
            public void Insert_Current_Setting(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;
                Data.Data_Table = "Current_Setting";
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
             
                    stringA[i].Clear();
                    Insert_Current_Setting_Thread(i);
                    //ThreadFlags[i] = new ManualResetEvent(false);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Current_Setting_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                  //  Wait[i] = ThreadFlags[i].WaitOne();
                  //  stringA[i] = new StringBuilder();

                }
            }
            public void Insert_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                cmd[i] = new SQLiteCommand(conn[i]);
                stringA[i].Clear();


                if (i == 0)
                {
                    stringA[i].Append("CREATE TABLE IF Not Exists " + Data.Data_Table + "(" + Data.New_Header[0] + " VARCAHR(5),");
                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "'" + " VARCAHR(5),");
                    }

                    for (int j = 10; j < Count; j++)
                    {


                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j - 9] + " VARCHAR(5),");


                    }
                }
                else
                {
                    stringA[i].Append("CREATE TABLE IF Not Exists " + Data.Data_Table + "(" + Data.New_Header[Data.DB_Column_Limit * i - 9] + " VARCAHR(5),");

                    for (int j = 1; j < Count; j++)
                    {


                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j - 9] + " VARCHAR(5),");


                    }
                }

         

                stringA[i].Append("SubLot VARCAHR(5), id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");
                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].CommandText = "";
                cmd[i].Dispose();


            }
            public void Insert_Current_Setting_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                cmd[i] = new SQLiteCommand(conn[i]);
                stringA[i].Clear();


                if (i == 0)
                {
                    stringA[i].Append("CREATE TABLE IF Not Exists " + Data.Data_Table + "(" + Data.New_Header[0] + " VARCAHR(5),");
                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "'" + " VARCAHR(5),");
                    }

                    for (int j = 10; j < Count; j++)
                    {


                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j - 9] + " VARCHAR(5),");


                    }
                }
                else
                {
                    stringA[i].Append("CREATE TABLE IF Not Exists " + Data.Data_Table + "(" + Data.New_Header[Data.DB_Column_Limit * i - 9] + " VARCAHR(5),");

                    for (int j = 1; j < Count; j++)
                    {


                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j - 9] + " VARCHAR(5),");


                    }
                }



                stringA[i].Append("SubLot VARCAHR(5), id VARCAHR(5), LOTID VARCAHR(5), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");
                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].CommandText = "";
                cmd[i].Dispose();


            }
            public void Insert_Spec_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {
                Data = Data_Edit;
                this.Table = Table;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                    Insert_Spec_Data_Thread(i);
                  //  ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
             //       Wait[i] = ThreadFlags[i].WaitOne();
                }

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Clotho_Spec_Max_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}

            }
            public void Insert_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();
                stringA[i].Clear();
                SampleCount = 1;

                cmd[i] = new SQLiteCommand(conn[i]);

                int Count = Data.Per_DB_Column_Count[i];
    
               
                int k = 0;

     
                if(Table.ToUpper() == "CLOTHO_SPEC")
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Min[Spec_Count] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Min[Spec_Count] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                    }


 

                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }
                else
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[0] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Min[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;

                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }
                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }

              

 
             //   ThreadFlags[i].Set();
            }

            public void Insert_Current_Setting_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {
                Data = Data_Edit;
                this.Table = Table;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  cmd[i].Reset();
                   ThreadFlags[i] = new ManualResetEvent(false);
                    //Insert_Current_Setting_Data_Thread(i);
                     ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Current_Setting_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                     Wait[i] = ThreadFlags[i].WaitOne();
                }

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Clotho_Spec_Max_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}

            }

            public void Insert_Current_Setting_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();
                stringA[i].Clear();
                SampleCount = 1;

                cmd[i] = new SQLiteCommand(conn[i]);

                int Count = Data.Per_DB_Column_Count[i];

              //  Count = Count - 6;

                int k = 0;



                #region No_Idex

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + No_Index[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 10; k < Count; k++)
                    {
                        if(k == 1992)
                        {

                        }

                        stringA[i].Append("'" + No_Index[k - 9]  + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                  stringA[i].Clear();
                }

                else
                {


                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + No_Index[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + No_Index[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region Paraname

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + Paraname[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {
                        if (k == 1991)
                        {

                        }

                        stringA[i].Append("'" + Paraname[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + Paraname[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + Paraname[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region SpecMin

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + SpecMin[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + SpecMin[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + SpecMin[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + SpecMin[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region SpecMax

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + SpecMax[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + SpecMax[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + SpecMax[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + SpecMax[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region DataMin

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + DataMin[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + DataMin[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + DataMin[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + DataMin[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region DataMedian

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + DataMedian[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + DataMedian[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + DataMedian[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + DataMedian[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region DataMax

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + DataMax[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + DataMax[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + DataMax[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + DataMax[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region CPK

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + CPK[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + CPK[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + CPK[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + CPK[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region STD

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + STD[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + STD[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + STD[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + STD[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region Percent

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + Percent[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + Percent[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + Percent[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + Percent[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                #region Fail

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + Fail[0] + "',");

                    for (int p = 0; p < 9; p++)
                    {
                        stringA[i].Append("'" + p + "',");
                    }

                    for (k = 1; k < Count - 9; k++)
                    {

                        stringA[i].Append("'" + Fail[k] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }
                else
                {
                    stringA[i].Append("INSERT INTO Current_Setting VALUES ('" + Fail[Data.DB_Column_Limit * i - 9] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + Fail[Data.DB_Column_Limit * i + k - 9] + "',");

                    }


                    stringA[i].Append("'0','" + Table + "','0','0', '0', '0');");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                #endregion

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;





               ThreadFlags[i].Set();
            }

            public void Insert_Clotho_Spec_Min_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];
                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;
                int k = 0;
                // for (int j = 0; j < 1; j++)
                for (int j = 0; j < Data.Clotho_Spcc_List[0].Max.Length; j++)
                {
                    stringA[i].Clear();

                    if (i == 0)
                    {
                        stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Min[j] + "',");

                    }
                    else
                    {
                        stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i].Min[j] + "',");
                    }

                    for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                    {
                        stringA[i].Append("'" + Data.Clotho_Spcc_List[(Data.DB_Column_Limit * i) + k].Min[j] + "',");
                        //    stringA[i].Append("'" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + k] + "',");

                    }

                    stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Min[j] + "', '" + (j + 1) + "' ,0,0,0,0);");
                    //   stringA[i].Append("'" + Data.New_LowSpec[Data.DB_Column_Limit * i + k] + "', '" + SampleCount + "' ,0,0,0,0);");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    // cmd[i].Reset();
                    SampleCount++;

                }

                k = 0;
                // for (int j = 0; j < 1; j++)
                for (int j = 0; j < Data.Clotho_Spcc_List[0].Min.Length; j++)
                {
                    stringA[i].Clear();

                    if (i == 0)
                    {
                        stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Max[j] + "',");

                    }
                    else
                    {
                        stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i].Max[j] + "',");
                    }

                    for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                    {
                        stringA[i].Append("'" + Data.Clotho_Spcc_List[(Data.DB_Column_Limit * i) + k].Max[j] + "',");
                        //    stringA[i].Append("'" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + k] + "',");

                    }


                    stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Max[j] + "', '" + (Data.Clotho_Spcc_List[0].Min.Length + j + 1) + "' ,0,0,0,0);");
                    //   stringA[i].Append("'" + Data.New_LowSpec[Data.DB_Column_Limit * i + k] + "', '" + SampleCount + "' ,0,0,0,0);");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    //  cmd[i].Reset();
                    SampleCount++;

                }
                ThreadFlags[i].Set();
            }
            public void Insert_Clotho_Spec_Max_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;
                // for (int j = 0; j < 1; j++)
                for (int j = 0; j < Data.Clotho_List[0].Min.Length; j++)
                {
                    stringA[i].Clear();

                    if (i == 0)
                    {
                        stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_List[0].Max[j] + "',");

                    }
                    else
                    {
                        stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_List[Data.DB_Column_Limit * i].Max[j] + "',");
                    }

                    for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                    {
                        stringA[i].Append("'" + Data.Clotho_List[(Data.DB_Column_Limit * i) + k].Max[j] + "',");
                        //    stringA[i].Append("'" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + k] + "',");

                    }


                    stringA[i].Append("'" + Data.Clotho_List[Data.DB_Column_Limit * i + k].Max[j] + "', '" + (SampleCount) + "' ,0,0,0,0);");
                    //   stringA[i].Append("'" + Data.New_LowSpec[Data.DB_Column_Limit * i + k] + "', '" + SampleCount + "' ,0,0,0,0);");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    cmd[i].Reset();
                    SampleCount++;

                }
                ThreadFlags[i].Set();
            }
            public void Insert_New_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_New_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }
            public void Insert_New_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE IF Not Exists newspec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE IF Not Exists newspec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY, Fail VARCHAR(5) , LOT_ID VARCHAR(5) , SUBLOT_ID VARCHAR(5) , BIN VARCHAR(5));");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }
            public void Insert_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                // sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];
                sqlitelock = new ReaderWriterLockSlim[Data.DB_Count];
                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                //Testdouble = new double[7][];

                //Testdouble[0] = new double[Data.DB_Column_Limit];
                //Testdouble[1] = new double[Data.DB_Column_Limit];
                //Testdouble[2] = new double[Data.DB_Column_Limit];
                //Testdouble[3] = new double[Data.DB_Column_Limit];
                //Testdouble[4] = new double[Data.DB_Column_Limit];
                //Testdouble[5] = new double[Data.DB_Column_Limit];
                //Testdouble[6] = new double[Data.Per_DB_Column_Count[6]];
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder(100000);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder(100000);
                    Testtime[i] = TestTime1.Elapsed.TotalMilliseconds;
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);
            }
            public void Insert_Ref_Header_Data(Data_Class.Data_Editing.INT Data_Edit)
            {


            }
            public void Insert_Data(long Sample)
            {
                SampleCount = Sample;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                Insert_ThreadFlags[0].Set();
            }
            public void Insert_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[0].Replace("PID-", "") + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count] + "',");

                }

                for (k = 1; k < Count; k++)
                {
                    stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k] + "',");

                }

                stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k] + "', '" + SampleCount + "' , '0' , '" + Lot_ID + "' , '" + SubLot_ID + "' , '" + Bin + "' , '" + Site + "' , '" + Tester_ID + "' );");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Data_Get_From_DB(int Sample)
            {
                SampleCount = Sample;

                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Get_From_DB_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Insert_Data_Get_From_DB_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Spec_Get_From_DB(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  Insert_Spec_Get_From_DB_Thread(0);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Get_From_DB_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();

                }

            }

            public void Insert_Spec_Get_From_DB_Thread(Object threadContext)
            {
                int i = (int)threadContext;


                int count = 0;


                stringA[i].Clear();
                stringA[i].Append("Select * from " + this.Data.Data_Table);

                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                count = 0;

                while (SqReader[i].Read())
                {

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);


                    string[] ConvertostringData = Array.ConvertAll<object, string>(values, Convert.ToString);

                    double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                    string Table = Data.Data_Table.Substring(5, 1);
                    int lot = Convert.ToInt16(Table);
                    int j = 0;


                    int Index_For = this.Data.Clotho_Spcc_List[0].Max.Length;
                    int ForCount = values.Length - 8;
                    int Db_Limit = Data.DB_Column_Limit;


                    if (i == 0)
                    {
                        if (count == 0) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].No[lot] = Convert.ToInt16(ConvertostringData[0]);
                        else if (count == 1) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Parameter[lot] = Convert.ToString(ConvertostringData[0]);
                        else if (count == 2) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Selector[lot] = Convert.ToString(ConvertostringData[0]);
                        else if (count == 3) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Selector[lot] = Convert.ToString(ConvertostringData[0]);
                        else if (count == 4) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 5) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 6) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Spec[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 7) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Spec[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 8) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Data[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 9) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Median_Data[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 10) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Data[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 11) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].CPK[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 12) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Std[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 13) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Persent[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 14) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Fail_Count[lot] = Convert.ToInt16(ConvertostringData[0]);
                        else if (count == 15) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Outlier[lot] = Convert.ToInt64(ConvertostringData[0]);


                    }
                    else
                    {

                        if (count == 0) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].No[lot] = Convert.ToInt16(ConvertostringData[0]);
                        else if (count == 1) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Parameter[lot] = Convert.ToString(ConvertostringData[0]);
                        else if (count == 2) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Min_Selector[lot] = Convert.ToString(ConvertostringData[0]);
                        else if (count == 3) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Max_Selector[lot] = Convert.ToString(ConvertostringData[0]);
                        else if (count == 4) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Min_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 5) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Max_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 6) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Min_Spec[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 7) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Max_Spec[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 8) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Min_Data[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 9) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Median_Data[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 10) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Max_Data[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 11) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].CPK[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 12) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Std[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 13) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Persent[lot] = Convert.ToDouble(ConvertostringData[0]);
                        else if (count == 14) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Fail_Count[lot] = Convert.ToInt16(ConvertostringData[0]);
                        else if (count == 15) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i]].Outlier[lot] = Convert.ToInt64(ConvertostringData[0]);

                    }

                    for (j = 1; j < ForCount; j++)
                    {
                        if (count == 0) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].No[lot] = Convert.ToInt16(ConvertostringData[j]);
                        else if (count == 1) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Parameter[lot] = Convert.ToString(ConvertostringData[j]);
                        else if (count == 2) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Selector[lot] = Convert.ToString(ConvertostringData[j]);
                        else if (count == 3) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Selector[lot] = Convert.ToString(ConvertostringData[j]);
                        else if (count == 4) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 5) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 6) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Spec[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 7) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Spec[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 8) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Data[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 9) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Median_Data[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 10) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Data[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 11) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].CPK[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 12) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Std[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 13) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Persent[lot] = Convert.ToDouble(ConvertostringData[j]);
                        else if (count == 14) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Fail_Count[lot] = Convert.ToInt16(ConvertostringData[j]);
                        else if (count == 15) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Outlier[lot] = Convert.ToInt64(ConvertostringData[j]);

                        //      if (this.Data.Data_Table == "B_Spec_Min") this.Data.Customor_Clotho_List[Db_Limit * i + j].Min[count] = doubles[j];
                        //       else if (this.Data.Data_Table == "B_Spec_Max") this.Data.Customor_Clotho_List[Db_Limit * i + j].Max[count] = doubles[j];

                    }

                    if (count == 0) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].No[lot] = Convert.ToInt16(ConvertostringData[j]);
                    else if (count == 1) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Parameter[lot] = Convert.ToString(ConvertostringData[j]);
                    else if (count == 2) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Selector[lot] = Convert.ToString(ConvertostringData[j]);
                    else if (count == 3) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Selector[lot] = Convert.ToString(ConvertostringData[j]);
                    else if (count == 4) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 5) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Spec_Control[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 6) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Spec[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 7) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Spec[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 8) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Min_Data[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 9) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Median_Data[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 10) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Max_Data[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 11) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].CPK[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 12) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Std[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 13) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Persent[lot] = Convert.ToDouble(ConvertostringData[j]);
                    else if (count == 14) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Fail_Count[lot] = Convert.ToInt16(ConvertostringData[j]);
                    else if (count == 15) For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Db_Limit * i + j]].Outlier[lot] = Convert.ToInt64(ConvertostringData[j]);

                    //    if (this.Data.Data_Table == "B_Spec_Min") this.Data.Customor_Clotho_List[Db_Limit * i + j].Min[count] = doubles[j];
                    //    else if (this.Data.Data_Table == "B_Spec_Max") this.Data.Customor_Clotho_List[Db_Limit * i + j].Max[count] = doubles[j];

                    count++;

                    double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;

                }
                SqReader[i].Close();

                stringA[i].Clear();
                cmd[i].CommandText = "";

                ThreadFlags[i].Set();


            }
            public void Insert_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    Insert_Spec_Data_Thread(i);
                  //  ThreadFlags[i] = new ManualResetEvent(false);
                  //  ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                  //  Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            //public void Insert_Spec_Data_Thread(Object threadContext)
            //{
            //    int i = (int)threadContext;
            //    int Count = Data.Per_DB_Column_Count[i] - 1;
            //    TestTime1[i].Restart();
            //    TestTime1[i].Start();

            //    int lndex = 0;

            //    if (Data.Data_Table == "B_Spec_Min") lndex = Data.Customor_Clotho_List[0].Min.Length;
            //    else if (Data.Data_Table == "B_Spec_Max") lndex = Data.Customor_Clotho_List[0].Min.Length;
            //    else if (Data.Data_Table == "C_Spec_Min") lndex = 1;
            //    else if (Data.Data_Table == "C_Spec_Max") lndex = 1;


            //    for (int j = 0; j < lndex; j++)
            //    {

            //        int k = 0;
            //        stringA[i].Clear();
            //        cmd[i].Reset();
            //        if (i == 0)
            //        {
            //            if (Data.Data_Table == "B_Spec_Min") stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[j] + "',");
            //            else if (Data.Data_Table == "B_Spec_Max") stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[j] + "',");
            //            else if (Data.Data_Table == "C_Spec_Min") stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[j] + "',");
            //            else if (Data.Data_Table == "C_Spec_Max") stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[j] + "',");


            //        }
            //        else
            //        {
            //            if (Data.Data_Table == "B_Spec_Min") stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count].Min[j] + "',");
            //            else if (Data.Data_Table == "B_Spec_Max") stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count].Max[j] + "',");
            //            else if (Data.Data_Table == "C_Spec_Min")
            //            {
            //                if (j == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count].Min[0] + "',");
            //            }
            //            else if (Data.Data_Table == "C_Spec_Max")
            //            {
            //                if (j == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count].Max[0] + "',");
            //            }


            //        }

            //        for (k = 1; k < Count; k++)
            //        {
            //            if (Data.Data_Table == "B_Spec_Min") stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[j] + "',");
            //            else if (Data.Data_Table == "B_Spec_Max") stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[j] + "',");
            //            else if (Data.Data_Table == "C_Spec_Min")
            //            {
            //                if (j == 0) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[0] + "',");
            //            }
            //            else if (Data.Data_Table == "C_Spec_Max")
            //            {
            //                if (j == 0) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[0] + "',");
            //            }


            //        }

            //        if (Data.Data_Table == "B_Spec_Min") stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[j] + "', '" + j + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
            //        else if (Data.Data_Table == "B_Spec_Max") stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[j] + "', '" + j + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
            //        else if (Data.Data_Table == "C_Spec_Min")
            //        {
            //            if (j == 0) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[0] + "', '" + j + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
            //        }
            //        else if (Data.Data_Table == "C_Spec_Max")
            //        {
            //            if (j == 0) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[0] + "', '" + j + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
            //        }



            //        cmd[i].CommandText = stringA[i].ToString();

            //        cmd[i].ExecuteNonQuery();
            //    }

            //    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

            //    stringA[i].Clear();
            //    ThreadFlags[i].Set();
            //}
            public void Make_table(string Tablename)
            {

            }
            public void Make_table2(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {

            }
            public void Make_table_For_Filename(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {
                Data = Data_Edit;
                Table = Tablename;

                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(_Make_Table_For_Filename), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void _Make_Table_For_Filename(Object threadContext)
            {
                int i = (int)threadContext;

                stringA[i].Append("CREATE TABLE " + Table + "(FIle VARCAHR(20))");


                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].CommandText = "";

                stringA[i].Append(",");

                ThreadFlags[i].Set();
            }

            public void Make_table_For_Trace(string Tablename, string Chan, bool Flag)
            {
                stringA[0].Clear();
                stringA[0].Append("CREATE TABLE " + Tablename + "( FIRST VARCAHR(5), END VARCAHR(5), DBCOUNT VARCHAR(5), COLUMNCOUNT VARCHAR(5) );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";

                stringA[0].Clear();
                stringA[0].Append("INSERT INTO INF VALUES ('" + TheFirst_Trashes_Header_Count + "' , '" + TheEnd_Trashes_Header_Count + "' , '" + Data.Per_DB_Column_Count.Length + "' , '" + Data.Per_DB_Column_Count[Data.Per_DB_Column_Count.Length - 1] + "' );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";
            }
            public void Delete_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Delete_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Delete_Lot_Data(string Query)
            {

            }
            public void Delete_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;



                stringA[i].Append("Delete from " + Table + " where id = 1");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                stringA[i].Clear();

                stringA[i].Append("Delete from " + Table + " where id = 2");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Save_table(Data_Class.Data_Editing.INT Data_Edit)
            {

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //  Insert_table_Data_Thread(i);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_table_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }

            public void Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Insert_Customer_Spec_table_Data_Thread(i);
                  //  ThreadFlags[i] = new ManualResetEvent(false);
                //   ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Customer_Spec_table_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                  //  Wait[i] = ThreadFlags[i].WaitOne();
                }


            }

            public void Insert_table_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();

                string Table = Data.Data_Table.Substring(5, 1);
                int j = Convert.ToInt16(Table);

                int Index_for_Table_item = 16;

                for (int ii = 0; ii < Index_for_Table_item; ii++)
                {
                    //  for (int j = 0; j < lndex; j++)
                    //   {

                    int k = 0;

                    stringA[i].Clear();
                    cmd[i].Reset();

                    if (i == 0)
                    {
                        if (ii == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].No[j] + "',");
                        else if (ii == 1) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Parameter[j] + "',");
                        else if (ii == 2) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Selector[j] + "',");
                        else if (ii == 3) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Selector[j] + "',");
                        else if (ii == 4) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Spec_Control[j] + "',");
                        else if (ii == 5) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Spec_Control[j] + "',");
                        else if (ii == 6) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Spec[j] + "',");
                        else if (ii == 7) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Spec[j] + "',");
                        else if (ii == 8) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Min_Data[j] + "',");
                        else if (ii == 9) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Median_Data[j] + "',");
                        else if (ii == 10) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Max_Data[j] + "',");
                        else if (ii == 11) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].CPK[j] + "',");
                        else if (ii == 12) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Std[j] + "',");
                        else if (ii == 13) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Persent[j] + "',");
                        else if (ii == 14) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Fail_Count[j] + "',");
                        else if (ii == 15) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].Outlier[j] + "',");


                    }
                    else
                    {
                        if (ii == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].No[j] + "',");
                        else if (ii == 1) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Parameter[j] + "',");
                        else if (ii == 2) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Min_Selector[j] + "',");
                        else if (ii == 3) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Max_Selector[j] + "',");
                        else if (ii == 4) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Min_Spec_Control[j] + "',");
                        else if (ii == 5) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Max_Spec_Control[j] + "',");
                        else if (ii == 6) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Min_Spec[j] + "',");
                        else if (ii == 7) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Max_Spec[j] + "',");
                        else if (ii == 8) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Min_Data[j] + "',");
                        else if (ii == 9) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Median_Data[j] + "',");
                        else if (ii == 10) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Max_Data[j] + "',");
                        else if (ii == 11) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].CPK[j] + "',");
                        else if (ii == 12) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Std[j] + "',");
                        else if (ii == 13) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Persent[j] + "',");
                        else if (ii == 14) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Fail_Count[j] + "',");
                        else if (ii == 15) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * i]].Outlier[j] + "',");


                    }

                    for (k = 1; k < Count; k++)
                    {

                        if (ii == 0) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].No[j] + "',");
                        else if (ii == 1) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Parameter[j] + "',");
                        else if (ii == 2) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Selector[j] + "',");
                        else if (ii == 3) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Selector[j] + "',");
                        else if (ii == 4) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Spec_Control[j] + "',");
                        else if (ii == 5) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Spec_Control[j] + "',");
                        else if (ii == 6) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Spec[j] + "',");
                        else if (ii == 7) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Spec[j] + "',");
                        else if (ii == 8) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Data[j] + "',");
                        else if (ii == 9) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Median_Data[j] + "',");
                        else if (ii == 10) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Data[j] + "',");
                        else if (ii == 11) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].CPK[j] + "',");
                        else if (ii == 12) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Std[j] + "',");
                        else if (ii == 13) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Persent[j] + "',");
                        else if (ii == 14) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Fail_Count[j] + "',");
                        else if (ii == 15) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Outlier[j] + "',");


                    }

                    if (ii == 0) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].No[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 1) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Parameter[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 2) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Selector[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 3) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Selector[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 4) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Spec_Control[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 5) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Spec_Control[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 6) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Spec[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 7) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Spec[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 8) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Min_Data[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 9) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Median_Data[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 10) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Max_Data[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 11) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].CPK[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 12) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Std[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 13) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Persent[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 14) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Fail_Count[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");
                    else if (ii == 15) stringA[i].Append("'" + For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]].Outlier[j] + "', '" + ii + "' , '0' , '0' , '0' , '0' , '0' , '0' );");


                    cmd[i].CommandText = stringA[i].ToString();

                    cmd[i].ExecuteNonQuery();
                    //   }
                }
                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void Insert_Customer_Spec_table_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();

                string Table = Data.Data_Table.Substring(5, 1);
                //   int j = Convert.ToInt16(Table);

                cmd[i] = new SQLiteCommand(conn[i]);

              
                    //  for (int j = 0; j < lndex; j++)
                    //   {

                    int k = 0;

                    stringA[i].Clear();
                    cmd[i].Reset();

                if (Data.Data_Table.ToUpper() == "CLOTHO_SPEC")
                {
                    for (int ii = 0; ii < 2; ii++)
                    {
                        for (int j = 0; j < Data.Clotho_Spcc_List[0].Min.Length; j++)
                        {
                            stringA[i].Clear();

                            if (i == 0)
                            {
                                if (ii == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Min[j] + "',");
                                else if (ii == 1) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Max[j] + "',");

                                for (int p = 0; p < 9; p++)
                                {
                                    stringA[i].Append("'" + p + "' ,");
                                }

                            }
                            else
                            {
                                if (ii == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i].Min[j] + "',");
                                else if (ii == 1) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i].Max[j] + "',");


                            }

                            for (k = 1; k < Count; k++)
                            {
                                if (ii == 0) stringA[i].Append("'" + Data.Clotho_Spcc_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[j] + "',");
                                else if (ii == 1) stringA[i].Append("'" + Data.Clotho_Spcc_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[j] + "',");

                            }

                            if (ii == 0) stringA[i].Append("'" + Data.Clotho_Spcc_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[j] + "',  '0', '" + j + "' , '0' , '0' , '0' , '0' );");
                            else if (ii == 1) stringA[i].Append("'" + Data.Clotho_Spcc_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[j] + "', '0','" + Data.Clotho_Spcc_List[0].Min.Length + j + "'  , '0' , '0' , '0' , '0' );");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                        }


                    }



                }
                else
                {
                    for (int ii = 0; ii < 2; ii++)
                    {
                        for (int j = 0; j < Data.Clotho_Spcc_List[0].Min.Length; j++)
                        {
                            stringA[i].Clear();
                            if (i == 0)
                            {
                                if (ii == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[j] + "',");
                                else if (ii == 1) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[j] + "',");



                            }
                            else
                            {
                                if (ii == 0) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i].Min[j] + "',");
                                else if (ii == 1) stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i].Max[j] + "',");


                            }

                            for (k = 1; k < Count; k++)
                            {
                                if (ii == 0) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[j] + "',");
                                else if (ii == 1) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[j] + "',");

                            }

                            if (ii == 0) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Min[j] + "',  '0', '" + j + "' , '0' , '0' , '0' , '0' );");
                            else if (ii == 1) stringA[i].Append("'" + Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k].Max[j] + "', '0','" + Data.Clotho_Spcc_List[0].Min.Length + j + "'  , '0' , '0' , '0' , '0' );");



                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();

                        }
                    }
                }

                cmd[i].Dispose();
                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
             //   ThreadFlags[i].Set();
            }
            public void Insert_Files_Name(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Road_Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    Road_Save_Customer_Spec_table_Thread(i);
                  //  ThreadFlags[i] = new ManualResetEvent(false);
                  // ThreadPool.QueueUserWorkItem(new WaitCallback(Road_Save_Customer_Spec_table_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                 //   Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Road_Save_Customer_Spec_table_Thread(Object threadContext)
            {
                int i = (int)threadContext;


                int count = 0;


                cmd[i] = new SQLiteCommand(conn[i]);
                sqlAdapter[i] = new SQLiteDataAdapter();
                stringA[i].Clear();
                stringA[i].Append("Select * from " + this.Data.Data_Table);

                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                count = 0;
                int length = Data.Clotho_Spcc_List[0].Min.Length;

                while (SqReader[i].Read())
                {

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);


                    double[] ConvertoDoubleData = Array.ConvertAll<object, double>(values, Convert.ToDouble);

                    double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                    string Table = Data.Data_Table.Substring(5, 1);

                    int j = 0;

                    if(Data.Data_Table.ToUpper() == "CLOTHO_SPEC")
                    {
                        int Index_For = Data.Clotho_Spcc_List[0].Max.Length;
                        int ForCount = values.Length - 6;
                        int Db_Limit = Data.DB_Column_Limit;

                     
                        if (i == 0)
                        {
                            for (int index = 0; index < Index_For; index++)
                            {
                                if (count < Data.Clotho_Spcc_List[0].Min.Length) Data.Clotho_Spcc_List[0].Min[index] = ConvertoDoubleData[0];
                                else  Data.Clotho_Spcc_List[0].Max[index] = ConvertoDoubleData[0];
                            }


                        }
                        else
                        {
                            for (int index = 0; index < Index_For; index++)
                            {
                               if(count < Data.Clotho_Spcc_List[0].Min.Length) Data.Clotho_Spcc_List[Db_Limit * i - 9].Min[index] = ConvertoDoubleData[0];
                                else  Data.Clotho_Spcc_List[Db_Limit * i - 9].Max[index] = ConvertoDoubleData[0];

                                //  if (count == 0) Data.Clotho_List[Db_Limit * i - 9].Min[index] = 0;
                                // else if (count == 1) Data.Clotho_List[Db_Limit * i - 9].Max[index] = 0;
                            }


                        }


                        if (i == 0)
                        {
                            for (j = 10; j < ForCount; j++)
                            {
                                for (int index = 0; index < Index_For; index++)
                                {
                                    if(count < Data.Clotho_Spcc_List[0].Min.Length) Data.Clotho_Spcc_List[Db_Limit * i + j - 9].Min[index] = ConvertoDoubleData[j];
                                    else  Data.Clotho_Spcc_List[Db_Limit * i + j - 9].Max[index] = ConvertoDoubleData[j];

                                }



                            }
                        }
                        else
                        {
                            for (j = 1; j < ForCount; j++)
                            {
                                for (int index = 0; index < Index_For; index++)
                                {
                                    if(count < Data.Clotho_Spcc_List[0].Min.Length) Data.Clotho_Spcc_List[Db_Limit * i + j - 9].Min[index] = ConvertoDoubleData[j];
                                    else  Data.Clotho_Spcc_List[Db_Limit * i + j - 9].Max[index] = ConvertoDoubleData[j];
                                }

                            }
                        }


                        count++;

                        double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                    }
                    else
                    {
                        int Index_For = Data.Customor_Clotho_List[0].Max.Length;
                        int ForCount = values.Length - 6;
                        int Db_Limit = Data.DB_Column_Limit;


                        if (i == 0)
                        {
                            for (int index = 0; index < Index_For; index++)
                            {
                                if (count < Data.Clotho_Spcc_List[0].Min.Length) Data.Customor_Clotho_List[0].Min[index] = ConvertoDoubleData[0];
                                else  Data.Customor_Clotho_List[0].Max[index] = ConvertoDoubleData[0];
                            }


                        }
                        else
                        {
                            for (int index = 0; index < Index_For; index++)
                            {
                                if (count < Data.Clotho_Spcc_List[0].Min.Length) Data.Customor_Clotho_List[Db_Limit * i - 9].Min[index] = ConvertoDoubleData[0];
                                else  Data.Customor_Clotho_List[Db_Limit * i - 9].Max[index] = ConvertoDoubleData[0];

                                //  if (count == 0) Data.Clotho_List[Db_Limit * i - 9].Min[index] = 0;
                                // else if (count == 1) Data.Clotho_List[Db_Limit * i - 9].Max[index] = 0;
                            }


                        }


                        if (i == 0)
                        {
                            for (j = 10; j < ForCount; j++)
                            {
                                for (int index = 0; index < Index_For; index++)
                                {
                                    if (count < Data.Clotho_Spcc_List[0].Min.Length) Data.Customor_Clotho_List[Db_Limit * i + j - 9].Min[index] = ConvertoDoubleData[j];
                                    else  Data.Customor_Clotho_List[Db_Limit * i + j - 9].Max[index] = ConvertoDoubleData[j];

                                }



                            }
                        }
                        else
                        {
                            for (j = 1; j < ForCount; j++)
                            {
                                for (int index = 0; index < Index_For; index++)
                                {
                                    if (count < Data.Clotho_Spcc_List[0].Min.Length) Data.Customor_Clotho_List[Db_Limit * i + j - 9].Min[index] = ConvertoDoubleData[j];
                                    else  Data.Customor_Clotho_List[Db_Limit * i + j - 9].Max[index] = ConvertoDoubleData[j];
                                }

                            }
                        }


                     

                        double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                    }

                    count++;

                }
                SqReader[i].Close();

                stringA[i].Clear();
                cmd[i].Dispose();

          //      ThreadFlags[i].Set();
            }

            public void LOTID_Update(string Query, string Query2, string CellID)
            {

            }
            public void Gross_Update_Data(object data)
            {
                Update_Data_ID = data;

                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }

            }
            public void Gross_Update_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (object o in (Array)Update_Data_ID)
                {
                    cmd[i].CommandText = "Update data set FAIL = '1'  where id = " + o.ToString();
                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Gross_Update_Datas(List<string> data)
            {
                Update_Datas_ID = data.ToArray();
                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                      //  Gross_Update_Datas_Thread(i);
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Datas_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }
            }
            public void Gross_Update_Datas_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    cmd[i] = new SQLiteCommand(conn[i]);

                    Dictionary<string, List<string>> tests = key.Value;

                    int count = 0;
                    int j = 0;
                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {
                        foreach (object o in (Array)Update_Datas_ID)
                        {

                            if (j == 0)
                            {
                 

                                stringA[i].Clear();
                                stringA[i].Append("Update " + key.Key + " set FAIL = '1'  where id = '" + o.ToString() + "'");

                            }
                            else
                            {
                                stringA[i].Append(" or id = '" + o.ToString() + "'");
                            }
                            j++;
                        }

                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();

                        cmd[i].Dispose();
                    }

                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Chnaged_Spec_Update_Data(int DB, int Index, string Parameter, double Spec, int GetId)
            {
                stringA[DB].Clear();
                stringA[DB].Append("Update newspec set " + Parameter + " = " + Spec + " where id = " + GetId);

                cmd[DB].CommandText = stringA[DB].ToString();

                cmd[DB].ExecuteNonQuery();
                cmd[DB].Reset();

                stringA[DB].Clear();
            }
            public Dictionary<string, double[]> Chnaged_Spec_Anl_Yield(int DB, int Index, string Parameter)
            {
                Dictionary<string, double[]> Dic_Change_Spec = new Dictionary<string, double[]>();

        
                object[] GetData = new object[0];
                object[] GetData_Ref = new object[0];

         
                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
           
                    Dictionary<string, List<string>> tests = key.Value;

                    int count = 0;
                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {
                        stringA[DB].Clear();

                        // conn[DB] = new SQLiteConnection(strConn[DB]);
                        cmd[DB] = new SQLiteCommand(conn[DB]);
                        sqlAdapter[DB] = new SQLiteDataAdapter();
                        // conn[DB].Open();
                        // cmd[DB] = new SQLiteCommand(conn[DB]);
                        stringA[DB].Append("Select " + Parameter + " from " + key.Key + " where Fail not like '1'");

                        cmd[DB].CommandText = stringA[DB].ToString();
                        ds[DB] = new DataSet();

                        sqlAdapter[DB].SelectCommand = cmd[DB];
                        sqlAdapter[DB].Fill(ds[DB]);

                        GetData = new object[ds[DB].Tables[0].Rows.Count];
                 
                        foreach (DataRow dr in ds[DB].Tables[0].Rows)
                        {
                            GetData[count] = dr.ItemArray[0];
                            count++;
                        }

                        GetData_Ref = GetData_Ref.Concat(GetData).ToArray();
                    }
                }


                //for (int loop = 0; loop < Table_Count; loop++)
                //{

                //    stringA[DB].Clear();

                //    // conn[DB] = new SQLiteConnection(strConn[DB]);
                //    // cmd[DB] = new SQLiteCommand(conn[DB]);
                //    // sqlAdapter[DB] = new SQLiteDataAdapter();
                //    // conn[DB].Open();
                //    // cmd[DB] = new SQLiteCommand(conn[DB]);
                //    stringA[DB].Append("Select " + Parameter + " from data" + loop + " where Fail not like '1'");

                //    cmd[DB].CommandText = stringA[DB].ToString();
                //    ds[DB] = new DataSet();

                //    sqlAdapter[DB].SelectCommand = cmd[DB];
                //    sqlAdapter[DB].Fill(ds[DB]);

                //    if (loop == 0) GetData = new object[ds[DB].Tables[0].Rows.Count * Table_Count];
                //    foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //    {
                //        GetData[count] = dr.ItemArray[0];
                //        count++;
                //    }


                //    // cmd[DB].Dispose();
                //}
                //Array.Resize(ref GetData, count);
                double[] Toduble_Data = Array.ConvertAll<object, double>(GetData_Ref, Convert.ToDouble);

                Dic_Change_Spec.Add("DATA", Toduble_Data);

                sqlAdapter[DB].Dispose();
                cmd[DB].Dispose();

                stringA[DB].Clear();

                return Dic_Change_Spec;
            }
            public void Get_Ave_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                DB_DataSet_Values = new List<double[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    DB_DataSet_Values[i] = new List<double[]>();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_Thread), i);
                  //  Get_Ave_Data_Thread(i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();

                    //      conn[i] = new SQLiteConnection(strConn[i]);
                    //     cmd[i] = new SQLiteCommand(conn[i]);
                    //      conn[i].Open();

                }
                double testtime1 = TestTime1.Elapsed.TotalMilliseconds;
            }
            public void Get_Ave_Data_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  Get_Ave_Data_For_New_Spec_Thread(i);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_For_New_Spec_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();

                    //   conn[i] = new SQLiteConnection(strConn[i]);
                    //  cmd[i] = new SQLiteCommand(conn[i]);
                    //   conn[i].Open();

                }
            }
            public void Set_Refer_for_Anlyzer(Data_Class.Data_Editing.INT Data_Edit)
            {
                stringA[0].Clear();
                stringA[0].Append("Select id from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                ForCampare_Yield_List1 = new List<List<int>[]>();
                for (int k = 0; k < Value.Length; k++)
                {
                    ForCampare_Yield_List = new List<int>[Data.DB_Count];

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ForCampare_Yield_List[i] = new List<int>();
                    }

                    for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                    {
                        for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                        {
                            ForCampare_Yield_List[i].Add(0);
                        }
                    }

                    ForCampare_Yield_List1.Add(ForCampare_Yield_List);
                }
            }
            public void Get_Ave_Data2(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    Get_Ave_Data_Thread2(i);
                }
            }
            public void Get_Ave_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                int count = 0;
                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    Dictionary<string, List<string>> tests = key.Value;


                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {
            

                   
                           // conn[i] = new SQLiteConnection(strConn[i]);
                            cmd[i] = new SQLiteCommand(conn[i]);
                         //   conn[i].Open();
                      
          

                        stringA[i].Clear();
                        stringA[i].Append("Select * from " + key.Key + " where Fail = '0'");

                

                        cmd[i].CommandText = stringA[i].ToString();
                        SqReader[i] = cmd[i].ExecuteReader();

                        object[] Std_Value = new object[SqReader[i].FieldCount];
                        double[] Std_Value_Convert = new double[SqReader[i].FieldCount];

                        while (SqReader[i].Read())
                        {
                            SqReader[i].GetValues(Std_Value);

                            if (i == Data.DB_Count - 1)
                            {
                                // ForCount = values.Length - 12;
                                Std_Value[Std_Value.Length - 11] = 0;
                                Std_Value[Std_Value.Length - 10] = 0;
                                Std_Value[Std_Value.Length - 9] = 0;
                                Std_Value[Std_Value.Length - 8] = 0;
                                Std_Value[Std_Value.Length - 7] = 0;
                                Std_Value[Std_Value.Length - 6] = 0;
                                Std_Value[Std_Value.Length - 5] = 0;
                                Std_Value[Std_Value.Length - 4] = 0;
                                Std_Value[Std_Value.Length - 3] = 0;
                                Std_Value[Std_Value.Length - 2] = 0;
                                Std_Value[Std_Value.Length - 1] = 0;
                            }
                            else if (i == 0)
                            {
                                // ForCount = values.Length - 6;

                                Std_Value[0] = 0;
                                Std_Value[5] = 0;
                                Std_Value[8] = 0;
                                Std_Value[9] = 0;

                                Std_Value[Std_Value.Length - 6] = 0;
                                Std_Value[Std_Value.Length - 5] = 0;
                                Std_Value[Std_Value.Length - 4] = 0;
                                Std_Value[Std_Value.Length - 3] = 0;
                                Std_Value[Std_Value.Length - 2] = 0;
                                Std_Value[Std_Value.Length - 1] = 0;
                            }
                            else
                            {
                                //  ForCount = values.Length - 6;

                                Std_Value[Std_Value.Length - 6] = 0;
                                Std_Value[Std_Value.Length - 5] = 0;
                                Std_Value[Std_Value.Length - 4] = 0;
                                Std_Value[Std_Value.Length - 3] = 0;
                                Std_Value[Std_Value.Length - 2] = 0;
                                Std_Value[Std_Value.Length - 1] = 0;
                            }



                            Std_Value_Convert = Array.ConvertAll<object, double>(Std_Value, Convert.ToDouble);
                            DB_DataSet_Values[i].Add(Std_Value_Convert);

                            count++;

                        }
                        SqReader[i].Close();
                        cmd[i].Dispose();
                    }



                }


                double testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                STDEVandMedian(DB_DataSet_Values[i], i, count);

                DB_DataSet_Values[i] = new List<double[]>();
                double testtime2 = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
          
              //  conn[i].Dispose();
                ThreadFlags[i].Set();
            }
            public void Get_Ave_Data_For_New_Spec_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                int count = 0;
                List<double[]> DataSet_Values = new List<double[]>();

                for (int loop = 0; loop < Table_Count; loop++)
                {
                    stringA[i].Clear();
                    stringA[i].Append("Select * from data" + loop + " where Fail not like '1'");
                    //  conn[i] = new SQLiteConnection(strConn[i]);
                    //    cmd[i] = new SQLiteCommand(conn[i]);
                    //  conn[i].Open();

                    cmd[i].CommandText = stringA[i].ToString();
                    SqReader[i] = cmd[i].ExecuteReader();


                    double testtime = TestTime1.Elapsed.TotalMilliseconds;


                    while (SqReader[i].Read())
                    {
                        object[] values = new object[SqReader[i].FieldCount];
                        SqReader[i].GetValues(values);
                        values[Data.Per_DB_Column_Count[i] + 2] = 0;
                        values[Data.Per_DB_Column_Count[i] + 3] = 0;
                        values[Data.Per_DB_Column_Count[i] + 6] = 0;
                        double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);
                        DataSet_Values.Add(doubles);

                        count++;

                    }
                    SqReader[i].Close();

                    //   cmd[i].Dispose();
                    //   conn[i].Close();

                }
                double testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                STDEVandMedian_For_New_Spec(DataSet_Values, i, count);

                double testtime2 = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();

                ThreadFlags[i].Set();
            }
            public void Get_Ave_Data_Thread2(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from data where Fail not like '1'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                List<double[]> DataSet_Values = new List<double[]>();
                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    values[Data.Per_DB_Column_Count[i] + 2] = 0;
                    values[Data.Per_DB_Column_Count[i] + 3] = 0;
                    double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);
                    DataSet_Values.Add(doubles);

                    count++;

                }
                SqReader[i].Close();

                STDEVandMedian(DataSet_Values, i, count);

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();
            }
            public void Get_Saved_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                DataSet_Value = new List<string[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    DataSet_Value[i] = new List<string[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Saved_Spec_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }
            public void Get_Saved_Spec_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from newspec");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    DataSet_Value[i].Add(stringD);

                    count++;

                }
                SqReader[i].Close();

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();

            }
            public void Get_Rows_Data(Data_Class.Data_Editing.INT Data_Edit)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                DataSet_Value = new List<string[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    DataSet_Value[i] = new List<string[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Rows_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
            }
            public void Get_Rows_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from data where id = '" + Data.Set_ID + "'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    DataSet_Value[i].Add(stringD);
                    count++;
                    break;
                }


                SqReader[i].Close();

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();

            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Interface;
                Dic_Test = new Dictionary<string, For_Box>[this.Data.DB_Count];


                stringA[0].Clear();


                SampleCount = 0;

                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    Dictionary<string, List<string>> tests = key.Value;


                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {
                        string Query = "Select count(id) from " + key.Key + "  where Fail = '0'";

                        SampleCount += Get_Sample_Count(0, Query);
                    }
                }





                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Dic_Test[i] = new Dictionary<string, For_Box>();
                    stringA[i].Clear();
                   ThreadFlags[i] = new ManualResetEvent(false);
                 //  Get_Selected_Para_Thread(i);
                   ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Selected_Para_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                   Wait[i] = ThreadFlags[i].WaitOne();
                }

                if(Dic_Test[0].Count == 0)
                {
                   

                    foreach (Dictionary<string, CSV_Class.For_Box> _T in Dic_Test)
                    {

                        foreach (KeyValuePair<string, CSV_Class.For_Box> _D in _T)
                        {
                            object[] s = Get_Selected_Para_Thread();


                            _D.Value.WAFER_ID = Array.ConvertAll<object, string>(s, Convert.ToString);
                        }

                    }


                }
            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface, DataTable dt)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Interface;
                this.dt_test = dt;

                Dic_Test = new Dictionary<string, For_Box>[this.Data.DB_Count];


                stringA[0].Clear();


                SampleCount = 0;

                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    Dictionary<string, List<string>> tests = key.Value;


                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {
                        string Query = "Select count(id) from " + key.Key + "  where Fail = '0'";

                        SampleCount += Get_Sample_Count(0, Query);
                    }
                }





                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Dic_Test[i] = new Dictionary<string, For_Box>();
                    stringA[i].Clear();
                 //    ThreadFlags[i] = new ManualResetEvent(false);
                    Get_Selected_Para_Thread_Dt(i);
                  //   ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Selected_Para_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                  //  Wait[i] = ThreadFlags[i].WaitOne();
                }

                if (Dic_Test[0].Count == 0)
                {


                    foreach (Dictionary<string, CSV_Class.For_Box> _T in Dic_Test)
                    {

                        foreach (KeyValuePair<string, CSV_Class.For_Box> _D in _T)
                        {
                            object[] s = Get_Selected_Para_Thread();


                            _D.Value.WAFER_ID = Array.ConvertAll<object, string>(s, Convert.ToString);
                        }

                    }


                }
            }
            public void Get_Selected_Para_Thread_Dt(Object threadContext)
            {

                int DB = (int)threadContext;
                bool Flag = false;

                object[] ID_Dummy = new object[0];
                object[] WAFERID_Dummy = new object[0];
                object[] LOTID_Dummy = new object[0];
                object[] SITEID_Dummy = new object[0];
                object[] Value_Dummy = new object[0];

                object[] ID_Test = new object[0];
                object[] WAFERID_Test = new object[0];
                object[] LOTID_Test = new object[0];
                object[] SITEID_Test = new object[0];
                object[] Value_Test = new object[0];

             //   Dic_Test = new Dictionary<string, For_Box>[9];

                Dictionary<int, string> T = new Dictionary<int, string>();


                List<List<object[]>> D_Data = new List<List<object[]>>();


                int indext = 0;
                int Find_DB = 0;
                for (int C = 0; C < Line.Length; C++)
                {
                    DataRow[] Resultsrow = dt_test.Select("[Parameter] = '" + Line[C] + "'");

                    indext = Convert.ToInt16(Resultsrow[0].ItemArray[0]);

                    if (indext < 11)
                    {
                        if(DB == 0)
                        {
                            T.Add(indext, Line[C]);
                        }
                  

                    }
                    else if (Data.New_Header.Length < indext - 9)
                    {
                        if (DB == Data.DB_Count - 1)
                            T.Add(indext, Line[C]);

                    }
                    else
                    {
                        for (int i = 0; i < Data.Reference_Header.Length; i++)
                        {
                            if (Line[C] == Data.Reference_Header[i])
                            {
                                Find_DB = 0;
                                if (i > this.Data.DB_Column_Limit * DB - 10)
                                {

                                    if (i <= this.Data.Per_DB_Column_Count_End[DB] - 9)
                                    {
                                        Find_DB = DB;
                                        if (Find_DB == DB)
                                            T.Add(i + 10, this.Data.New_Header[i]);
                                        Flag = true;
                                        break;
                                    }

                                }
                                else
                                {
                                    Find_DB = DB;
                                    if (Find_DB == 0)
                                        T.Add(i + 10, this.Data.New_Header[i]);
                                    Flag = true;
                                    break;
                                }
                            }
                        }
                    }

                }


                for (int d = 0; d < T.Count + 4; d++)
                {
                    object[] s = new object[SampleCount];
                    List<object[]> L1 = new List<object[]>();
                    L1.Add(s);

                    D_Data.Add(L1);
                }

                if (T.Count != 0)
                {
                    int row = 0;

                    foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                    {
                        int count = 0;
                        Dictionary<string, List<string>> tests = key.Value;
                        List<List<object[]>> D = new List<List<object[]>>();

                        foreach (KeyValuePair<string, List<string>> ts in tests)
                        {

                            stringA[DB].Clear();
                            if (DB == 0)
                            {


                                stringA[DB].Append("Select id, WAFER_ID, LOTID, SITEID, ");
                                int k = 0;
                                foreach (KeyValuePair<int, string> H in T)
                                {

                                    if (k != T.Count - 1)
                                    {
                                        stringA[DB].Append(H.Value + ",");

                                    }
                                    else
                                    {
                                        stringA[DB].Append(H.Value);

                                    }
                                    k++;
                                }

                                stringA[DB].Append(" from " + key.Key + " where Fail not like '1'");


                                cmd[DB] = new SQLiteCommand(conn[DB]);
                                sqlAdapter[DB] = new SQLiteDataAdapter();


                                cmd[DB].CommandText = stringA[DB].ToString();
                                ds[DB] = new DataSet();

                                sqlAdapter[DB].SelectCommand = cmd[DB];
                                sqlAdapter[DB].Fill(ds[DB]);


                                k = 0;

                                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                                {
                                    int r = 0;

                                    for (int n = 0; n < dr.ItemArray.Length; n++)
                                    {
                                        D_Data[r][0][row] = dr.ItemArray[n];
                                        r++;
                                    }

                                    row++;
                                    r++;
                                }

                            }
                            else
                            {

                                count = 0;

                                cmd[DB] = new SQLiteCommand(conn[DB]);
                                stringA[0].Clear();
                                stringA[DB].Append("Select id, FAIL , LOTID, SITEID, ");
                                int k = 0;
                                foreach (KeyValuePair<int, string> H in T)
                                {

                                    if (k != T.Count - 1)
                                    {
                                        stringA[DB].Append(H.Value + ",");

                                    }
                                    else
                                    {
                                        stringA[DB].Append(H.Value);

                                    }
                                    k++;
                                }

                                stringA[DB].Append(" from " + key.Key + " where Fail not like '1'");

                                cmd[DB] = new SQLiteCommand(conn[DB]);
                                sqlAdapter[DB] = new SQLiteDataAdapter();


                                cmd[DB].CommandText = stringA[DB].ToString();
                                ds[DB] = new DataSet();

                                sqlAdapter[DB].SelectCommand = cmd[DB];
                                sqlAdapter[DB].Fill(ds[DB]);

                                k = 0;

                                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                                {
                                    int r = 0;

                                    for (int n = 0; n < dr.ItemArray.Length; n++)
                                    {
                                        D_Data[r][0][row] = dr.ItemArray[n];
                                        r++;
                                    }

                                    row++;
                                    r++;
                                }



                            }

                        }

                    }

                    string[] ID = new string[0];
                    string[] WAFER_ID = new string[0];
                    string[] LOT_ID = new string[0];
                    string[] SITE_ID = new string[0];
                    int m = 0;

                    m = 4;

                    foreach (KeyValuePair<int, string> _L in T)
                    {
                        if (_L.Key < 11 || Data.New_Header.Length < _L.Key - 9)
                        {

                            ID = Array.ConvertAll<object, string>(D_Data[0][0], Convert.ToString);
                            WAFER_ID = Array.ConvertAll<object, string>(D_Data[1][0], Convert.ToString);
                            LOT_ID = Array.ConvertAll<object, string>(D_Data[2][0], Convert.ToString);
                            SITE_ID = Array.ConvertAll<object, string>(D_Data[3][0], Convert.ToString);

                        

                            string[] Data = Array.ConvertAll<object, string>(D_Data[m][0], Convert.ToString);

                            int i = _L.Key;

                            CSV_Class.For_Box Box = new CSV_Class.For_Box(_L.Value, Data, ID, WAFER_ID, SITE_ID, LOT_ID, 0, 0, "", "", "", "", "", "", "");

                            Dic_Test[DB].Add(_L.Value, Box);



                            m++;

                        }
                        else
                        {
                            ID = Array.ConvertAll<object, string>(D_Data[0][0], Convert.ToString);
                            WAFER_ID = Array.ConvertAll<object, string>(D_Data[1][0], Convert.ToString);

                            LOT_ID = Array.ConvertAll<object, string>(D_Data[2][0], Convert.ToString);
                            SITE_ID = Array.ConvertAll<object, string>(D_Data[3][0], Convert.ToString);



                            double[] Data = Array.ConvertAll<object, double>(D_Data[m][0], Convert.ToDouble);

                            int i = _L.Key;

                            string NPI_Spec_Min = Convert.ToString(this.Data.Clotho_Spcc_List[_L.Key - 10].Min[0]);
                            string NPI_Spec_Max = Convert.ToString(this.Data.Clotho_Spcc_List[_L.Key - 10].Max[0]);

                            string Customer_Spec_Min = Convert.ToString(this.Data.Customor_Clotho_List[_L.Key - 10].Min[0]);
                            string Customer_Spec_Max = Convert.ToString(this.Data.Customor_Clotho_List[_L.Key - 10].Max[0]);

                            CSV_Class.For_Box Box = new CSV_Class.For_Box(_L.Value, Data, ID, WAFER_ID, SITE_ID, LOT_ID, 0, 0, "", "", "", Customer_Spec_Min, Customer_Spec_Max, NPI_Spec_Min, NPI_Spec_Max);

                            Dic_Test[DB].Add(_L.Value, Box);
                            m++;


                        }
                    }


         

                    sqlAdapter[DB].Dispose();

                    stringA[DB].Clear();
                }




     //           ThreadFlags[DB].Set();
            }

            public void Get_Selected_Para_Thread(Object threadContext)
            {

                int DB = (int)threadContext;
                bool Flag = false;

                object[] ID_Dummy = new object[0];
                object[] WAFERID_Dummy = new object[0];
                object[] LOTID_Dummy = new object[0];
                object[] SITEID_Dummy = new object[0];
                object[] Value_Dummy = new object[0];

                object[] ID_Test = new object[0];
                object[] WAFERID_Test = new object[0];
                object[] LOTID_Test = new object[0];
                object[] SITEID_Test = new object[0];
                object[] Value_Test = new object[0];

                Dictionary<int, string> T = new Dictionary<int, string>();


                List<List<object[]>> D_Data = new List<List<object[]>>();



                int Find_DB = 0;
                for (int C = 0; C < Line.Length; C++)
                {
                    for (int i = 0; i < Data.Reference_Header.Length; i++)
                    {
                        if (Line[C] == Data.Reference_Header[i])
                        {
                            Find_DB = 0;
                            if (i > this.Data.DB_Column_Limit * DB - 10)
                            {

                                if (i <= this.Data.Per_DB_Column_Count_End[DB] - 9)
                                {
                                    Find_DB = DB;
                                    if (Find_DB == DB)
                                        T.Add(i, this.Data.New_Header[i]);
                                    Flag = true;
                                    break;
                                }

                            }
                            else
                            {
                                Find_DB = DB;
                                if (Find_DB == 0)
                                    T.Add(i, this.Data.New_Header[i]);
                                Flag = true;
                                break;
                            }
                        }
                    }
                }

                for (int d = 0; d < T.Count + 4; d++)
                {
                    object[] s = new object[SampleCount];
                    List<object[]> L1 = new List<object[]>();
                    L1.Add(s);

                    D_Data.Add(L1);
                }

                if (T.Count != 0)
                {
                    int row = 0;

                    foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                    {
                        int count = 0;
                        Dictionary<string, List<string>> tests = key.Value;
                        List<List<object[]>> D = new List<List<object[]>>();

                        foreach (KeyValuePair<string, List<string>> ts in tests)
                        {

                            stringA[DB].Clear();
                            if (DB == 0)
                            {


                                stringA[DB].Append("Select id, WAFER_ID, LOTID, SITEID, ");
                                int k = 0;
                                foreach (KeyValuePair<int, string> H in T)
                                {

                                    if (k != T.Count - 1)
                                    {
                                        stringA[DB].Append(H.Value + ",");

                                    }
                                    else
                                    {
                                        stringA[DB].Append(H.Value);

                                    }
                                    k++;
                                }

                                stringA[DB].Append(" from " + key.Key + " where Fail not like '1'");


                                cmd[DB] = new SQLiteCommand(conn[DB]);
                                sqlAdapter[DB] = new SQLiteDataAdapter();


                                cmd[DB].CommandText = stringA[DB].ToString();
                                ds[DB] = new DataSet();

                                sqlAdapter[DB].SelectCommand = cmd[DB];
                                sqlAdapter[DB].Fill(ds[DB]);


                                k = 0;

                                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                                {
                                    int r = 0;

                                    for (int n = 0; n < dr.ItemArray.Length; n++)
                                    {
                                        D_Data[r][0][row] = dr.ItemArray[n];
                                        r++;
                                    }

                                    row++;
                                    r++;
                                }

                            }
                            else
                            {

                                count = 0;

                                cmd[DB] = new SQLiteCommand(conn[DB]);
                                stringA[0].Clear();
                                stringA[DB].Append("Select id, FAIL , LOTID, SITEID, ");
                                int k = 0;
                                foreach (KeyValuePair<int, string> H in T)
                                {

                                    if (k != T.Count - 1)
                                    {
                                        stringA[DB].Append(H.Value + ",");

                                    }
                                    else
                                    {
                                        stringA[DB].Append(H.Value);

                                    }
                                    k++;
                                }

                                stringA[DB].Append(" from " + key.Key + " where Fail not like '1'");

                                cmd[DB] = new SQLiteCommand(conn[DB]);
                                sqlAdapter[DB] = new SQLiteDataAdapter();


                                cmd[DB].CommandText = stringA[DB].ToString();
                                ds[DB] = new DataSet();

                                sqlAdapter[DB].SelectCommand = cmd[DB];
                                sqlAdapter[DB].Fill(ds[DB]);

                                k = 0;

                                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                                {
                                    int r = 0;

                                    for (int n = 0; n < dr.ItemArray.Length; n++)
                                    {
                                        D_Data[r][0][row] = dr.ItemArray[n];
                                        r++;
                                    }

                                    row++;
                                    r++;
                                }



                            }

                        }

                    }

                    if (DB == 0)
                    {
                        string[] ID = Array.ConvertAll<object, string>(D_Data[0][0], Convert.ToString);
                        string[] WAFER_ID = Array.ConvertAll<object, string>(D_Data[1][0], Convert.ToString);
                        string[] LOT_ID = Array.ConvertAll<object, string>(D_Data[2][0], Convert.ToString);
                        string[] SITE_ID = Array.ConvertAll<object, string>(D_Data[3][0], Convert.ToString);

                        int m = 4;

                        foreach (KeyValuePair<int, string> _L in T)
                        {
                            double[] Data = Array.ConvertAll<object, double>(D_Data[m][0], Convert.ToDouble);

                            int i = _L.Key;

                            string NPI_Spec_Min = Convert.ToString(this.Data.Clotho_Spcc_List[i].Min[0]);
                            string NPI_Spec_Max = Convert.ToString(this.Data.Clotho_Spcc_List[i].Max[0]);

                            string Customer_Spec_Min = Convert.ToString(this.Data.Customor_Clotho_List[i].Min[0]);
                            string Customer_Spec_Max = Convert.ToString(this.Data.Customor_Clotho_List[i].Max[0]);

                            CSV_Class.For_Box Box = new CSV_Class.For_Box(this.Data.Reference_Header[i], Data, ID, WAFER_ID, SITE_ID, LOT_ID, 0, 0, "", "", "", Customer_Spec_Min, Customer_Spec_Max, NPI_Spec_Min, NPI_Spec_Max);

                            Dic_Test[DB].Add(this.Data.Reference_Header[i], Box);
                            m++;
                        }
                    }
                    else
                    {
                        string[] ID = Array.ConvertAll<object, string>(D_Data[0][0], Convert.ToString);
                        string[] WAFER_ID = Array.ConvertAll<object, string>(D_Data[1][0], Convert.ToString);

                        string[] LOT_ID = Array.ConvertAll<object, string>(D_Data[2][0], Convert.ToString);
                        string[] SITE_ID = Array.ConvertAll<object, string>(D_Data[3][0], Convert.ToString);

                        int m = 4;

                        foreach (KeyValuePair<int, string> _L in T)
                        {
                            double[] Data = Array.ConvertAll<object, double>(D_Data[m][0], Convert.ToDouble);

                            int i = _L.Key;

                            string NPI_Spec_Min = Convert.ToString(this.Data.Clotho_Spcc_List[i].Min[0]);
                            string NPI_Spec_Max = Convert.ToString(this.Data.Clotho_Spcc_List[i].Max[0]);

                            string Customer_Spec_Min = Convert.ToString(this.Data.Customor_Clotho_List[i].Min[0]);
                            string Customer_Spec_Max = Convert.ToString(this.Data.Customor_Clotho_List[i].Max[0]);

                            CSV_Class.For_Box Box = new CSV_Class.For_Box(this.Data.Reference_Header[i], Data, ID, WAFER_ID, SITE_ID, LOT_ID, 0, 0, "", "", "", Customer_Spec_Min, Customer_Spec_Max, NPI_Spec_Min, NPI_Spec_Max);

                            Dic_Test[DB].Add(this.Data.Reference_Header[i], Box);
                            m++;
                        }
                    }




                    sqlAdapter[DB].Dispose();

                    stringA[DB].Clear();
                }




                ThreadFlags[DB].Set();
            }

            public object[] Get_Selected_Para_Thread()
            {
                object[] ID_Dummy = new object[0];
                object[] WAFERID_Dummy = new object[0];
                object[] LOTID_Dummy = new object[0];
                object[] SITEID_Dummy = new object[0];
                object[] Value_Dummy = new object[0];

                object[] ID_Test = new object[0];
                object[] WAFERID_Test = new object[0];
                object[] LOTID_Test = new object[0];
                object[] SITEID_Test = new object[0];
                object[] Value_Test = new object[0];

                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    int count = 0;
                    Dictionary<string, List<string>> tests = key.Value;


                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {
                        stringA[0].Clear();
                        stringA[0].Append("Select WAFER_ID from " + key.Key + " where Fail not like '1'");

                        cmd[0] = new SQLiteCommand(conn[0]);
                        sqlAdapter[0] = new SQLiteDataAdapter();


                        cmd[0].CommandText = stringA[0].ToString();
                        ds[0] = new DataSet();

                        sqlAdapter[0].SelectCommand = cmd[0];
                        sqlAdapter[0].Fill(ds[0]);

                        WAFERID_Dummy = new object[ds[0].Tables[0].Rows.Count];


                        foreach (DataRow dr in ds[0].Tables[0].Rows)
                        {
                            WAFERID_Dummy[count] = dr.ItemArray[0];
                            count++;
                        }

                        WAFERID_Test = WAFERID_Test.Concat(WAFERID_Dummy).ToArray();
                        sqlAdapter[0].Dispose();
                    }
                }

                return WAFERID_Test;
            }

            public void Get_Selected_Para(int DB, string Select_Para, bool Flag, string Selector)
            {
                ID = new object[0];
                Value = new object[0];
                Variation = new object[0];

                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    int count = 0;
                    Dictionary<string, List<string>> tests = key.Value;


                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {
                        stringA[DB].Clear();

                        if (Selector == "BIN")
                        {
                            stringA[DB].Append("Select id, " + Select_Para + " , BIN from " + key.Key + "  where FAIL not like '1'");
                        }
                        else if (Selector == "SITE")
                        {
                            stringA[DB].Append("Select id, " + Select_Para + " , SITEID from " + key.Key + "  where FAIL not like '1'");
                        }
                        else if (Selector == "LOT")
                        {
                            stringA[DB].Append("Select id, " + Select_Para + " , LOTID from " + key.Key + "  where FAIL not like '1'");
                        }

                        cmd[DB] = new SQLiteCommand(conn[DB]);
                        sqlAdapter[DB] = new SQLiteDataAdapter();
                        cmd[DB].CommandText = stringA[DB].ToString();
                        ds[DB] = new DataSet();

                        sqlAdapter[DB].SelectCommand = cmd[DB];
                        sqlAdapter[DB].Fill(ds[DB]);

                        object[] ID_Dummy = new object[ds[DB].Tables[0].Rows.Count];
                        object[] Value_Dummy = new object[ds[DB].Tables[0].Rows.Count];
                        object[] Variation_Dummy = new object[ds[DB].Tables[0].Rows.Count];

                        foreach (DataRow dr in ds[DB].Tables[0].Rows)
                        {
                            ID_Dummy[count] = dr.ItemArray[0];
                            Value_Dummy[count] = dr.ItemArray[1];
                            Variation_Dummy[count] = dr.ItemArray[2];
                            count++;
                        }
                        ID = ID.Concat(ID_Dummy).ToArray();
                        Value = Value.Concat(Value_Dummy).ToArray();
                        Variation = Variation.Concat(Variation_Dummy).ToArray();

                        sqlAdapter[DB].Dispose();
                        cmd[DB].Dispose();
                    }

                }


          
      
                stringA[DB].Clear();

            }
            public double[] Get_Find_Bin(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);

                stringA[0].Clear();
                return doubles;
            }
            public List<object[]> Get_Data_By_Querys(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                List<object[]> _Data = new List<object[]>(); 

                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;


                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    _Data.Add(dr.ItemArray);
                    count++;
                }
               

                sqlAdapter[0].Dispose();
                cmd[0].Dispose();
                stringA[0].Clear();

                return _Data;
            }
            public string[] Get_Data_By_Query(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;

                if (Query.Contains("PRAGMA"))
                {
                    foreach (DataRow dr in ds[0].Tables[0].Rows)
                    {
                        Value[count] = dr.ItemArray[1];
                        count++;
                    }
                }
                else
                {
                    foreach (DataRow dr in ds[0].Tables[0].Rows)
                    {
                        Value[count] = dr.ItemArray[0];
                        count++;
                    }
                }
              
         

                string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);
                sqlAdapter[0].Dispose();
                cmd[0].Dispose();
                stringA[0].Clear();
                return _string;
            }
            public string[] Get_Data_By_Query(string Query, int DB)
            {
                stringA[DB].Clear();
                stringA[DB].Append(Query);

                //  sqlAdapter[DB] = new SQLiteDataAdapter();
                //  cmd[DB] = new SQLiteCommand(conn[DB]);
                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                Value = new object[ds[DB].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);
                //    cmd[0].Dispose();
                stringA[DB].Clear();
                return _string;
            }
            public Dictionary<string, object[]> Get_Data_By_Query_S4PD(string Query, string Chan)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                // cmd[0].CommandText = stringA[0].ToString();
                // ds[0] = new DataSet();

                // sqlAdapter[0].SelectCommand = cmd[0];
                // sqlAdapter[0].Fill(ds[0]);

                // Value = new object[ds[0].Tables[0].Rows.Count];

                //// int count = 0;
                // foreach (DataRow dr in ds[0].Tables[0].Rows)
                // {
                //     Value[count] = dr.ItemArray[0];
                //     count++;
                // }

                //  string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);
                // SqReader[0] = cmd[0].ExecuteReader();
                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[500000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    Value1[count] = stringD[0];

                    count++;

                }

                Array.Resize(ref Value1, count);

                cmd[0].Dispose();
                SqReader[0].Close();

                string[] _string = Array.ConvertAll<object, string>(Value1, Convert.ToString);


                stringA[0].Clear();
                return null;
            }
            public void Get_Defined_Para(object[,] DummyData, string key, Data_Class.Data_Editing.INT Data_InterFace)
            {


            }
            public void Get_Gross_Check_Para(Data_Class.Data_Editing.INT Data_Edit, string Select_Para, double Persent, string Selector, int SelectedBin)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                Get_Gross_Para = Select_Para;
                Get_Gross_Persent = Persent;
                Get_Gross_Selector = Selector;
                Get_Gross_Selectedbin = SelectedBin;

                NB = 0;
                for (int s = 0; s < Table_Count; s++)
                {
                    Query = "Select count(id) from data" + s + " where Fail = '0'";

                    NB += Get_Sample_Count(0, Query);
                }

                ID = new object[NB];
                Value = new object[NB];

                //  Gross = ForGross_Fail_Unit;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = false;
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Gross_Check_Para_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
                double test = TestTime1.Elapsed.TotalMilliseconds;

                int count = 0;
                ID = new object[NB];

                for (int loop = 0; loop < Table_Count; loop++)
                {

                    stringA[0].Clear();
                    stringA[0].Append("Select id from data" + loop + " where FAIL not like '1'");
                    //stringA[0].Append("Select id from data");

                    cmd[0].CommandText = stringA[0].ToString();
                    ds[0] = new DataSet();

                    sqlAdapter[0].SelectCommand = cmd[0];
                    sqlAdapter[0].Fill(ds[0]);


                    foreach (DataRow dr in ds[0].Tables[0].Rows)
                    {
                        ID[count] = dr.ItemArray[0];
                        count++;
                    }
                }
                stringA[0].Clear();
                List_Gross_Values.Add(Gross_Values1);
            }
            public void Get_Gross_Check_Para_Thread(Object threadContext)
            {
                int i = (int)threadContext;



                int k = 0;
                for (k = 0; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    string[] Split_Dummy = Data.Reference_Header[Data.DB_Column_Limit * i + k].Split('_');
                    if (Split_Dummy.Length != 1)
                    {
                        if (Split_Dummy[1].ToUpper() == Get_Gross_Para.ToUpper())
                        {

                            object[] DataValue = new object[NB];
                            List<double[]> DataSet_Values = new List<double[]>();
                            DataValue = new object[NB];
                            int count = 0;

                            for (int loop = 0; loop < Table_Count; loop++)
                            {

                                ds[i] = new DataSet();
                                stringA[i].Clear();
                                //    conn[i] = new SQLiteConnection(strConn[i]);
                                //    cmd[i] = new SQLiteCommand(conn[i]);
                                //     conn[i].Open();
                                stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data" + loop + " where Fail not like '1'");

                                string a = "Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data where Fail not like '1'";
                                //  stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data");
                                cmd[i].CommandText = stringA[i].ToString();

                                sqlAdapter[i].SelectCommand = cmd[i];
                                sqlAdapter[i].Fill(ds[i]);



                                foreach (DataRow dr in ds[i].Tables[0].Rows)
                                {
                                    DataValue[count] = dr.ItemArray[0];

                                    count++;
                                }

                            }
                            double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

                            double DataMin = doubles.Min();
                            double DataMax = doubles.Max();
                            double DataAve = doubles.Average();

                            double DataMinindex = doubles.ToList().IndexOf(DataMin);
                            double DataMaxindex = doubles.ToList().IndexOf(DataMax);

                            double Divide = DataMax / DataMin;

                            string[] test;
                            string _Substring = Get_Gross_Para.Substring(0, 1);

                            double MinSpec = 0f;
                            bool Define_Flag = false;

                            if (Get_Gross_Selector == "MAX/MIN")
                            {
                                Define_Flag = true;
                                test = Convert.ToString(Get_Gross_Persent).Split('.');
                                MinSpec = 1 - (Convert.ToDouble(test[1]) / 10);
                            }
                            else if (Get_Gross_Selector == "MAX-MIN")
                            {
                                Define_Flag = false;
                                MinSpec = Convert.ToDouble(Get_Gross_Persent) * -1;
                            }

                            double std = STDEVandMedian(doubles, i, count);
                            Gross Gross_data = new Gross(doubles, std, Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Min[Get_Gross_Selectedbin], Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k].Max[Get_Gross_Selectedbin]);

                            Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], Gross_data);

                            //if (Define_Flag)
                            //{
                            //    for (int j = 0; j < doubles.Length; j++)
                            //    {
                            //        if (DataAve / doubles[j] > Get_Gross_Persent || DataAve / doubles[j] < MinSpec)
                            //        {
                            //            if (!Gross_Values1[i].ContainsKey(Convert.ToString(j + 1)))
                            //            {
                            //                double std = STDEVandMedian(doubles, i, count);
                            //                Gross Gross_data = new Gross(doubles, std);

                            //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], Gross_data); break;
                            //            }
                            //        }
                            //    }
                            //}
                            //else
                            //{
                            //    for (int j = 0; j < doubles.Length; j++)
                            //    {
                            //        if (DataAve - doubles[j] > Get_Gross_Persent || doubles[j] - DataAve < MinSpec)
                            //        {
                            //            if (!Gross_Values1[i].ContainsKey(Convert.ToString(j + 1)))
                            //            {
                            //                double std = STDEVandMedian(doubles, i, count);
                            //                Gross Gross_data = new Gross(doubles, std);

                            //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], Gross_data); break;
                            //            }
                            //        }
                            //    }
                            //}

                            stringA[i].Clear();
                            cmd[i].CommandText = "";
                        }


                    }
                }
                ThreadFlags[i].Set();
            }
            public void Get_Current_Setting(Data_Class.Data_Editing.INT Data_Edit, int NB)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;
                this.Count_Current_Setting = NB;

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Get_Current_Setting_Thread(i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                 //   Wait[i] = ThreadFlags[i].WaitOne();
                }


            }

            //public void Get_Current_Setting_Thread(Object threadContext)
            //{
            //    int i = (int)threadContext;

            //    int count = 0;



            //    cmd[i] = new SQLiteCommand(conn[i]);
            //    //   conn[i].Open();

            //    stringA[i].Clear();
            //    stringA[i].Append("Select * from Current_Setting where id = " + this.Count_Current_Setting);

            //    cmd[i].CommandText = stringA[i].ToString();
            //    SqReader[i] = cmd[i].ExecuteReader();

            //    count = 0;
            //    int k = 0;

            //    int Db_Limit = Data.DB_Column_Limit;

            //    while (SqReader[i].Read())
            //    {

            //        Stopwatch TestTime1 = new Stopwatch();
            //        TestTime1.Restart();
            //        TestTime1.Start();


            //        object[] values = new object[SqReader[i].FieldCount];
            //        SqReader[i].GetValues(values);

            //        int ForCount = 0;
            //        int j = 0;
            //        if (i == Data.DB_Count - 1)
            //        {
            //            ForCount = values.Length - 6;
            //        }
            //        else if (i == 0)
            //        {
            //            ForCount = values.Length - 6;

            //        }
            //        else
            //        {
            //            ForCount = values.Length - 6;
            //        }


            //        if (i == 0)
            //        {
            //            if (k == 0)
            //            {
            //                No_Index[0] = values[0].ToString();
            //            }
            //            else if (k == 1)
            //            {
            //                Paraname[0] = values[0].ToString();
            //            }
            //            else if (k == 2)
            //            {
            //                SpecMin[0] = values[0].ToString();
            //            }
            //            else if (k == 3)
            //            {
            //                SpecMax[0] = values[0].ToString();
            //            }
            //            else if (k == 4)
            //            {
            //                DataMin[0] = values[0].ToString();
            //            }
            //            else if (k == 5)
            //            {
            //                DataMedian[0] = values[0].ToString();
            //            }

            //            else if (k == 6)
            //            {
            //                DataMax[0] = values[0].ToString();
            //            }
            //            else if (k == 7)
            //            {
            //                CPK[0] = values[0].ToString();
            //            }
            //            else if (k == 8)
            //            {
            //                STD[0] = values[0].ToString();
            //            }
            //            else if (k == 9)
            //            {
            //                Percent[0] = values[0].ToString();
            //            }
            //            else if (k == 10)
            //            {
            //                Fail[0] = values[0].ToString();
            //            }
            //            for (j = 10; j < ForCount; j++)
            //            {
            //                if(k == 0)
            //                {
            //                    No_Index[j - 9] = values[j].ToString();
            //                }
            //                else if(k == 1)
            //                {
            //                    if( k == 1)
            //                    {

            //                    }
            //                    if(j == ForCount - 1)
            //                    {

            //                    }
            //                    Paraname[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 2)
            //                {
            //                    SpecMin[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 3)
            //                {
            //                    SpecMax[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 4)
            //                {
            //                    DataMin[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 5)
            //                {
            //                    DataMedian[j - 9] = values[j].ToString();
            //                }

            //                else if (k == 6)
            //                {
            //                    DataMax[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 7)
            //                {
            //                    CPK[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 8)
            //                {
            //                    STD[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 9)
            //                {
            //                    Percent[j - 9] = values[j].ToString();
            //                }
            //                else if (k == 10)
            //                {
            //                    Fail[j - 9] = values[j].ToString();
            //                }

            //            }

            //        }
            //        else
            //        {
            //            if (k == 0)
            //            {
            //                No_Index[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 1)
            //            {
            //                Paraname[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 2)
            //            {
            //                SpecMin[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 3)
            //            {
            //                SpecMax[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 4)
            //            {
            //                DataMin[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 5)
            //            {
            //                DataMedian[Db_Limit * i - 9] = values[j].ToString();
            //            }

            //            else if (k == 6)
            //            {
            //                DataMax[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 7)
            //            {
            //                CPK[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 8)
            //            {
            //                STD[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 9)
            //            {
            //                Percent[Db_Limit * i - 9] = values[j].ToString();
            //            }
            //            else if (k == 10)
            //            {
            //                Fail[Db_Limit * i - 9] = values[j].ToString();
            //            }

            //            for (j = 1; j < ForCount; j++)
            //            {
            //                if (k == 0)
            //                {
            //                    No_Index[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 1)
            //                {
            //                    Paraname[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 2)
            //                {
            //                    SpecMin[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 3)
            //                {
            //                    SpecMax[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 4)
            //                {
            //                    DataMin[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 5)
            //                {
            //                    DataMedian[Db_Limit * i + j - 9] = values[j].ToString();
            //                }

            //                else if (k == 6)
            //                {
            //                    DataMax[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 7)
            //                {
            //                    CPK[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 8)
            //                {
            //                    STD[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 9)
            //                {
            //                    Percent[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //                else if (k == 10)
            //                {
            //                    Fail[Db_Limit * i + j - 9] = values[j].ToString();
            //                }
            //            }

            //        }

            //        k++;

            //        double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
            //        count++;

            //    }

            //    SqReader[i].Close();

            //    stringA[i].Clear();
            //    cmd[i].CommandText = "";
            //    cmd[i].Dispose();
            //    //   conn[i].Dispose();



            //   // ThreadFlags[i].Set();


            //}

            public void Get_Current_Setting_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                int count = 0;



                cmd[i] = new SQLiteCommand(conn[i]);
                //   conn[i].Open();

                stringA[i].Clear();
                stringA[i].Append("Select * from Current_Setting where id = " + this.Count_Current_Setting);

                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                count = 0;
                int k = 0;

                int Db_Limit = Data.DB_Column_Limit;

                while (SqReader[i].Read())
                {

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);

                    int ForCount = 0;
                    int j = 0;
                    if (i == Data.DB_Count - 1)
                    {
                        ForCount = values.Length - 6;
                    }
                    else if (i == 0)
                    {
                        ForCount = values.Length - 6;

                    }
                    else
                    {
                        ForCount = values.Length - 6;
                    }


                    if (i == 0)
                    {
                        if (k == 0)
                        {
                            No_Index[0] = values[0].ToString();
                        }
                        else if (k == 1)
                        {
                            Paraname[0] = values[0].ToString();
                        }
                        else if (k == 2)
                        {
                            SpecMin[0] = values[0].ToString();
                        }
                        else if (k == 3)
                        {
                            SpecMax[0] = values[0].ToString();
                        }
                        else if (k == 4)
                        {
                            DataMin[0] = values[0].ToString();
                        }
                        else if (k == 5)
                        {
                            DataMedian[0] = values[0].ToString();
                        }

                        else if (k == 6)
                        {
                            DataMax[0] = values[0].ToString();
                        }
                        else if (k == 7)
                        {
                            CPK[0] = values[0].ToString();
                        }
                        else if (k == 8)
                        {
                            STD[0] = values[0].ToString();
                        }
                        else if (k == 9)
                        {
                            Percent[0] = values[0].ToString();
                        }
                        else if (k == 10)
                        {
                            Fail[0] = values[0].ToString();
                        }
                        for (j = 10; j < ForCount; j++)
                        {
                            if (k == 0)
                            {
                                No_Index[j - 9] = values[j].ToString();
                            }
                            else if (k == 1)
                            {
                                if (k == 1)
                                {

                                }
                                if (j == ForCount - 1)
                                {

                                }
                                Paraname[j - 9] = values[j].ToString();
                            }
                            else if (k == 2)
                            {
                                SpecMin[j - 9] = values[j].ToString();
                            }
                            else if (k == 3)
                            {
                                SpecMax[j - 9] = values[j].ToString();
                            }
                            else if (k == 4)
                            {
                                DataMin[j - 9] = values[j].ToString();
                            }
                            else if (k == 5)
                            {
                                DataMedian[j - 9] = values[j].ToString();
                            }

                            else if (k == 6)
                            {
                                DataMax[j - 9] = values[j].ToString();
                            }
                            else if (k == 7)
                            {
                                CPK[j - 9] = values[j].ToString();
                            }
                            else if (k == 8)
                            {
                                STD[j - 9] = values[j].ToString();
                            }
                            else if (k == 9)
                            {
                                Percent[j - 9] = values[j].ToString();
                            }
                            else if (k == 10)
                            {
                                Fail[j - 9] = values[j].ToString();
                            }

                        }

                    }
                    else
                    {
                        if (k == 0)
                        {
                            No_Index[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 1)
                        {
                            Paraname[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 2)
                        {
                            SpecMin[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 3)
                        {
                            SpecMax[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 4)
                        {
                            DataMin[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 5)
                        {
                            DataMedian[Db_Limit * i - 9] = values[j].ToString();
                        }

                        else if (k == 6)
                        {
                            DataMax[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 7)
                        {
                            CPK[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 8)
                        {
                            STD[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 9)
                        {
                            Percent[Db_Limit * i - 9] = values[j].ToString();
                        }
                        else if (k == 10)
                        {
                            Fail[Db_Limit * i - 9] = values[j].ToString();
                        }

                        for (j = 1; j < ForCount; j++)
                        {
                            if (k == 0)
                            {
                                No_Index[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 1)
                            {
                                Paraname[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 2)
                            {
                                SpecMin[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 3)
                            {
                                SpecMax[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 4)
                            {
                                DataMin[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 5)
                            {
                                DataMedian[Db_Limit * i + j - 9] = values[j].ToString();
                            }

                            else if (k == 6)
                            {
                                DataMax[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 7)
                            {
                                CPK[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 8)
                            {
                                STD[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 9)
                            {
                                Percent[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                            else if (k == 10)
                            {
                                Fail[Db_Limit * i + j - 9] = values[j].ToString();
                            }
                        }

                    }

                    k++;

                    double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                    count++;

                }

                SqReader[i].Close();

                stringA[i].Clear();
                cmd[i].CommandText = "";
                cmd[i].Dispose();
                //   conn[i].Dispose();



                // ThreadFlags[i].Set();


            }

            public void Get_From_Db_Data_for_Anly(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;
                Each_Thread_Count = new int[Data.DB_Count];
                ForCampare_Yield = new List<List<int>[]>[Data.DB_Count];
                For_Any_Yield = new List<List<int>>[Data.DB_Count];

                For_Any_Yield_For_Lot = new List<List<List<int>>>[Data.DB_Count];
                For_Any_Yield_For_SITE = new List<List<List<int>>>[Data.DB_Count];

                For_Any_Yield_Percent = new List<List<int>[]>[Data.DB_Count];
                Yield_Test = new List<List<RowAndPass>[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield[i] = new List<List<int>[]>();
                    For_Any_Yield[i] = new List<List<int>>();
                    For_Any_Yield_For_Lot[i] = new List<List<List<int>>>();
                    For_Any_Yield_For_SITE[i] = new List<List<List<int>>>();
                    For_Any_Yield_Percent[i] = new List<List<int>[]>();
                    Yield_Test[i] = new List<List<RowAndPass>[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    for (int k = 0; k < this.Data.Clotho_Spcc_List[1].Max.Length; k++)
                    {
                        For_Any_Yield_For_Lot[i].Add(new List<List<int>>());
                        For_Any_Yield_For_SITE[i].Add(new List<List<int>>());
                    }

                }



                for (int i = 0; i < Data.DB_Count; i++)
                {
                    int k = 0;
                    for (k = 0; k < 1; k++)
                    {
                        List<int> dummy = new List<int>();

                        for (k = 0; k < this.Data.Clotho_Spcc_List[1].Max.Length; k++)
                        {
                            dummy = new List<int>();

                            for (int n = 0; n < Data.Per_DB_Column_Count[i]; n++)
                            {
                                dummy.Add(0);

                            }

                            For_Any_Yield[i].Add(dummy);
                        }
                    }
                }


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    int k = 0;
                    for (k = 0; k < 1; k++)
                    {
                        List<int> dummy = new List<int>();

                        for (k = 0; k < this.Data.Clotho_Spcc_List[1].Max.Length; k++)
                        {
                            for (int e = 0; e < Lot_Dic.Count; e++)
                            {
                                dummy = new List<int>();

                                for (int n = 0; n < Data.Per_DB_Column_Count[i]; n++)
                                {
                                    //  foreach (KeyValuePair<string, int> data in Lot_Dic)
                                    //  {
                                    dummy.Add(0);
                                    //    }
                                }

                                For_Any_Yield_For_Lot[i][k].Add(dummy);
                            }
                        }
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    int k = 0;
                    for (k = 0; k < 1; k++)
                    {
                        List<int> dummy = new List<int>();

                        for (k = 0; k < this.Data.Clotho_Spcc_List[1].Max.Length; k++)
                        {
                            for (int e = 0; e < Site_Dic.Count; e++)
                            {
                                dummy = new List<int>();

                                for (int n = 0; n < Data.Per_DB_Column_Count[i]; n++)
                                {
                                    //  foreach (KeyValuePair<string, int> data in Site_Dic)
                                    //   {
                                    dummy.Add(0);
                                    //   }
                                }

                                For_Any_Yield_For_SITE[i][k].Add(dummy);
                            }
                        }
                    }
                }


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                //   Get_From_Db_Data_for_Anly_Thread(3);
                   ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data_for_Anly_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Get_From_Db_Data_for_Anly_Thread(Object threadContext)
            {
                int i = (int)threadContext;


                int count = 0;


                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    Dictionary<string, List<string>> tests = key.Value;


                    foreach (KeyValuePair<string, List<string>> ts in tests)
                    {

                    //    conn[i] = new SQLiteConnection(strConn[i]);
                        cmd[i] = new SQLiteCommand(conn[i]);
                     //   conn[i].Open();

                        stringA[i].Clear();
                        stringA[i].Append("Select * from " + key.Key + " where Fail = '0'");

                        cmd[i].CommandText = stringA[i].ToString();
                        SqReader[i] = cmd[i].ExecuteReader();

                        count = 0;

                        while (SqReader[i].Read())
                        {

                            Stopwatch TestTime1 = new Stopwatch();
                            TestTime1.Restart();
                            TestTime1.Start();


                            object[] values = new object[SqReader[i].FieldCount];
                            SqReader[i].GetValues(values);

                            string Lot = values[values.Length - 4].ToString();
                            string Site = values[values.Length - 3].ToString();

                            int Lot_Int = Lot_Dic[Lot];
                            int Site_Int = Site_Dic[Site];
                            long SN = Convert.ToInt64(values[values.Length - 5]);
                            int ForCount = 0;

                            if (i == Data.DB_Count - 1)
                            {
                                ForCount = values.Length - 12;
                                values[values.Length - 11] = 0;
                                values[values.Length - 10] = 0;
                                values[values.Length - 10] = 0;
                                values[values.Length - 9] = 0;
                                values[values.Length - 8] = 0;
                                values[values.Length - 7] = 0;
                                values[values.Length - 6] = 0;
                                values[values.Length - 5] = 0;
                                values[values.Length - 4] = 0;
                                values[values.Length - 3] = 0;
                                values[values.Length - 2] = 0;
                                values[values.Length - 1] = 0;
                            }
                            else if (i == 0)
                            {
                                ForCount = values.Length - 6;
                                values[5] = 0;
                                values[8] = 0;
                                values[9] = 0;
                                values[values.Length - 6] = 0;
                                values[values.Length - 5] = 0;
                                values[values.Length - 4] = 0;
                                values[values.Length - 3] = 0;
                                values[values.Length - 2] = 0;
                                values[values.Length - 1] = 0;
                                values[0] = values[0].ToString().Remove(0, 4);
                            }
                            else
                            {
                                ForCount = values.Length - 6;


                                values[values.Length - 6] = 0;
                                values[values.Length - 5] = 0;
                                values[values.Length - 4] = 0;
                                values[values.Length - 3] = 0;
                                values[values.Length - 2] = 0;
                                values[values.Length - 1] = 0;

                            }


                            double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);

                            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                            List<RowAndPass>[] Check3 = new List<RowAndPass>[this.Data.Clotho_Spcc_List[1].Max.Length];
                            int k = 0;
                            for (k = 0; k < this.Data.Clotho_Spcc_List[1].Max.Length; k++)
                            {
                                Check3[k] = new List<RowAndPass>();
                            }

                            double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;


                            int j = 0;
                            int m = 0;

                            int Index_For = this.Data.Clotho_Spcc_List[0].Max.Length;
                            int Db_Limit = Data.DB_Column_Limit;


                            if (i == 0)
                            {

                                for (j = 10; j < ForCount; j++)
                                {
                                    for (m = 0; m < Index_For; m++)
                                    {
                                        if (this.Data.Clotho_Spcc_List[Db_Limit * i + j - 9].Max[m] < doubles[j] || this.Data.Clotho_Spcc_List[(Db_Limit * i) + j - 9].Min[m] > doubles[j])
                                        {
                                 
                                            //
                                            RowAndPass data = new RowAndPass(0, 0, 0);
                                            data.SN = SN;
                                            data.Row = j;
                                            data.Pass = 1;

                                            Check3[m].Add(data);
                                            For_Any_Yield[i][m][j] += 1;
                                            For_Any_Yield_For_Lot[i][m][Lot_Int][j] += 1;
                                            For_Any_Yield_For_SITE[i][m][Site_Int][j] += 1;


                                        }
                                    }
                                }

                            }
                            else
                            {

                                for (j = 0; j < ForCount; j++)
                                {
                                    for (m = 0; m < Index_For; m++)
                                    {
                                        if (this.Data.Clotho_Spcc_List[Db_Limit * i + j - 9].Max[m] < doubles[j] || this.Data.Clotho_Spcc_List[Db_Limit * i + j - 9].Min[m] > doubles[j])
                                        {
                         
                                            RowAndPass data = new RowAndPass(0, 0, 0);
                                            data.SN = SN;
                                            data.Row = j;
                                            data.Pass = 1;

                                            Check3[m].Add(data);
                                            For_Any_Yield[i][m][j] += 1;
                                            For_Any_Yield_For_Lot[i][m][Lot_Int][j] += 1;
                                            For_Any_Yield_For_SITE[i][m][Site_Int][j] += 1;
                                        }
                                    }
                                }

                            }

                            Yield_Test[i].Add(Check3);

                            double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                            count++;


                            Each_Thread_Count[i]++;
                        }
                        SqReader[i].Close();

                        stringA[i].Clear();
                        cmd[i].CommandText = "";
                        cmd[i].Dispose();
                     //   conn[i].Dispose();
                    }

                }


               // cmd[i] = new SQLiteCommand(conn[i]);
               // conn[i].Open();

                ThreadFlags[i].Set();


            }
            public void Get_From_Db_Data_for_Anly_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;

                // For_New_Spec_ForCampare_Yield = new List<List<int>[]>[Data.DB_Count];
                For_Any_Yield_For_New_Spec = new List<List<int>>[Data.DB_Count];
                // For_Any_Yield_Percent_For_New_Spec = new List<List<int>[]>[Data.DB_Count];
                // For_New_Spec_ForCampare_Yield2 = new List<List<int>>[Data.DB_Count];
                Yield_Test_New_Spec = new List<List<RowAndPass>[]>[Data.DB_Count];


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //  For_New_Spec_ForCampare_Yield[i] = new List<List<int>[]>();
                    For_Any_Yield_For_New_Spec[i] = new List<List<int>>();
                    //    For_Any_Yield_Percent_For_New_Spec[i] = new List<List<int>[]>();
                    //   For_New_Spec_ForCampare_Yield2[i] = new List<List<int>>();
                    Yield_Test_New_Spec[i] = new List<List<RowAndPass>[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    int k = 0;
                    for (k = 0; k < 1; k++)
                    {
                        List<int> dummy = new List<int>();
                        List<int> dummy2 = new List<int>();

                        for (k = 0; k < this.Data.Clotho_List[1].Max.Length; k++)
                        {
                            dummy = new List<int>();
                            dummy2 = new List<int>();
                            for (int n = 0; n < Data.Per_DB_Column_Count[i]; n++)
                            {
                                dummy.Add(0);
                                dummy2.Add(0);
                            }

                            For_Any_Yield_For_New_Spec[i].Add(dummy);
                            //   For_New_Spec_ForCampare_Yield2[i].Add(dummy2);
                        }
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data_for_Anly_For_New_Spec_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    //   conn[i] = new SQLiteConnection(strConn[i]);
                    //  cmd[i] = new SQLiteCommand(conn[i]);
                    //  conn[i].Open();
                }

                //for (int i = 0; i < 1; i++)
                //{
                //    stringA[6].Clear();
                //    ThreadFlags[6] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data_for_Anly_For_New_Spec_Thread), 6);
                //}
                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[6] = ThreadFlags[i].WaitOne();
                //}

            }
            public void Get_From_Db_Data_for_Anly_For_New_Spec_Thread(Object threadContext)
            {
                int i = (int)threadContext;


                int count = 0;

                for (int loop = 0; loop < Table_Count; loop++)
                {

                    stringA[i].Append("Select * from data" + loop + " where Fail not like '1'");

                    //  conn[i] = new SQLiteConnection(strConn[i]);
                    //  cmd[i] = new SQLiteCommand(conn[i]);
                    //  conn[i].Open();

                    cmd[i].CommandText = stringA[i].ToString();
                    SqReader[i] = cmd[i].ExecuteReader();

                    count = 0;

                    while (SqReader[i].Read())
                    {

                        Stopwatch TestTime1 = new Stopwatch();
                        TestTime1.Restart();
                        TestTime1.Start();


                        object[] values = new object[SqReader[i].FieldCount];
                        SqReader[i].GetValues(values);

                        string Lot = values[Data.Per_DB_Column_Count[i] + 2].ToString();
                        string Site = values[Data.Per_DB_Column_Count[i] + 5].ToString();
                        int SN = Convert.ToInt16(values[Data.Per_DB_Column_Count[i]]);

                        values[Data.Per_DB_Column_Count[i] + 2] = 0;
                        values[Data.Per_DB_Column_Count[i] + 3] = 0;
                        values[Data.Per_DB_Column_Count[i] + 6] = 0;

                        double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);

                        //    List<int>[] Check = new List<int>[this.Data.Clotho_List[1].Max.Length];
                        //  List<int>[] Check2 = new List<int>[this.Data.Clotho_List[1].Max.Length];

                        List<RowAndPass>[] Check3 = new List<RowAndPass>[this.Data.Customor_Clotho_List[1].Max.Length];

                        int k = 0;
                        for (k = 0; k < this.Data.Customor_Clotho_List[1].Max.Length; k++)
                        {
                            //       Check[k] = new List<int>();
                            //     Check2[k] = new List<int>();
                            Check3[k] = new List<RowAndPass>();
                            //       Check2[k].Add(0);
                            //for (int n = 0; n < Data.Per_DB_Column_Count[i]; n++)
                            //{
                            //    Check[k].Add(0);
                            //}
                        }

                        //   double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;


                        int j = 0;
                        int m = 0;

                        int Index_For = this.Data.Customor_Clotho_List[0].Max.Length;

                        if (i == 0)
                        {
                            //        Check[0][0] = 0;
                        }
                        else
                        {
                            for (m = 0; m < Index_For; m++)
                            {
                                if (this.Data.Customor_Clotho_List[Data.DB_Column_Limit * i].Max[m] < doubles[0] || this.Data.Customor_Clotho_List[Data.DB_Column_Limit * i].Min[m] > doubles[0])
                                {
                                    //     Check[m][j] = 1;
                                    //      Check2[m][0] = 1;


                                    RowAndPass data = new RowAndPass(0, 0, 0);
                                    data.SN = SN;
                                    data.Row = 0;
                                    data.Pass = 1;

                                    Check3[m].Add(data);
                                    //      Yield_Test[i].Add(Check3);
                                    For_Any_Yield_For_New_Spec[i][m][0] += 1;
                                    //      For_New_Spec_ForCampare_Yield2[i][m][0] += 1;
                                }
                            }

                        }

                        for (j = 1; j < values.Length - 8; j++)
                        {
                            for (m = 0; m < Index_For; m++)
                            {
                                if (this.Data.Customor_Clotho_List[Data.DB_Column_Limit * i + j].Max[m] < doubles[j] || this.Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + j].Min[m] > doubles[j])
                                {

                                    //     Check[m][j] = 1;
                                    //      Check2[m][0] = 1;


                                    RowAndPass data = new RowAndPass(0, 0, 0);
                                    data.SN = SN;
                                    data.Row = j;
                                    data.Pass = 1;


                                    Check3[m].Add(data);
                                    // Yield_Test[i].Add(Check3);
                                    For_Any_Yield_For_New_Spec[i][m][j] += 1;
                                    //          For_New_Spec_ForCampare_Yield2[i][m][j] += 1;
                                }
                            }
                        }
                        for (m = 0; m < Index_For; m++)
                        {
                            if (this.Data.Customor_Clotho_List[Data.DB_Column_Limit * i + j].Max[m] < doubles[j] || this.Data.Customor_Clotho_List[(Data.DB_Column_Limit * i) + j].Min[m] > doubles[j])
                            {
                                //   Check[m][j] = 1;
                                //   Check2[m][0] = 1;


                                RowAndPass data = new RowAndPass(0, 0, 0);
                                data.SN = SN;
                                data.Row = j;
                                data.Pass = 1;
                                Check3[m].Add(data);
                                //   Yield_Test[i].Add(Check3);
                                For_Any_Yield_For_New_Spec[i][m][j] += 1;
                                //      For_New_Spec_ForCampare_Yield2[i][m][j] += 1;
                            }
                        }

                        Yield_Test_New_Spec[i].Add(Check3);
                        //   For_New_Spec_ForCampare_Yield[i].Add(Check);
                        //  For_Any_Yield_Percent_For_New_Spec[i].Add(Check2);

                        count++;

                        double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;

                    }
                    SqReader[i].Close();

                    //   cmd[i].Dispose();
                    //  conn[i].Close();

                    stringA[i].Clear();

                    // cmd[i].Dispose();
                }
                ThreadFlags[i].Set();


            }
            //public void Get_From_Db_Ref_Header(Data_Class.Data_Editing.INT Data_Edit)
            //{

            //    Stopwatch TestTime1 = new Stopwatch();
            //    TestTime1.Restart();
            //    TestTime1.Start();




            //    this.Data = Data_Edit;


            //    this.Data.Reference_Header_List = new List<string>();

            //    //for (int i = 0; i < Data.DB_Count; i++)
            //    //{
            //    //    count += Data.Per_DB_Column_Count[i];
            //    //}


            //    //this.Data.Reference_Header = new string[count];

            //    for (int i = 0; i < Data.DB_Count; i++)
            //    {
            //        stringA[i].Clear();
            //       // ThreadFlags[i] = new ManualResetEvent(false);
            //        Get_From_Db_Ref_Header_Thread(i);
            //       //   ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Ref_Header_Thread), i);
            //    }
            //    for (int i = 0; i < Data.DB_Count; i++)
            //    {//
            //     //   Wait[i] = ThreadFlags[i].WaitOne();
            //    }



            //}

            public void Get_From_Db_Ref_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();




                this.Data = Data_Edit;


                this.Data.Reference_Header_List = new List<string>();

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    count += Data.Per_DB_Column_Count[i];
                //}


                //this.Data.Reference_Header = new string[count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    // ThreadFlags[i] = new ManualResetEvent(false);
                    Get_From_Db_Ref_Header_Thread(i);
                    //   ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Ref_Header_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {//
                 //   Wait[i] = ThreadFlags[i].WaitOne();
                }



            }
            public void Get_From_Db_Ref_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                //  int Count_Data = 0;
                int count = 0;

                stringA[i].Clear();

                stringA[i].Append("Select * from REFHEADER");


                cmd[i] = new SQLiteCommand(conn[i]);
                sqlAdapter[i] = new SQLiteDataAdapter();

                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                count = 0;

                while (SqReader[i].Read())
                {

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    int ForCount = 0;

                    ForCount = values.Length - 5;

                    if(i == 0)
                    {
                        this.Data.Reference_Header_List.Add(Convert.ToString(values[0]));
                        for (int j = 10; j < ForCount; j++)
                        {
                            this.Data.Reference_Header_List.Add(Convert.ToString(values[j]));

                        }

                    }
                    else if (i == Data.DB_Count - 1 )
                    {           
                        for (int j = 0; j < ForCount -6; j++)
                        {
                            this.Data.Reference_Header_List.Add(Convert.ToString(values[j]));

                        }

                    }
                    else
                    {
                        for (int j = 0; j < ForCount; j++)
                        {
                            this.Data.Reference_Header_List.Add(Convert.ToString(values[j]));

                        }

                    }




                    double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                    count++;
                }
                SqReader[i].Close();
                cmd[i].Dispose();
                stringA[i].Clear();



               // ThreadFlags[i].Set();


            }
            public int Get_Sample_Count(int DB, string Query)
            {

                //conn[0] = new SQLiteConnection(strConn[0]);
                //cmd[0] = new SQLiteCommand(conn[0]);
                //sqlAdapter[0] = new SQLiteDataAdapter();


                stringA[DB].Clear();
                stringA[DB].Append(Query);

                cmd[DB] = new SQLiteCommand(conn[DB]);
                sqlAdapter[DB] = new SQLiteDataAdapter();

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                }

                //   sqlAdapter[0].Dispose();
                //   cmd[0].Dispose();

                //   conn[0].Dispose();

                sqlAdapter[DB].Dispose();
                //stringA[0].Clear();

                cmd[DB].Dispose();
                // conn[0].Close();


                int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Data_Count[0];

            }
            public int Get_Column_Count(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {
                return 0;
            }
            public void Close(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();
                    conn[i].Close();

                }
            }
            public void Read_Dispose(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    //sqlAdapter[i].Dispose();
                    //cmd[i].Dispose();
                    //conn[i].Dispose();
                }
            }
            public void Set_Conn(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    //conn[i] = new SQLiteConnection(strConn[i]);
                    //conn[i].Open();
                    //sqlAdapter[0] = new SQLiteDataAdapter();
                    //cmd[i] = new SQLiteCommand(conn[i]);
                    //cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    //     cmd[0].ExecuteNonQuery();


                }
            }
            public void trans(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;
                tran = new SQLiteTransaction[Data.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    tran[i] = conn[i].BeginTransaction();
                    cmd[i].Transaction = tran[i];
                }
            }
            public void Commit(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    tran[i].Commit();
                }


            }
            public void STDEVandMedian(List<double[]> Ds, int DB, int RowCount)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();


                double dummytesttime2 = TestTime1.Elapsed.TotalMilliseconds;
                int Para_Count = 0;

                double L_AVG = 0f;
                double H_AVG = 0f;
                double average = 0f;
                double Median = 0f;
                double dummyi = 0f;
                double dummyj = 0f;
                int GetMedian_i = 0;
                double minusSquareSummary = 0.0;

                int Count = 0;
                int Low_Count = 0;
                int High_Count = 0;

                double L_minusSquareSummary = 0f;
                double H_minusSquareSummary = 0f;

                int d = 0;


                double stdev = 0f;

                double L_stdev = 0f;
                double H_stdev = 0f;


                double[] values = new double[Ds.Count];

                int z = 0;
                int x = 0;
                int Offset = 0;
                int Offset2 = -5;


                if (DB == 0)
                {
                    z = 10;
                    x = 5;
                    Offset = 0;
                }
                else if (DB == Data.DB_Count - 1)
                {
                    z = 0;
                    x = 6;
                    Offset = -10;
                    Offset2 = -13;

                }
                else
                {
                    z = 0;
                    x = 6;
                    Offset = -10;

                }


                //  for (z = z; z < Ds[0].Length - (z + x); z++)
                for (z = z; z < Ds[0].Length + (Offset2); z++)
                {
                    for (int w = 0; w < Ds.Count; w++)
                    {
                        values[w] = Ds[w][z];
                    }

                    average = values.Average();
                    Array.Sort(values);

                    if (values.Length % 2 == 0)
                    {
                        dummyi = values[(values.Length / 2) - 1];
                        dummyj = values[values.Length / 2];
                        Median = (dummyi + dummyj) / 2;
                    }
                    else
                    {
                        GetMedian_i = (values.Length) / 2;
                        Median = values[GetMedian_i];

                    }


                    minusSquareSummary = 0.0;

                    Count = 0;
                    Low_Count = 0;
                    High_Count = 0;

                    //L_AVG = new double();
                    //H_AVG = new double();

                    foreach (double source in values)
                    {
                        minusSquareSummary += (source - average) * (source - average);

                        //if (Count < values.Length / 2)
                        //{
                        //    L_AVG += source;
                        //    Low_Count++;
                        //}
                        //else
                        //{
                        //    H_AVG += source;
                        //    High_Count++;
                        //}
                        Count++;
                    }

                    //L_AVG = L_AVG / Low_Count;
                    //H_AVG = H_AVG / High_Count;

                    //L_minusSquareSummary = 0f;
                    //H_minusSquareSummary = 0f;

                    d = 0;


                    //for (d = 0; d < Low_Count; d++)
                    //{
                    //    L_minusSquareSummary += (values[d] - L_AVG) * (values[d] - L_AVG);
                    //}

                    //for (d = Low_Count; d < values.Length; d++)
                    //{
                    //    H_minusSquareSummary += (values[d] - H_AVG) * (values[d] - H_AVG);
                    //}


                    stdev = Math.Sqrt(minusSquareSummary / (values.Length - 1));

                    //L_stdev = Math.Sqrt(L_minusSquareSummary / (Low_Count - 1));
                    //H_stdev = Math.Sqrt(H_minusSquareSummary / (High_Count - 1));


                    for (int i = 0; i < Cal_Value_by_rowsdata[Data.Reference_Header[0]].CPK.Length; i++)
                    {

                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].Std[i] = stdev;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].Median_Data[i] = Median;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].Min_Data[i] = values.Min();
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].Max_Data[i] = values.Max();
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].Avg[i] = values.Average();

                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].L_Avg[i] = 0f;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].H_Avg[i] = 0f;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].L_Std[i] = 0f;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count + 1 + Offset]].H_Std[i] = 0f;
                    }


                    Para_Count++;

                }


                Ds.Clear();
                double dummytesttime3323 = TestTime1.Elapsed.TotalMilliseconds;


                #region
                //for (int i = 0; i < ReturnValue.Length; i++)
                //{
                //    average = ReturnValue[i].Average();
                //    Median = 0f;

                //    Array.Sort(ReturnValue[i]);

                //    if (ReturnValue[i].Length % 2 == 0)
                //    {
                //        dummyi = ReturnValue[i][(ReturnValue[i].Length / 2) - 1];
                //        dummyj = ReturnValue[i][ReturnValue[i].Length / 2];
                //        Median = (dummyi + dummyj) / 2;
                //    }
                //    else
                //    {
                //        GetMedian_i = (ReturnValue[i].Length) / 2;
                //        Median = ReturnValue[i][GetMedian_i];

                //    }

                //    minusSquareSummary = 0.0;

                //    Count = 0;
                //    Low_Count = 0;
                //    High_Count = 0;

                //    L_AVG = new double();
                //    H_AVG = new double();

                //    foreach (double source in ReturnValue[i])
                //    {
                //        minusSquareSummary += (source - average) * (source - average);

                //        if (Count < ReturnValue[i].Length / 2)
                //        {
                //            L_AVG += source;
                //            Low_Count++;
                //        }
                //        else
                //        {
                //            H_AVG += source;
                //            High_Count++;
                //        }
                //        Count++;
                //    }

                //    L_AVG = L_AVG / Low_Count;
                //    H_AVG = H_AVG / High_Count;

                //    L_minusSquareSummary = 0f;
                //    H_minusSquareSummary = 0f;

                //    d = 0;

                //    for (d = 0; d < Low_Count; d++)
                //    {
                //        L_minusSquareSummary += (ReturnValue[i][d] - L_AVG) * (ReturnValue[i][d] - L_AVG);
                //    }

                //    for (d = Low_Count; d < ReturnValue[i].Length; d++)
                //    {
                //        H_minusSquareSummary += (ReturnValue[i][d] - H_AVG) * (ReturnValue[i][d] - H_AVG);
                //    }


                //    stdev = Math.Sqrt(minusSquareSummary / (ReturnValue[i].Length - 1));

                //    L_stdev = Math.Sqrt(L_minusSquareSummary / (Low_Count - 1));
                //    H_stdev = Math.Sqrt(H_minusSquareSummary / (High_Count - 1));

                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Std = stdev;
                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Median = Median;
                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Min = ReturnValue[i].Min();
                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Max = ReturnValue[i].Max();
                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Avg = ReturnValue[i].Average();

                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Avg = L_AVG;
                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Avg = H_AVG;
                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Std = L_stdev;
                //    Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Std = H_stdev;

                //    Para_Count++;

                //}
                #endregion


                double dummytesttime3 = TestTime1.Elapsed.TotalMilliseconds;

            }
            public double STDEVandMedian(double[] Ds, int DB, int RowCount)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                //  double[][] ReturnValue = new double[Data.Per_DB_Column_Count[DB]][];

                //for (int i = 0; i < Data.Per_DB_Column_Count[DB]; i++)
                //{
                //    ReturnValue[i] = new double[RowCount];
                //}
                //  double dummytesttime1 = TestTime1.Elapsed.TotalMilliseconds;


                double L_AVG = 0f;
                double H_AVG = 0f;


                double average = Ds.Average();
                double Median = 0f;

                List<double> Ds_Data = Ds.ToList();

                Ds_Data.Sort();
                // Array.Sort(Ds_Data);

                if (Ds_Data.Count % 2 == 0)
                {
                    double dummyi = (Ds_Data.Count / 2) - 1;
                    double dummyj = Ds_Data.Count / 2;
                    Median = (dummyi + dummyj) / 2;
                }
                else
                {
                    int GetMedian_i = (Ds_Data.Count) / 2;
                    Median = Ds_Data[GetMedian_i];

                }

                double minusSquareSummary = 0.0;

                int Count = 0;
                int Low_Count = 0;
                int High_Count = 0;

                L_AVG = new double();
                H_AVG = new double();

                foreach (double source in Ds)
                {
                    minusSquareSummary += (source - average) * (source - average);

                    if (Count < Ds_Data.Count / 2)
                    {
                        L_AVG += source;
                        Low_Count++;
                    }
                    else
                    {
                        H_AVG += source;
                        High_Count++;
                    }
                    Count++;
                }

                L_AVG = L_AVG / Low_Count;
                H_AVG = H_AVG / High_Count;

                double L_minusSquareSummary = 0f;
                double H_minusSquareSummary = 0f;

                int d = 0;

                for (d = 0; d < Low_Count; d++)
                {
                    L_minusSquareSummary += (Ds_Data[d] - L_AVG) * (Ds_Data[d] - L_AVG);
                }

                for (d = Low_Count; d < Ds_Data.Count; d++)
                {
                    H_minusSquareSummary += (Ds_Data[d] - H_AVG) * (Ds_Data[d] - H_AVG);
                }


                double stdev = Math.Sqrt(minusSquareSummary / (Ds_Data.Count - 1));

                double L_stdev = Math.Sqrt(L_minusSquareSummary / (Low_Count - 1));
                double H_stdev = Math.Sqrt(H_minusSquareSummary / (High_Count - 1));

                double dummytesttime3 = TestTime1.Elapsed.TotalMilliseconds;

                return stdev;



            }
            public void STDEVandMedian_For_New_Spec(List<double[]> Ds, int DB, int RowCount)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();



                double dummytesttime2 = TestTime1.Elapsed.TotalMilliseconds;
                int Para_Count = 0;

                double L_AVG = 0f;
                double H_AVG = 0f;
                double average = 0f;
                double Median = 0f;
                double dummyi = 0f;
                double dummyj = 0f;
                int GetMedian_i = 0;
                double minusSquareSummary = 0.0;

                int Count = 0;
                int Low_Count = 0;
                int High_Count = 0;

                double L_minusSquareSummary = 0f;
                double H_minusSquareSummary = 0f;

                int d = 0;


                double stdev = 0f;

                double L_stdev = 0f;
                double H_stdev = 0f;


                double[] values = new double[Ds.Count];

                int Q1_index = 0;
                int Q3_index = 0;

                double LowQ = 0f;
                double HowQ = 0f;

                double IQR = 0f;

                double Lout = 0f;
                double Hout = 0f;

                string[] SN = new string[Ds.Count];

                for (int z = 0; z < Ds[0].Length - 7; z++)
                {
                    for (int w = 0; w < Ds.Count; w++)
                    {
                        values[w] = Ds[w][z];
                    }

                    SN = new string[Ds.Count];

                    average = values.Average();
                    Array.Sort(values);

                    if (values.Length % 2 == 0)
                    {
                        dummyi = values[(values.Length / 2) - 1];
                        dummyj = values[values.Length / 2];
                        Median = (dummyi + dummyj) / 2;

                        GetMedian_i = (values.Length) / 2;
                    }
                    else
                    {
                        GetMedian_i = (values.Length) / 2;
                        Median = values[GetMedian_i];

                    }


                    minusSquareSummary = 0.0;

                    Count = 0;
                    Low_Count = 0;
                    High_Count = 0;

                    L_AVG = new double();
                    H_AVG = new double();

                    foreach (double source in values)
                    {
                        minusSquareSummary += (source - average) * (source - average);

                        if (Count < values.Length / 2)
                        {
                            L_AVG += source;
                            Low_Count++;
                        }
                        else
                        {
                            H_AVG += source;
                            High_Count++;
                        }
                        Count++;
                    }

                    L_AVG = L_AVG / Low_Count;
                    H_AVG = H_AVG / High_Count;

                    L_minusSquareSummary = 0f;
                    H_minusSquareSummary = 0f;

                    d = 0;


                    for (d = 0; d < Low_Count; d++)
                    {
                        L_minusSquareSummary += (values[d] - L_AVG) * (values[d] - L_AVG);
                    }

                    for (d = Low_Count; d < values.Length; d++)
                    {
                        H_minusSquareSummary += (values[d] - H_AVG) * (values[d] - H_AVG);
                    }


                    Q1_index = GetMedian_i / 2;
                    Q3_index = (values.Length - GetMedian_i) / 2 + GetMedian_i;

                    LowQ = values[Q1_index];
                    HowQ = values[Q3_index];

                    IQR = HowQ - LowQ;

                    Lout = LowQ - IQR * DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR;
                    Hout = HowQ + IQR * DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR;

                    stdev = Math.Sqrt(minusSquareSummary / (values.Length - 1));

                    L_stdev = Math.Sqrt(L_minusSquareSummary / (Low_Count - 1));
                    H_stdev = Math.Sqrt(H_minusSquareSummary / (High_Count - 1));

                    //for (int i = 0; i < For_New_Spec_Cal_Value_by_rowsdata[0].CPK.Length; i++)
                    //{
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Std = stdev;
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Median_Data = Median;
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Min_Data = values.Min();
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Max_Data = values.Max();
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Avg = values.Average();

                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Avg = L_AVG;
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Avg = H_AVG;
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Std = L_stdev;
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Std = H_stdev;

                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR_Value = Lout;
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR_Value = Hout;


                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR_Value_Array.Add(1.5, Lout);
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR_Value_Array.Add(-999, -999);

                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR_Value_Array.Add(1.5, Hout);
                    //    //For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR_Value_Array.Add(999, 999);

                    //}
                    int Sn_Count = 0;



                    for (int zz = 0; zz < values.Length; zz++)
                    {
                        if (Hout < Ds[zz][z] || Lout > values[zz])
                        {
                            SN[Sn_Count] = Convert.ToString(zz + 1); Sn_Count++;
                        }
                    }

                    Array.Resize(ref SN, Sn_Count);

                    DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].SN = SN;
                    Para_Count++;

                }
                Para_Count = 0;



                for (int z = 0; z < Ds[0].Length - 7; z++)
                {
                    int dummy_Count = 0;
                    values = new double[Ds.Count];
                    for (int w = 0; w < values.Length; w++)
                    {
                        if (DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].SN.Length == 0)
                        {
                            values[w] = Ds[w][z];
                            dummy_Count++;
                        }
                        else
                        {

                            if (!DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].SN.Contains(Convert.ToString(w + 1)))
                            {
                                values[dummy_Count] = Ds[w][z];
                                dummy_Count++;
                            }

                        }

                    }

                    Array.Resize(ref values, dummy_Count);
                    SN = new string[Ds.Count];

                    average = values.Average();
                    Array.Sort(values);

                    if (values.Length % 2 == 0)
                    {
                        dummyi = values[(values.Length / 2) - 1];
                        dummyj = values[values.Length / 2];
                        Median = (dummyi + dummyj) / 2;

                        GetMedian_i = (values.Length) / 2;
                    }
                    else
                    {
                        GetMedian_i = (values.Length) / 2;
                        Median = values[GetMedian_i];

                    }


                    minusSquareSummary = 0.0;

                    Count = 0;
                    Low_Count = 0;
                    High_Count = 0;

                    L_AVG = new double();
                    H_AVG = new double();

                    foreach (double source in values)
                    {
                        minusSquareSummary += (source - average) * (source - average);

                        if (Count < values.Length / 2)
                        {
                            L_AVG += source;
                            Low_Count++;
                        }
                        else
                        {
                            H_AVG += source;
                            High_Count++;
                        }
                        Count++;
                    }

                    L_AVG = L_AVG / Low_Count;
                    H_AVG = H_AVG / High_Count;

                    L_minusSquareSummary = 0f;
                    H_minusSquareSummary = 0f;

                    d = 0;


                    for (d = 0; d < Low_Count; d++)
                    {
                        L_minusSquareSummary += (values[d] - L_AVG) * (values[d] - L_AVG);
                    }

                    for (d = Low_Count; d < values.Length; d++)
                    {
                        H_minusSquareSummary += (values[d] - H_AVG) * (values[d] - H_AVG);
                    }


                    Q1_index = GetMedian_i / 2;
                    Q3_index = (values.Length - GetMedian_i) / 2 + GetMedian_i;

                    LowQ = values[Q1_index];
                    HowQ = values[Q3_index];

                    IQR = HowQ - LowQ;

                    Lout = LowQ - IQR * DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR;
                    Hout = HowQ + IQR * DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR;

                    stdev = Math.Sqrt(minusSquareSummary / (values.Length - 1));

                    L_stdev = Math.Sqrt(L_minusSquareSummary / (Low_Count - 1));
                    H_stdev = Math.Sqrt(H_minusSquareSummary / (High_Count - 1));

                    for (int i = 0; i < For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].CPK.Length; i++)
                    {
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Std[i] = stdev;
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Median_Data[i] = Median;
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Min_Data[i] = values.Min();
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Max_Data[i] = values.Max();
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Avg[i] = values.Average();

                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Avg[i] = 0f;
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Avg[i] = 0f;
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Std[i] = 0f;
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Std[i] = 0f;

                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR_Value[i] = Lout;
                        For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR_Value[i] = Hout;


                        //   For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR_Value_Array.Add(1.5, Lout);
                        //    For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_IQR_Value_Array.Add(-999, -999);
                        //
                        //   For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR_Value_Array.Add(1.5, Hout);
                        //   For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_IQR_Value_Array.Add(999, 999);

                    }
                    //int Sn_Count = 0;



                    //for (int zz = 0; zz < values.Length; zz++)
                    //{
                    //    if (Hout < Ds[zz][z] || Lout > values[zz])
                    //    {
                    //        SN[Sn_Count] = Convert.ToString(zz + 1); Sn_Count++;
                    //    }
                    //}

                    //Array.Resize(ref SN, Sn_Count);

                    //DIC_IQR[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].SN = SN;
                    Para_Count++;

                }


                double dummytesttime3323 = TestTime1.Elapsed.TotalMilliseconds;


            }
            static double[] STDEVandMedian(DataSet ds)
            {
                List<double> DataSet_Values = new List<double>();
                double[] ReturnValue = new double[2];

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    DataSet_Values.Add(Convert.ToDouble(dr.ItemArray[0]));
                }

                double average = DataSet_Values.Average();
                double Median = 0f;

                if (DataSet_Values.Count % 2 == 0)
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = DataSet_Values.Count / 2;
                    Median = DataSet_Values[GetMedian_i];
                }
                else
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = (DataSet_Values.Count + 1) / 2;
                    Median = DataSet_Values[GetMedian_i];
                }

                double minusSquareSummary = 0.0;

                foreach (double source in DataSet_Values)
                {
                    minusSquareSummary += (source - average) * (source - average);
                }

                double stdev = Math.Sqrt(minusSquareSummary / (DataSet_Values.Count - 1));

                ReturnValue[0] = stdev; ReturnValue[1] = Median;

                return ReturnValue;
            }
            public string Get_Data_From_Table(string Table, string header)
            {

                stringA[0].Clear();
                stringA[0].Append("select " + header + " from " + Table);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }
                stringA[0].Clear();

                return Convert.ToString(Value[0]);
            }


        }

        public class BOXPLOT : INT
        {
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ReaderWriterLockSlim[] sqlitelock { get; set; }
            public string[] strConn { get; set; }
            public SQLiteConnection[] conn { get; set; }
            public SQLiteCommand[] cmd { get; set; }

            public SQLiteDataAdapter[] sqlAdapter { get; set; }
            public SQLiteCommandBuilder[] sqlcmdbuilder { get; set; }
            public SQLiteDataReader[] SqReader { get; set; }

            public DbDataReader[] DbReader { get; set; }
            public DataSet[] ds { get; set; }
            public DataTable dt_test { get; set; }
            public DataTable[] dt { get; set; }
            public SQLiteTransaction[] tran { get; set; }

            public ManualResetEvent[] ThreadFlags { get; set; }
            public ManualResetEvent[] Insert_ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public int Limit { get; set; }
            public int Limit_Count { get; set; }
            public int Table_Count { get; set; }
            public bool[] Insert_Thread_Wait { get; set; }
            public double[] Testtime { get; set; }

            public double[][] test { get; set; }
            public string Filename { get; set; }
            double[] Testtime1 { get; set; }
            double[] Testtime2 { get; set; }
            double[] Testtime3 { get; set; }
            public string[][] Teststring { get; set; }
            public double[][] Testdouble { get; set; }

            public object[] ID { get; set; }
            public object[] Value { get; set; }
            public object[] WAFER_ID { get; set; }
            public object[] LOT_ID { get; set; }
            public object[] SITE_ID { get; set; }
            public Dictionary<string, double[]> Selected_Parameter_Distribution { get; set; }

            public object[] Variation { get; set; }
            public Dictionary<string, IQR> DIC_IQR { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test_New_Spec { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent { get; set; }
            public List<List<int>>[] For_Any_Yield { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_Lot { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_SITE { get; set; }
            public List<List<int>[]>[] ForCampare_Yield { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent_For_New_Spec { get; set; }
            public List<List<int>>[] For_Any_Yield_For_New_Spec { get; set; }
            public List<List<int>[]>[] For_New_Spec_ForCampare_Yield { get; set; }
            public List<int[]>[] ForCampare_Yield_Fro_DB { get; set; }
            public List<List<int[]>>[] ForCampare_Yield_Fro_DB_List { get; set; }
            public List<List<List<List<int>[]>>>[] ForCampare_Yield_DB_LotVariation { get; set; }
            public List<List<List<int[]>>>[] ForCampare_Yield_Fro_DB_List_LotVariation { get; set; }
            public Dictionary<string, int> Refer_Site_And_Num { get; set; }
            public Dictionary<string, int> Refer_Lot_And_Num { get; set; }
            public List<int>[] ForCampare_Yield_List { get; set; }
            public List<List<int>[]> ForCampare_Yield_List1 { get; set; }
            public List<List<int>>[] For_New_Spec_ForCampare_Yield2 { get; set; }
            public List<List<int>[]>[] ForCampare_Yield_List2 { get; set; }
            public Dictionary<string, Values> Values { get; set; }
            public Dictionary<string, Data_Calculation> Cal_Value_by_rowsdata { get; set; }
            public Dictionary<string, Data_Calculation> For_New_Spec_Cal_Value_by_rowsdata { get; set; }
            public List<int>[] Check { get; set; }
            public List<List<int>[]> Test { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public List<double[]>[] DB_DataSet_Values { get; set; }
            public Dictionary<string, CSV_Class.For_Box> Dic_Test_For_Spec_Gen { get; set; }
            public Dictionary<string, CSV_Class.For_Box>[] Dic_Test { get; set; }
            public Dictionary<string, int> Lot_Dic { get; set; }
            public Dictionary<string, int> Site_Dic { get; set; }
            public Dictionary<string, int> Bin_Dic { get; set; }
            public Dictionary<string, Dictionary<string, List<string>>> Matching_Lots { get; set; }
            public Dictionary<string, List<string>> Matching_Lot { get; set; }
            public Stopwatch[] TestTime1 { get; set; }
            public Stopwatch[] TestTime2 { get; set; }
            public Stopwatch[] TestTime3 { get; set; }
            public Stopwatch[] TestTime4 { get; set; }
            public Stopwatch[] TestTime5 { get; set; }
            public long SampleCount { get; set; }
            public object Update_Data_ID { get; set; }
            public string[] Update_Datas_ID { get; set; }
            public string Get_Gross_Para { get; set; }

            public double Get_Gross_Persent { get; set; }
            public string Get_Gross_Selector { get; set; }
            public List<Dictionary<string, Gross>[]> List_Gross_Values { get; set; }
            public Dictionary<string, Gross>[] Gross_Values1 { get; set; }
            public object[] Std_Value { get; set; }
            public double[] Std_Value_Convert { get; set; }
            public long NB { get; set; }
            public string Table { get; set; }

            public double[] Make_New_Spec_For_Yield_Min { get; set; }
            public double[] Make_New_Spec_For_Yield_Max { get; set; }
            public List<string> Gross { get; set; }
            public List<string[]>[] DataSet_Value { get; set; }
            public List<double[]>[] DataSet_Double_Value { get; set; }

            public int[] Each_Thread_Count { get; set; }

            public string Lot_ID { get; set; }
            public string SubLot_ID { get; set; }
            public string Tester_ID { get; set; }
            public string Site { get; set; }
            public string Bin { get; set; }
            public string ID_Unit { get; set; }
            public int Bin_place { get; set; }
            public string Query { get; set; }
            public bool _From_Db { get; set; }
            public int Spec_Table_Count { get; set; }
            public bool _Flag { get; set; }
            public bool _SUBLOT_Flag { get; set; }
            public bool Clotho_Spec_Flag { get; set; }
            public string Before_Lot_ID { get; set; }
            public string Changed_Lot_ID { get; set; }


            public string[] No_Index { get; set; }
            public string[] Paraname { get; set; }
            public string[] SpecMin { get; set; }
            public string[] SpecMax { get; set; }
            public string[] DataMin { get; set; }
            public string[] DataMedian { get; set; }
            public string[] DataMax { get; set; }
            public string[] CPK { get; set; }
            public string[] STD { get; set; }
            public string[] Percent { get; set; }
            public string[] Fail { get; set; }
            public int Count_Current_Setting { get; set; }

            public string[] Line { get; set; }

            public void Open_DB(string FileName, Data_Class.Data_Editing.INT Data_Edit)
            {
                string Filename = FileName.Substring(FileName.LastIndexOf("\\") + 1);
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    cmd[i].ExecuteNonQuery();

                }


                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                dt = new DataTable[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    dt[i] = new DataTable();
                    cmd[i].CommandText = "PRAGMA synchronous";
                    SqReader[i] = cmd[i].ExecuteReader();
                    dt[i].Load(SqReader[i]);
                }

                ForCampare_Yield_List = new List<int>[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data_Edit.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }
                ForCampare_Yield_List1 = new List<List<int>[]>();

            }

            public void Open_DB(string[] FileName, Data_Class.Data_Editing.INT Data_Edit)
            {

                Data_Edit.DB_Count = FileName.Length;
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    string Filename = FileName[i].Substring(FileName[i].LastIndexOf("\\") + 1);

                    int length = Filename.Length;
                    Filename = Filename.Substring(0, length - 4);

                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    //cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    //cmd[i].ExecuteNonQuery();

                }


                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                dt = new DataTable[Data_Edit.DB_Count];

                ForCampare_Yield_List = new List<int>[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data_Edit.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }
                ForCampare_Yield_List1 = new List<List<int>[]>();

            }

            public void DropTable(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {
                try
                {
                    for (int i = 0; i < Data_Edit.DB_Count; i++)
                    {
                        cmd[i].CommandText = "";
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                    }
                }
                catch { }

            }

            public void Insert_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(MakecolumnsThread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void MakecolumnsThread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY, Fail VARCHAR(5) , LOT_ID VARCHAR(5) , SUBLOT_ID VARCHAR(5) , BIN VARCHAR(5));");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void Insert_Current_Setting(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void Insert_Current_Setting_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {
                Data = Data_Edit;
                this.Table = Table;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  cmd[i].Reset();
                    //    ThreadFlags[i] = new ManualResetEvent(false);
                    Insert_Current_Setting_Data_Thread(i);
                    //  ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //       Wait[i] = ThreadFlags[i].WaitOne();
                }

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Clotho_Spec_Max_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}

            }

            public void Insert_Current_Setting_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();
                stringA[i].Clear();
                SampleCount = 1;

                cmd[i] = new SQLiteCommand(conn[i]);

                int Count = Data.Per_DB_Column_Count[i];


                int k = 0;


                if (Table.ToUpper() == "CLOTHO_SPEC")
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Min[Spec_Count] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Min[Spec_Count] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                    }




                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }
                else
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[0] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Min[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;

                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }
                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }




                //   ThreadFlags[i].Set();
            }
            public void Insert_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE IF Not Exists spec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE IF Not Exists spec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY , Fail VARCHAR(5), LOT_ID VARCHAR(5) , SUBLOT_ID VARCHAR(5) , BIN VARCHAR(5));");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {

            }

            public void Insert_New_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_New_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void Insert_New_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE IF Not Exists newspec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE IF Not Exists newspec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY , Fail VARCHAR(5), LOT_ID VARCHAR(5) , SUBLOT_ID VARCHAR(5) , BIN VARCHAR(5));");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                // sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];
                sqlitelock = new ReaderWriterLockSlim[Data.DB_Count];
                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                //Testdouble = new double[7][];

                //Testdouble[0] = new double[Data.DB_Column_Limit];
                //Testdouble[1] = new double[Data.DB_Column_Limit];
                //Testdouble[2] = new double[Data.DB_Column_Limit];
                //Testdouble[3] = new double[Data.DB_Column_Limit];
                //Testdouble[4] = new double[Data.DB_Column_Limit];
                //Testdouble[5] = new double[Data.DB_Column_Limit];
                //Testdouble[6] = new double[Data.Per_DB_Column_Count[6]];
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder(100000);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder(100000);
                    Testtime[i] = TestTime1.Elapsed.TotalMilliseconds;
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);
            }
            public void Insert_Ref_Header_Data(Data_Class.Data_Editing.INT Data_Edit)
            {


            }
            public void Insert_Data(long Sample)
            {
                SampleCount = Sample;

                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Insert_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO data VALUES ('" + Data.Getstring[0].Replace("PID-", "") + "',");
                    // Testdouble[i][0] = Data.New_Data[0];
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    stringA[i].Append("INSERT INTO data VALUES ('" + Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count] + "',");
                    // Testdouble[i][0] = Data.New_Data[Data.DB_Column_Limit * i];
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {

                    stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k] + "',");
                    // Testdouble[i][j] = Data.New_Data[Data.DB_Column_Limit * i + j];
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }
                //  stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k] + "');");
                stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k] + "', '" + SampleCount + "' , '0' , '" + Lot_ID + "' , '" + SubLot_ID + "' , '" + Bin + "');");
                //  Testdouble[i][j] = Data.New_Data[Data.DB_Column_Limit * i + j];

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(Data.Getstring[(Data.DB_Column_Limit * i) + TheFirst_Trashes_Header_Count + k]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Data_Get_From_DB(int Sample)
            {
                SampleCount = Sample;

                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Get_From_DB_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }

            public void Insert_Data_Get_From_DB_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Get_From_DB(Data_Class.Data_Editing.INT Data_Edit)
            {


                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Get_From_DB_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }

            public void Insert_Spec_Get_From_DB_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Insert_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;
                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[0] + "',");

                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[Data.DB_Column_Limit * i] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_LowSpec[Data.DB_Column_Limit * i + k] + "', '" + SampleCount + "' ,0,0,0,0);");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                Thread.Sleep(100);
                stringA[i].Clear();
                cmd[i].Reset();
                k = 0;
                SampleCount = 2;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[0] + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[Data.DB_Column_Limit * i] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_HighSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_HighSpec[Data.DB_Column_Limit * i + k] + "', '" + SampleCount + "' ,0,0,0,0);");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }
            public void Insert_Files_Name(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Make_table(string Tablename)
            {

            }
            public void Make_table2(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {

            }
            public void Make_table_For_Filename(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {
                Data = Data_Edit;
                Table = Tablename;

                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(_Make_Table_For_Filename), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void _Make_Table_For_Filename(Object threadContext)
            {
                int i = (int)threadContext;

                stringA[i].Append("CREATE TABLE " + Table + "(FIle VARCAHR(20))");


                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].CommandText = "";

                stringA[i].Append(",");

                ThreadFlags[i].Set();
            }

            public void Make_table_For_Trace(string Tablename, string Chan, bool Flag)
            {
                stringA[0].Clear();
                stringA[0].Append("CREATE TABLE " + Tablename + "( FIRST VARCAHR(5), END VARCAHR(5), DBCOUNT VARCHAR(5), COLUMNCOUNT VARCHAR(5) );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";

                stringA[0].Clear();
                stringA[0].Append("INSERT INTO INF VALUES ('" + TheFirst_Trashes_Header_Count + "' , '" + TheEnd_Trashes_Header_Count + "' , '" + Data.Per_DB_Column_Count.Length + "' , '" + Data.Per_DB_Column_Count[Data.Per_DB_Column_Count.Length - 1] + "' );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";
            }
            public void Delete_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Delete_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Delete_Lot_Data(string Query)
            {

            }
            public void Delete_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;



                stringA[i].Append("Delete from " + Table + " where id = 1");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                stringA[i].Clear();

                stringA[i].Append("Delete from " + Table + " where id = 2");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Save_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                //Update_Data_ID = data;

                //if (data != null)
                //{
                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        ThreadFlags[i] = new ManualResetEvent(false);
                //        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                //    }

                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        Wait[i] = ThreadFlags[i].WaitOne();
                //    }
                //}

            }
            public void Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    //  Insert_table_Data_Thread(i);
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_table_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}


            }

            public void Road_Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                //  SampleCount = Sample;

                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Get_From_DB_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Road_Save_Customer_Spec_table_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void LOTID_Update(string Query, string Query2, string CellID)
            {

            }

            public void Gross_Update_Data(object data)
            {
                Update_Data_ID = data;

                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }

            }
            public void Gross_Update_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (object o in (Array)Update_Data_ID)
                {
                    cmd[i].CommandText = "Update data set FAIL = '1'  where id = " + o.ToString();
                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Gross_Update_Datas(List<string> data)
            {
                Update_Datas_ID = data.ToArray();
                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Datas_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }
            }
            public void Gross_Update_Datas_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (object o in (Array)Update_Datas_ID)
                {
                    cmd[i].CommandText = "Update data set FAIL = '1'  where id = '" + o.ToString() + "'";
                    cmd[i].ExecuteNonQuery();
                    //      cmd[i].Reset();
                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Chnaged_Spec_Update_Data(int DB, int Index, string Parameter, double Spec, int GetId)
            {
                stringA[DB].Clear();
                stringA[DB].Append("Update newspec set " + Parameter + " = " + Spec + " where id = " + GetId);

                cmd[DB].CommandText = stringA[DB].ToString();

                cmd[DB].ExecuteNonQuery();
                cmd[DB].Reset();

                stringA[DB].Clear();
            }
            public Dictionary<string, double[]> Chnaged_Spec_Anl_Yield(int DB, int Index, string Parameter)
            {
                stringA[DB].Clear();
                Dictionary<string, double[]> Dic_Change_Spec = new Dictionary<string, double[]>();


                stringA[DB].Append("Select " + Parameter + " from newspec");

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                object[] GetSpec = new object[ds[DB].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    GetSpec[count] = dr.ItemArray[0];
                    count++;
                }

                double[] Toduble_Spec = Array.ConvertAll<object, double>(GetSpec, Convert.ToDouble);

                Dic_Change_Spec.Add("SPEC", Toduble_Spec);
                stringA[DB].Clear();

                stringA[DB].Append("Select " + Parameter + " from data where Fail not like '1'");

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                object[] GetData = new object[ds[DB].Tables[0].Rows.Count];
                count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    GetData[count] = dr.ItemArray[0];
                    count++;
                }

                double[] Toduble_Data = Array.ConvertAll<object, double>(GetData, Convert.ToDouble);

                Dic_Change_Spec.Add("DATA", Toduble_Data);

                stringA[DB].Clear();

                return Dic_Change_Spec;
            }
            public void Get_Ave_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                }

            }
            public void Get_Ave_Data_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                }

            }
            public void Set_Refer_for_Anlyzer(Data_Class.Data_Editing.INT Data_Edit)
            {
                stringA[0].Clear();
                stringA[0].Append("Select id from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                ForCampare_Yield_List1 = new List<List<int>[]>();
                for (int k = 0; k < Value.Length; k++)
                {
                    ForCampare_Yield_List = new List<int>[Data.DB_Count];

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ForCampare_Yield_List[i] = new List<int>();
                    }

                    for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                    {
                        for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                        {
                            ForCampare_Yield_List[i].Add(0);
                        }
                    }

                    ForCampare_Yield_List1.Add(ForCampare_Yield_List);
                }
            }
            public void Get_Ave_Data2(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    Get_Ave_Data_Thread2(i);
                }
            }

            public void Get_Ave_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from data where Fail not like '1'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                List<double[]> DataSet_Values = new List<double[]>();
                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    values[Data.Per_DB_Column_Count[i] + 2] = 0;
                    values[Data.Per_DB_Column_Count[i] + 3] = 0;
                    double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);
                    DataSet_Values.Add(doubles);

                    count++;

                }
                SqReader[i].Close();

                double testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                STDEVandMedian(DataSet_Values, i, count);

                double testtime2 = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();
            }

            public void Get_Ave_Data_Thread2(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from data where Fail not like '1'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                List<double[]> DataSet_Values = new List<double[]>();
                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    values[Data.Per_DB_Column_Count[i] + 2] = 0;
                    values[Data.Per_DB_Column_Count[i] + 3] = 0;
                    double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);
                    DataSet_Values.Add(doubles);

                    count++;

                }
                SqReader[i].Close();

                STDEVandMedian(DataSet_Values, i, count);

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();
            }

            public void Get_Saved_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                DataSet_Value = new List<string[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    DataSet_Value[i] = new List<string[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Saved_Spec_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }

            public void Get_Saved_Spec_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from newspec");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    DataSet_Value[i].Add(stringD);

                    count++;

                }
                SqReader[i].Close();

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();

            }
            public void Get_Rows_Data(Data_Class.Data_Editing.INT Data_Edit)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                DataSet_Value = new List<string[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    DataSet_Value[i] = new List<string[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Rows_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
            }

            public void Get_Rows_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from data where id = '" + Data.Set_ID + "'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    DataSet_Value[i].Add(stringD);
                    count++;
                    break;
                }


                SqReader[i].Close();

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();

            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface)
            {
                //stringA[DB].Clear();
                //stringA[DB].Append("Select id, " + Select_Para + " from data where fail not like '1'");

                //cmd[DB].CommandText = stringA[DB].ToString();
                //ds[DB] = new DataSet();

                //sqlAdapter[DB].SelectCommand = cmd[DB];
                //sqlAdapter[DB].Fill(ds[DB]);

                //ID = new object[ds[DB].Tables[0].Rows.Count];
                //Value = new object[ds[DB].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //{
                //    ID[count] = dr.ItemArray[0];
                //    Value[count] = dr.ItemArray[1];

                //    count++;
                //}

                //stringA[DB].Clear();
            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface, DataTable dt)
            {
                //stringA[DB].Clear();
                //stringA[DB].Append("Select id, " + Select_Para + " from data");

                //cmd[DB].CommandText = stringA[DB].ToString();
                //ds[DB] = new DataSet();

                //sqlAdapter[DB].SelectCommand = cmd[DB];
                //sqlAdapter[DB].Fill(ds[DB]);

                //ID = new object[ds[DB].Tables[0].Rows.Count];
                //Value = new object[ds[DB].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //{
                //    ID[count] = dr.ItemArray[0];
                //    Value[count] = dr.ItemArray[1];

                //    count++;
                //}

                //double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);


                //stringA[DB].Clear();
            }

            public void Get_Selected_Para(int DB, string Select_Para, bool Flag, string Selector)
            {
                stringA[DB].Clear();
                stringA[DB].Append("Select id, " + Select_Para + ",LOT_ID from data");

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                ID = new object[ds[DB].Tables[0].Rows.Count];
                Value = new object[ds[DB].Tables[0].Rows.Count];
                Variation = new object[ds[DB].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    ID[count] = dr.ItemArray[0];
                    Value[count] = dr.ItemArray[1];
                    Variation[count] = dr.ItemArray[2];
                    count++;
                }

                stringA[DB].Clear();

            }
            public double[] Get_Find_Bin(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);

                stringA[0].Clear();
                return doubles;
            }
            public List<object[]> Get_Data_By_Querys(string Query)
            {
                return null;
            }
            public Dictionary<string, object[]> Get_Data_By_Query_S4PD(string Query, string Chan)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                // cmd[0].CommandText = stringA[0].ToString();
                // ds[0] = new DataSet();

                // sqlAdapter[0].SelectCommand = cmd[0];
                // sqlAdapter[0].Fill(ds[0]);

                // Value = new object[ds[0].Tables[0].Rows.Count];

                //// int count = 0;
                // foreach (DataRow dr in ds[0].Tables[0].Rows)
                // {
                //     Value[count] = dr.ItemArray[0];
                //     count++;
                // }

                //  string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);
                // SqReader[0] = cmd[0].ExecuteReader();
                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[500000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    Value1[count] = stringD[0];

                    count++;

                }

                Array.Resize(ref Value1, count);

                cmd[0].Dispose();
                SqReader[0].Close();

                string[] _string = Array.ConvertAll<object, string>(Value1, Convert.ToString);


                stringA[0].Clear();
                return null;
            }
            public string[] Get_Data_By_Query(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);

                stringA[0].Clear();
                return _string;
            }
            public string[] Get_Data_By_Query(string Query, int DB)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);

                stringA[0].Clear();
                return _string;
            }
            public void Get_Defined_Para(object[,] DummyData, string key, Data_Class.Data_Editing.INT Data_InterFace)
            {


            }

            public void Get_Gross_Check_Para(Data_Class.Data_Editing.INT Data_Edit, string Select_Para, double Persent, string Selector, int SelectedBin)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                Get_Gross_Para = Select_Para;
                Get_Gross_Persent = Persent;
                //   Gross = ForGross_Fail_Unit;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = false;
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Gross_Check_Para_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
                double test = TestTime1.Elapsed.TotalMilliseconds;


                stringA[0].Append("Select id from data where FAIL not like '1'");
                //stringA[0].Append("Select id from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                ID = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    ID[count] = dr.ItemArray[0];
                    count++;
                }

                stringA[0].Clear();
                //    List_Gross_Values.Add(Gross_Values1);
            }

            public void Get_Gross_Check_Para_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                int k = 0;
                for (k = 0; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    string[] Split_Dummy = Data.Reference_Header[Data.DB_Column_Limit * i + k].Split('_');
                    if (Split_Dummy.Length != 1)
                    {
                        if (Split_Dummy[1].ToUpper() == Get_Gross_Para.ToUpper())
                        {
                            ds[i] = new DataSet();
                            stringA[i].Clear();
                            stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data where Fail not like '1'");
                            //  stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data");
                            cmd[i].CommandText = stringA[i].ToString();

                            sqlAdapter[i].SelectCommand = cmd[i];
                            sqlAdapter[i].Fill(ds[i]);

                            object[] DataValue = new object[ds[i].Tables[0].Rows.Count];

                            int count = 0;
                            foreach (DataRow dr in ds[i].Tables[0].Rows)
                            {
                                DataValue[count] = dr.ItemArray[0];
                                count++;
                            }


                            double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

                            double DataMin = doubles.Min();
                            double DataMax = doubles.Max();
                            double DataAve = doubles.Average();

                            double DataMinindex = doubles.ToList().IndexOf(DataMin);
                            double DataMaxindex = doubles.ToList().IndexOf(DataMax);

                            double Divide = DataMax / DataMin;

                            string[] test;
                            string _Substring = Get_Gross_Para.Substring(0, 1);

                            double MinSpec = 0f;
                            bool Define_Flag = false;

                            if (Get_Gross_Para.ToUpper().Contains("IBATT") || Get_Gross_Para.ToUpper().Contains("ICC") || Get_Gross_Para.ToUpper().Contains("IDD"))
                            {
                                Define_Flag = true;
                                test = Convert.ToString(Get_Gross_Persent).Split('.');
                                MinSpec = 1 - (Convert.ToDouble(test[1]) / 10);
                            }
                            else
                            {
                                Define_Flag = false;
                                MinSpec = Convert.ToDouble(Get_Gross_Persent) * -1;
                            }

                            if (Define_Flag)
                            {
                                for (int j = 0; j < doubles.Length; j++)
                                {
                                    if (DataAve / doubles[j] > Get_Gross_Persent || DataAve / doubles[j] < MinSpec)
                                    {
                                        if (!Gross.Contains(Convert.ToString(j + 1)))
                                        {
                                            //         Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                for (int j = 0; j < doubles.Length; j++)
                                {
                                    if (DataAve - doubles[j] > Get_Gross_Persent || doubles[j] - DataAve < MinSpec)
                                    {
                                        if (!Gross.Contains(Convert.ToString(j + 1)))
                                        {
                                            //          Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                                        }
                                    }
                                }
                            }

                            stringA[i].Clear();
                            cmd[i].CommandText = "";
                        }
                        //if (Get_Gross_Para == "POUT" && Split_Dummy.Length > 7 && Split_Dummy[6].ToUpper() == "FIXEDPOUT" && Split_Dummy[1].ToUpper() == "POUT")
                        //{
                        //    ds[i] = new DataSet();
                        //    stringA[i].Clear();

                        //    stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data where Fail not like '%1%'");
                        //    //stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data");
                        //    cmd[i].CommandText = stringA[i].ToString();

                        //    sqlAdapter[i].SelectCommand = cmd[i];
                        //    sqlAdapter[i].Fill(ds[i]);

                        //    object[] DataValue = new object[ds[i].Tables[0].Rows.Count];

                        //    int count = 0;

                        //    foreach (DataRow dr in ds[i].Tables[0].Rows)
                        //    {
                        //        DataValue[count] = dr.ItemArray[0];
                        //        count++;
                        //    }

                        //    string remove = Split_Dummy[7].Replace("dBm", "");

                        //    double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

                        //    double DataMin = Convert.ToDouble(remove) - Get_Gross_Persent;
                        //    double DataMax = Convert.ToDouble(remove) + Get_Gross_Persent;

                        //    for (int j = 0; j < doubles.Length; j++)
                        //    {
                        //        if (doubles[j] < DataMin)
                        //        {
                        //            if (!Gross.Contains(Convert.ToString(j + 1)))
                        //            {
                        //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                        //            }

                        //        }
                        //        else if (doubles[j] > DataMax)
                        //        {
                        //            if (!Gross.Contains(Convert.ToString(j + 1)))
                        //            {
                        //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                        //            }
                        //        }
                        //    }

                        //    stringA[i].Clear();
                        //    cmd[i].CommandText = "";
                        //}

                    }
                }
                ThreadFlags[i].Set();
            }
            public void Get_From_Db_Data_for_Anly(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;

                ForCampare_Yield_Fro_DB = new List<int[]>[Data.DB_Count];
                ForCampare_Yield_Fro_DB_List = new List<List<int[]>>[Data.DB_Count];


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_Fro_DB_List[i] = new List<List<int[]>>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data_for_Anly_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }
            public void Get_From_Db_Data_for_Anly_Thread(Object threadContext)
            {
                int i = (int)threadContext;


                //    int Count_Data = 0;
                int count = 0;


                stringA[i].Append("Select * from data where Fail not like '1'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                count = 0;

                List<double[]> DataSet_Values = new List<double[]>();
                while (SqReader[i].Read())
                {

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);

                    string Lot = values[Data.Per_DB_Column_Count[i] + 2].ToString();

                    values[Data.Per_DB_Column_Count[i] + 2] = 0;
                    values[Data.Per_DB_Column_Count[i] + 3] = 0;


                    double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
                    double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);
                    double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

                    int[] Check = new int[SqReader[i].FieldCount];
                    List<int[]> Test = new List<int[]>();
                    ForCampare_Yield_Fro_DB[i] = new List<int[]>();


                    int j = 0;

                    int Select_Lot = Refer_Lot_And_Num[Lot];

                    if (i == 0)
                    {
                        Check[j] = 0;
                    }
                    else
                    {
                        if (Data.New_HighSpec[Data.DB_Column_Limit * i] < doubles[0] || Data.New_LowSpec[Data.DB_Column_Limit * i] > doubles[0])
                        {
                            Check[j] = 1;
                        }
                    }

                    j = 1;

                    for (j = 1; j < values.Length - 6; j++)
                    {

                        if (Data.New_HighSpec[Data.DB_Column_Limit * i + j] < doubles[j] || Data.New_LowSpec[(Data.DB_Column_Limit * i) + j] > doubles[j])
                        {
                            if (Select_Lot == 1)
                            {

                            }

                            Check[j] = 1;
                        }
                    }

                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + j] < doubles[j] || Data.New_LowSpec[(Data.DB_Column_Limit * i) + j] > doubles[j])
                    {
                        Check[j] = 1;
                    }

                    Test.Add(Check);
                    ForCampare_Yield_Fro_DB[i] = Test;
                    ForCampare_Yield_Fro_DB_List[i].Add(ForCampare_Yield_Fro_DB[i]);



                    ForCampare_Yield_Fro_DB_List_LotVariation[i][Select_Lot].Add(ForCampare_Yield_Fro_DB[i]);
                    count++;

                    double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;

                }
                SqReader[i].Close();

                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();


            }
            public void Get_From_Db_Data_for_Anly_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;

                ForCampare_Yield_Fro_DB = new List<int[]>[Data.DB_Count];
                ForCampare_Yield_Fro_DB_List = new List<List<int[]>>[Data.DB_Count];


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_Fro_DB_List[i] = new List<List<int[]>>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data_for_Anly_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }
            public void Get_From_Db_Ref_Header(Data_Class.Data_Editing.INT Data_Edit)
            {



            }
            public void Get_From_Db_Ref_Header_Thread(Object threadContext)
            {



            }
            public void Get_Current_Setting(Data_Class.Data_Editing.INT Data_Edit, int NB)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Count_Current_Setting = NB;
                this.Data = Data_Edit;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                   // Get_From_Db_Data(i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public int Get_Sample_Count(int DB, string Query)
            {

                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                }
                stringA[0].Clear();

                int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Data_Count[0];

            }

            public int Get_Column_Count(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {
                return 0;
            }

            public void Close(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();
                    conn[i].Close();

                }
            }

            public void Read_Dispose(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }

            public void Set_Conn(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }

            public void trans(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;
                tran = new SQLiteTransaction[Data.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    tran[i] = conn[i].BeginTransaction();
                    cmd[i].Transaction = tran[i];
                }
            }

            public void Commit(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    tran[i].Commit();
                }


            }

            public void STDEVandMedian(List<double[]> Ds, int DB, int RowCount)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                double[][] ReturnValue = new double[Data.Per_DB_Column_Count[DB]][];

                for (int i = 0; i < Data.Per_DB_Column_Count[DB]; i++)
                {
                    ReturnValue[i] = new double[RowCount];
                }
                double dummytesttime1 = TestTime1.Elapsed.TotalMilliseconds;
                int j = 0;
                int k = 0;


                foreach (double[] o in Ds)
                {
                    var t = o;
                    for (int q = 0; q < t.Length - 5; q++)
                    {
                        ReturnValue[j][k] = t[q];
                        j++;
                    }
                    j = 0;
                    k++;
                }
                double dummytesttime2 = TestTime1.Elapsed.TotalMilliseconds;
                int Para_Count = 0;

                double L_AVG = 0f;
                double H_AVG = 0f;

                for (int i = 0; i < ReturnValue.Length; i++)
                {
                    double average = ReturnValue[i].Average();
                    double Median = 0f;

                    Array.Sort(ReturnValue[i]);

                    if (ReturnValue[i].Length % 2 == 0)
                    {
                        double dummyi = ReturnValue[i][(ReturnValue[i].Length / 2) - 1];
                        double dummyj = ReturnValue[i][ReturnValue[i].Length / 2];
                        Median = (dummyi + dummyj) / 2;
                    }
                    else
                    {
                        int GetMedian_i = (ReturnValue[i].Length) / 2;
                        Median = ReturnValue[i][GetMedian_i];

                    }

                    double minusSquareSummary = 0.0;

                    int Count = 0;
                    int Low_Count = 0;
                    int High_Count = 0;

                    L_AVG = new double();
                    H_AVG = new double();
                    if (i == 20)
                    {

                    }
                    foreach (double source in ReturnValue[i])
                    {
                        minusSquareSummary += (source - average) * (source - average);

                        if (Count < ReturnValue[i].Length / 2)
                        {
                            L_AVG += source;
                            Low_Count++;
                        }
                        else
                        {
                            H_AVG += source;
                            High_Count++;
                        }
                        Count++;
                    }

                    L_AVG = L_AVG / Low_Count;
                    H_AVG = H_AVG / High_Count;

                    double L_minusSquareSummary = 0f;
                    double H_minusSquareSummary = 0f;

                    int d = 0;

                    for (d = 0; d < Low_Count; d++)
                    {
                        L_minusSquareSummary += (ReturnValue[i][d] - L_AVG) * (ReturnValue[i][d] - L_AVG);
                    }

                    for (d = Low_Count; d < ReturnValue[i].Length; d++)
                    {
                        H_minusSquareSummary += (ReturnValue[i][d] - H_AVG) * (ReturnValue[i][d] - H_AVG);
                    }


                    double stdev = Math.Sqrt(minusSquareSummary / (ReturnValue[i].Length - 1));

                    double L_stdev = Math.Sqrt(L_minusSquareSummary / (Low_Count - 1));
                    double H_stdev = Math.Sqrt(H_minusSquareSummary / (High_Count - 1));

                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Std = stdev;
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Median_Data = Median;
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Min_Data = ReturnValue[i].Min();
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Max_Data = ReturnValue[i].Max();
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Avg = ReturnValue[i].Average();

                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Avg = L_AVG;
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Avg = H_AVG;
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_Std = L_stdev;
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_Std = H_stdev;

                    Para_Count++;

                }


                double dummytesttime3 = TestTime1.Elapsed.TotalMilliseconds;

            }
            static double[] STDEVandMedian(DataSet ds)
            {
                List<double> DataSet_Values = new List<double>();
                double[] ReturnValue = new double[2];

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    DataSet_Values.Add(Convert.ToDouble(dr.ItemArray[0]));
                }

                double average = DataSet_Values.Average();
                double Median = 0f;

                if (DataSet_Values.Count % 2 == 0)
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = DataSet_Values.Count / 2;
                    Median = DataSet_Values[GetMedian_i];
                }
                else
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = (DataSet_Values.Count + 1) / 2;
                    Median = DataSet_Values[GetMedian_i];
                }

                double minusSquareSummary = 0.0;

                foreach (double source in DataSet_Values)
                {
                    minusSquareSummary += (source - average) * (source - average);
                }

                double stdev = Math.Sqrt(minusSquareSummary / (DataSet_Values.Count - 1));

                ReturnValue[0] = stdev; ReturnValue[1] = Median;

                return ReturnValue;
            }
            public string Get_Data_From_Table(string Table, string header)
            {

                return "";
            }

        }

        public class MERGE : INT
        {
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ReaderWriterLockSlim[] sqlitelock { get; set; }
            public string[] strConn { get; set; }
            public SQLiteConnection[] conn { get; set; }
            public SQLiteCommand[] cmd { get; set; }

            public SQLiteDataAdapter[] sqlAdapter { get; set; }
            public SQLiteCommandBuilder[] sqlcmdbuilder { get; set; }
            public SQLiteDataReader[] SqReader { get; set; }

            public DbDataReader[] DbReader { get; set; }
            public DataSet[] ds { get; set; }
            public DataTable dt_test { get; set; }
            public DataTable[] dt { get; set; }
            public SQLiteTransaction[] tran { get; set; }

            public ManualResetEvent[] ThreadFlags { get; set; }
            public ManualResetEvent[] Insert_ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }


            public string FilePath { get; set; }
            public string RefHeader { get; set; }
            public bool[] Wait { get; set; }

            public int Limit { get; set; }
            public int Limit_Count { get; set; }
            public int Table_Count { get; set; }
            public bool[] Insert_Thread_Wait { get; set; }
            public double[] Testtime { get; set; }

            public object[] ID { get; set; }
            public object[] Value { get; set; }
            public object[] WAFER_ID { get; set; }
            public object[] LOT_ID { get; set; }
            public object[] SITE_ID { get; set; }
            public Dictionary<string, double[]> Selected_Parameter_Distribution { get; set; }
            public double[][] test { get; set; }

            double[] Testtime1 { get; set; }
            double[] Testtime2 { get; set; }
            double[] Testtime3 { get; set; }
            public string[][] Teststring { get; set; }
            public double[][] Testdouble { get; set; }


            public Dictionary<string, IQR> DIC_IQR { get; set; }
            public object[] Variation { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test_New_Spec { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent { get; set; }
            public List<List<int>>[] For_Any_Yield { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_Lot { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_SITE { get; set; }

            public List<List<int>[]>[] ForCampare_Yield { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent_For_New_Spec { get; set; }
            public List<List<int>>[] For_Any_Yield_For_New_Spec { get; set; }
            public List<List<int>[]>[] For_New_Spec_ForCampare_Yield { get; set; }
            public List<int[]>[] ForCampare_Yield_Fro_DB { get; set; }
            public List<List<int[]>>[] ForCampare_Yield_Fro_DB_List { get; set; }
            public List<List<List<List<int>[]>>>[] ForCampare_Yield_DB_LotVariation { get; set; }
            public List<List<int>>[] For_New_Spec_ForCampare_Yield2 { get; set; }
            public List<List<List<int[]>>>[] ForCampare_Yield_Fro_DB_List_LotVariation { get; set; }
            public Dictionary<string, int> Refer_Site_And_Num { get; set; }
            public Dictionary<string, int> Refer_Lot_And_Num { get; set; }
            public List<int>[] ForCampare_Yield_List { get; set; }
            public List<List<int>[]> ForCampare_Yield_List1 { get; set; }
            public List<List<int>[]>[] ForCampare_Yield_List2 { get; set; }
            public Dictionary<string, Values> Values { get; set; }
            public Dictionary<string, Data_Calculation> Cal_Value_by_rowsdata { get; set; }
            public Dictionary<string, Data_Calculation> For_New_Spec_Cal_Value_by_rowsdata { get; set; }
            public List<int>[] Check { get; set; }
            public List<List<int>[]> Test { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public List<double[]>[] DB_DataSet_Values { get; set; }

            public Dictionary<string, int> Lot_Dic { get; set; }
            public Dictionary<string, int> Site_Dic { get; set; }
            public Dictionary<string, int> Bin_Dic { get; set; }
            public Dictionary<string, Dictionary<string, List<string>>> Matching_Lots { get; set; }
            public Dictionary<string, List<string>> Matching_Lot { get; set; }
            public Stopwatch[] TestTime1 { get; set; }
            public Stopwatch[] TestTime2 { get; set; }
            public Stopwatch[] TestTime3 { get; set; }
            public Stopwatch[] TestTime4 { get; set; }
            public Stopwatch[] TestTime5 { get; set; }
            public long SampleCount { get; set; }
            public object Update_Data_ID { get; set; }
            public string[] Update_Datas_ID { get; set; }
            public string Get_Gross_Para { get; set; }

            public double Get_Gross_Persent { get; set; }
            public string Get_Gross_Selector { get; set; }
            public object[] Std_Value { get; set; }
            public double[] Std_Value_Convert { get; set; }
            public List<Dictionary<string, Gross>[]> List_Gross_Values { get; set; }
            public Dictionary<string, Gross>[] Gross_Values1 { get; set; }
            public long NB { get; set; }


            public Dictionary<string, CSV_Class.For_Box>[] Dic_Test { get; set; }
            public  Dictionary<string, CSV_Class.For_Box> Dic_Test_For_Spec_Gen { get; set; }
            public string Table { get; set; }
            public string Filename { get; set; }

            public double[] Make_New_Spec_For_Yield_Min { get; set; }
            public double[] Make_New_Spec_For_Yield_Max { get; set; }
            public List<string> Gross { get; set; }
            public List<string[]>[] DataSet_Value { get; set; }
            public List<double[]>[] DataSet_Double_Value { get; set; }
            public string Lot_ID { get; set; }
            public string SubLot_ID { get; set; }
            public string Tester_ID { get; set; }
            public string Site { get; set; }
            public string Bin { get; set; }
            public string ID_Unit { get; set; }
            public int Bin_place { get; set; }

            public string Query { get; set; }
            public string Query2 { get; set; }
            public string CellID { get; set; }
            public bool _From_Db { get; set; }
            public int Spec_Table_Count { get; set; }
            public bool _Flag { get; set; }
            public bool _SUBLOT_Flag { get; set; }
            public bool Clotho_Spec_Flag { get; set; }
            public string Before_Lot_ID { get; set; }
            public string Changed_Lot_ID { get; set; }

            public int[] Each_Thread_Count { get; set; }

            public string[] No_Index { get; set; }
            public string[] Paraname { get; set; }
            public string[] SpecMin { get; set; }
            public string[] SpecMax { get; set; }
            public string[] DataMin { get; set; }
            public  string[] DataMedian { get; set; }
            public  string[] DataMax { get; set; }
            public  string[] CPK { get; set; }
            public  string[] STD { get; set; }
            public  string[] Percent { get; set; }
            public  string[] Fail { get; set; }

            public string[] Line { get; set; }

            public int Count_Current_Setting { get; set; }

            public void Open_DB(string FileName, Data_Class.Data_Editing.INT Data_Edit)
            {
                string Filename = FileName.Substring(FileName.LastIndexOf("\\") + 1);
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "\\" + Filename.Substring(0, FileName.Length - 4) + "_" + i + ".db";
                    //  strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    cmd[i].ExecuteNonQuery();

                }


                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                dt = new DataTable[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    dt[i] = new DataTable();
                    cmd[i].CommandText = "PRAGMA synchronous";
                    SqReader[i] = cmd[i].ExecuteReader();
                    dt[i].Load(SqReader[i]);
                }



            }

            public void Open_DB(string[] FileName, Data_Class.Data_Editing.INT Data_Edit)
            {
                Data_Edit.DB_Count = FileName.Length;
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    string Filename = FileName[i].Substring(FileName[i].LastIndexOf("\\") + 1);

                    int length = Filename.Length;
                    Filename = Filename.Substring(0, length - 5);

                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "\\" + Filename + i + ".db";
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + ".csv\\" + Filename.Substring(0, Filename.Length) + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    //cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    //cmd[i].ExecuteNonQuery();

                }

            }

            public void DropTable(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {

                try
                {
                    for (int i = 0; i < Data_Edit.DB_Count; i++)
                    {
                        cmd[i].CommandText = "";
                        cmd[i].CommandText = "drop TABLE " + Query;
                        cmd[i].ExecuteNonQuery();
                    }
                }
                catch { }
            }

            public void Insert_Header(Data_Class.Data_Editing.INT Data_Edit)
            {



                Lot_ID = Lot_ID.Replace('-', '_');
                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(MakecolumnsThread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();
                }

                if (Clotho_Spec_Flag)
                {

                    Data = Data_Edit;
                    ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                    Wait = new bool[Data_Edit.DB_Count];
                    Wait = new bool[Data_Edit.DB_Count];
                    Testtime = new double[Data_Edit.DB_Count];

                    Data.Data_Table = "Clotho_Spec";
                    for (int i = 0; i < Data_Edit.DB_Count; i++)
                    {
                        stringA[i].Clear();
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(MakecolumnsThread1), i);
                    }

                    for (int i = 0; i < Data_Edit.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                        stringA[i] = new StringBuilder();
                    }
                }
                //    Data.Data_Table = "data0";
            }

            public void MakecolumnsThread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];
                cmd[i] = new SQLiteCommand(conn[i]);

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE " + Lot_ID + "(" + Data.New_Header[0] + " VARCAHR(20)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE " + Lot_ID + "(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(20)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(20)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                      //  if (_SUBLOT_Flag == true)
                      //  {
                       //     stringA[i].Append(", id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");
                     //   }
                      //  else
                      //  {

                            stringA[i].Append(", SubLot VARCAHR(5), id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(20), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");
                     //   }

                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void MakecolumnsThread1(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE " + Data.Data_Table + "(" + Data.New_Header[0] + " VARCAHR(20)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE " + Data.Data_Table + "(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(20)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(20)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                    //    if (_SUBLOT_Flag == true)
                    //    {
                    //        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");

                    //    }
                    //    else
                    //    {
                            stringA[i].Append(", SubLot VARCAHR(5), id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");
                     //   }

                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }
            public void Insert_Current_Setting(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }
            public void Insert_Current_Setting_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {
                Data = Data_Edit;
                this.Table = Table;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  cmd[i].Reset();
                    //    ThreadFlags[i] = new ManualResetEvent(false);
                    Insert_Current_Setting_Data_Thread(i);
                    //  ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //       Wait[i] = ThreadFlags[i].WaitOne();
                }

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Clotho_Spec_Max_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}

            }

            public void Insert_Current_Setting_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();
                stringA[i].Clear();
                SampleCount = 1;

                cmd[i] = new SQLiteCommand(conn[i]);

                int Count = Data.Per_DB_Column_Count[i];


                int k = 0;


                if (Table.ToUpper() == "CLOTHO_SPEC")
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Min[Spec_Count] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Min[Spec_Count] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                    }




                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }
                else
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[0] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Min[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;

                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }
                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }




                //   ThreadFlags[i].Set();
            }
            public void Insert_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE spec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE spec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY );");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_New_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_New_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void Insert_New_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE newspec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE newspec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY );");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                // sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];
                sqlitelock = new ReaderWriterLockSlim[Data.DB_Count];
                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                //Testdouble = new double[7][];

                //Testdouble[0] = new double[Data.DB_Column_Limit];
                //Testdouble[1] = new double[Data.DB_Column_Limit];
                //Testdouble[2] = new double[Data.DB_Column_Limit];
                //Testdouble[3] = new double[Data.DB_Column_Limit];
                //Testdouble[4] = new double[Data.DB_Column_Limit];
                //Testdouble[5] = new double[Data.DB_Column_Limit];
                //Testdouble[6] = new double[Data.Per_DB_Column_Count[6]];
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder(100000);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder(100000);
                    Testtime[i] = TestTime1.Elapsed.TotalMilliseconds;
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);
            }
            public void Insert_Ref_Header_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data.Data_Table = "REFHEADER";

                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                // sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];
                sqlitelock = new ReaderWriterLockSlim[Data.DB_Count];


                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder(100000);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Ref_Header_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder(100000);
                    Testtime[i] = TestTime1.Elapsed.TotalMilliseconds;
                }

            }

            public void Insert_Ref_Header_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Reference_Header[0] + "',");
                    // stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[0].Replace("PID-", "") + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Reference_Header[(Data.DB_Column_Limit * i)] + "',");

                }

                for (k = 1; k < Count; k++)
                {
                    stringA[i].Append("'" + Data.Reference_Header[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.Reference_Header[(Data.DB_Column_Limit * i) + k] + "', '" + SubLot_ID + "', '" + SampleCount + "' , '" + Lot_ID + "' , '" + Site + "' ,'0');");

                // stringA[i].Append("'" + SubLot_ID + "', '" + SampleCount + "' , '0');");

                // stringA[i].Append("'1, '" + SubLot_ID + "', '" + SampleCount + "' , '0');");
                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void Insert_Data(long Sample)
            {
                SampleCount = Sample;

                Lot_ID = Lot_ID.Replace('-', '_');
                //ForCampare_Yield_List = new List<int>[Data.DB_Count];

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    ForCampare_Yield_List[i] = new List<int>();
                //}

                //for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                //{
                //    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                //    {
                //        ForCampare_Yield_List[i].Add(0);
                //    }
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    Insert_Data_NoThread(i);

                //}

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);

                   // Insert_Data_Thread(i);

                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                //   ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Insert_Data_NoThread(int DB)
            {
                int i = DB;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[0].Replace("PID-", "") + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[(Data.DB_Column_Limit * i)] + "',");

                }

                for (k = 1; k < Count; k++)
                {
                    stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "', '" + SubLot_ID + "', '" + SampleCount + "' , '0');");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();

                stringA[i].Clear();
                int k = 0;

                cmd[i] = new SQLiteCommand(conn[i]);
            

                if (i == 0)
                {

                    stringA[i].Append("INSERT INTO " + Lot_ID + " VALUES ('" + Data.Getstring[0] + "',");

                    for (k = 1; k < 5; k++)
                    {
                        stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "',");

                    }

                    stringA[i].Append("'" + Tester_ID + "',");
                    k++;

                    for (k = k; k < Count; k++)
                    {
                        stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "',");

                    }

                    // stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[0].Replace("PID-", "") + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Lot_ID + " VALUES ('" + Data.Getstring[(Data.DB_Column_Limit * i)] + "',");

                    for (k = 1; k < Count; k++)
                    {
                        stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "',");

                    }

                }







                stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "', '" + SubLot_ID + "', '" + SampleCount + "' , '" + Lot_ID + "' , '" + Site + "' ,'0', '" + Bin + "');");

                // stringA[i].Append("'" + SubLot_ID + "', '" + SampleCount + "' , '0');");

                // stringA[i].Append("'1, '" + SubLot_ID + "', '" + SampleCount + "' , '0');");
                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void Insert_Data_Get_From_DB(int Sample)
            {

            }
            public void Insert_Spec_Get_From_DB(Data_Class.Data_Editing.INT Data_Edit)
            {


                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Get_From_DB_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }

            public void Insert_Spec_Get_From_DB_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Insert_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;
                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[0] + "',");

                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[Data.DB_Column_Limit * i] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_LowSpec[Data.DB_Column_Limit * i + k] + "', '0', '0' ,'0','0', '0', '0');");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                stringA[i].Clear();
                cmd[i].Reset();
                k = 0;
                SampleCount = 2;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[0] + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[Data.DB_Column_Limit * i] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_HighSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_HighSpec[Data.DB_Column_Limit * i + k] + "', '1', '1' , '1', '1', '1', '1');");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {

            }


            public void Insert_Files_Name(string Filename)
            {

                this.Filename = Filename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(thread_Insert_File_Name), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }

            public void thread_Insert_File_Name(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();


                stringA[i].Append("INSERT INTO Files VALUES ('" + this.Filename + "');");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Make_table(string Tablename)
            {
                stringA[0].Clear();
                stringA[0].Append("CREATE TABLE " + Tablename + "( FIRST VARCAHR(5), END VARCAHR(5), DBCOUNT VARCHAR(5), COLUMNCOUNT VARCHAR(5) );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";

                stringA[0].Clear();
                stringA[0].Append("INSERT INTO INF VALUES ('" + TheFirst_Trashes_Header_Count + "' , '" + TheEnd_Trashes_Header_Count + "' , '" + Data.Per_DB_Column_Count.Length + "' , '" + Data.Per_DB_Column_Count[Data.Per_DB_Column_Count.Length - 1] + "' );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";
            }

            public void Make_table2(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {
                Data = Data_Edit;
                Table = Tablename;

                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(_Make_Table), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void _Make_Table(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE " + Table + "(" + Data.New_Header[0] + " VARCAHR(20)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE " + Table + "(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(20)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(20)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", SubLot VARCAHR(5), id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), Fail VARCHAR(20));");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Make_table_For_Filename(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {
                Data = Data_Edit;
                Table = Tablename;

                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(_Make_Table_For_Filename), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void _Make_Table_For_Filename(Object threadContext)
            {
                int i = (int)threadContext;

                stringA[i].Append("CREATE TABLE " + Table + "(FIle VARCAHR(20))");


                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].CommandText = "";

                stringA[i].Append(",");

                ThreadFlags[i].Set();
            }

            public void Make_table_For_Trace(string Tablename,string Chan, bool Flag)
            {
                if (!Flag)
                {
                    stringA[0].Clear();
                    stringA[0].Append("CREATE TABLE " + Tablename + "(Chan VARCAHR(5), Info VARCAHR(5));");

                    cmd[0] = new SQLiteCommand(conn[0]);
                    cmd[0].CommandText = stringA[0].ToString();


                    cmd[0].CommandText = stringA[0].ToString();
                    cmd[0].ExecuteNonQuery();

                    cmd[0].Dispose();
                    SqReader[0].Close();

                }
                if (Flag)
                {

                    stringA[0].Clear();
                    stringA[0].Append("INSERT INTO Trace_Info VALUES ('" + Chan + "','" +   Tablename +  "' );");

                    cmd[0] = new SQLiteCommand(conn[0]);
                    cmd[0].CommandText = stringA[0].ToString();


                    cmd[0].ExecuteNonQuery();
                    cmd[0].Dispose();
        
                }
            }
            public void Delete_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Delete_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Delete_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;
                int k = 0;


                stringA[i].Append("Delete from " + Table + " where id = 0");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                stringA[i].Clear();

                stringA[i].Append("Delete from " + Table + " where id = 1");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Delete_Lot_Data(string Query)
            {
                this.Query = Query;

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Delete_Lot_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
            }
            public void Delete_Lot_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                stringA[i].Append(this.Query);

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                ThreadFlags[i].Set();
            }

            public void Save_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                //Update_Data_ID = data;

                //if (data != null)
                //{
                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        ThreadFlags[i] = new ManualResetEvent(false);
                //        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                //    }

                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        Wait[i] = ThreadFlags[i].WaitOne();
                //    }
                //}

            }
            public void Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    //  Insert_table_Data_Thread(i);
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_table_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}


            }

            public void Road_Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                // SampleCount = Sample;

                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Road_Save_Customer_Spec_table_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Road_Save_Customer_Spec_table_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void LOTID_Update(string Query, string Query2, string CellID)
            {

                this.Query = Query;
                this.Query2 = Query2;
                this.CellID = CellID;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ThreadFlags[i] = new ManualResetEvent(false);
                    //   LOTID_Update_Thread(i);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(LOTID_Update_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }
            public void LOTID_Update_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                if (this.CellID == "DIE_X")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }

                }
                else if (this.CellID == "DIE_Y")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "TIME")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "TOTAL_TESTS")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "WAFER_ID")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "LOTID")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();

                        if (this.Query2 != null)
                        {
                            cmd[i].CommandText = this.Query2;
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                    else
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();

                    }
                }




                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }


            public void Gross_Update_Data(object data)
            {
                Update_Data_ID = data;

                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }

            }
            public void Gross_Update_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (object o in (Array)Update_Data_ID)
                {
                    cmd[i].CommandText = "Update data set FAIL = '1'  where id = " + o.ToString();
                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Gross_Update_Datas(List<string> data)
            {
                Update_Datas_ID = data.ToArray();
                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Datas_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }
            }
            public void Gross_Update_Datas_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (object o in (Array)Update_Datas_ID)
                {
                    cmd[i].CommandText = "Update data set FAIL = '1'  where id = " + o.ToString();
                    cmd[i].ExecuteNonQuery();
                    cmd[i].Reset();
                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Chnaged_Spec_Update_Data(int DB, int Index, string Parameter, double Spec, int GetId)
            {
                stringA[DB].Clear();
                stringA[DB].Append("Update newspec set " + Parameter + " = " + Spec + " where id = " + GetId);

                cmd[DB].CommandText = stringA[DB].ToString();

                cmd[DB].ExecuteNonQuery();
                cmd[DB].Reset();

                stringA[DB].Clear();
            }
            public Dictionary<string, double[]> Chnaged_Spec_Anl_Yield(int DB, int Index, string Parameter)
            {
                stringA[DB].Clear();
                Dictionary<string, double[]> Dic_Change_Spec = new Dictionary<string, double[]>();


                stringA[DB].Append("Select " + Parameter + " from newspec");

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                object[] GetSpec = new object[ds[DB].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    GetSpec[count] = dr.ItemArray[0];
                    count++;
                }

                double[] Toduble_Spec = Array.ConvertAll<object, double>(GetSpec, Convert.ToDouble);

                Dic_Change_Spec.Add("SPEC", Toduble_Spec);
                stringA[DB].Clear();

                stringA[DB].Append("Select " + Parameter + " from data where Fail not like '1'");

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                object[] GetData = new object[ds[DB].Tables[0].Rows.Count];
                count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    GetData[count] = dr.ItemArray[0];
                    count++;
                }

                double[] Toduble_Data = Array.ConvertAll<object, double>(GetData, Convert.ToDouble);

                Dic_Change_Spec.Add("DATA", Toduble_Data);

                stringA[DB].Clear();

                return Dic_Change_Spec;
            }
            public void Get_Ave_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                test = new double[Data.DB_Count][];
                double[] test1 = new double[Data.DB_Count];

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    test[i] = new double[10000];
                    Testtime[i] = new double();
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                    //ThreadFlags[i] = new ManualResetEvent(false);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //Wait[i] = ThreadFlags[i].WaitOne();
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                }


            }

            public void Get_Ave_Data_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                test = new double[Data.DB_Count][];
                double[] test1 = new double[Data.DB_Count];

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    test[i] = new double[10000];
                    Testtime[i] = new double();
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                    //ThreadFlags[i] = new ManualResetEvent(false);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //Wait[i] = ThreadFlags[i].WaitOne();
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                }


            }
            public void Set_Refer_for_Anlyzer(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_Ave_Data2(Data_Class.Data_Editing.INT Data_Edit)
            {

                test = new double[Data.DB_Count][];
                double[] test1 = new double[Data.DB_Count];

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    test[i] = new double[10000];
                    Testtime[i] = new double();
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                    //ThreadFlags[i] = new ManualResetEvent(false);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //Wait[i] = ThreadFlags[i].WaitOne();
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                }


            }
            public void Get_Ave_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from data where Fail not like '1'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                List<double[]> DataSet_Values = new List<double[]>();
                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);
                    DataSet_Values.Add(doubles);

                    count++;

                }
                SqReader[i].Close();

                STDEVandMedian(DataSet_Values, i, count);

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();
            }

            public void Get_Saved_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                DataSet_Value = new List<string[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    DataSet_Value[i] = new List<string[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Saved_Spec_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }

            public void Get_Saved_Spec_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                cmd[i] = new SQLiteCommand(conn[i]);
                stringA[i].Append("Select * from Clotho_Spec");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    DataSet_Value[i].Add(stringD);

                    count++;

                }
                SqReader[i].Close();

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].Dispose();
                ThreadFlags[i].Set();

            }

            public void Get_Rows_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface)
            {
                //stringA[DB].Clear();
                //stringA[DB].Append("Select id, " + Select_Para + " from data");

                //cmd[DB].CommandText = stringA[DB].ToString();
                //ds[DB] = new DataSet();

                //sqlAdapter[DB].SelectCommand = cmd[DB];
                //sqlAdapter[DB].Fill(ds[DB]);

                //ID = new object[ds[DB].Tables[0].Rows.Count];
                //Value = new object[ds[DB].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //{
                //    ID[count] = dr.ItemArray[0];
                //    Value[count] = dr.ItemArray[1];

                //    count++;
                //}

                //double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);

     
                //stringA[DB].Clear();
            }

            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface, DataTable dt)
            {
                //stringA[DB].Clear();
                //stringA[DB].Append("Select id, " + Select_Para + " from data");

                //cmd[DB].CommandText = stringA[DB].ToString();
                //ds[DB] = new DataSet();

                //sqlAdapter[DB].SelectCommand = cmd[DB];
                //sqlAdapter[DB].Fill(ds[DB]);

                //ID = new object[ds[DB].Tables[0].Rows.Count];
                //Value = new object[ds[DB].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //{
                //    ID[count] = dr.ItemArray[0];
                //    Value[count] = dr.ItemArray[1];

                //    count++;
                //}

                //double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);


                //stringA[DB].Clear();
            }
            public void Get_Selected_Para(int DB, string Select_Para, bool Flag, string Selector)
            {


            }
            public double[] Get_Find_Bin(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);
                sqlAdapter[0].Dispose();
                cmd[0].Dispose();
                stringA[0].Clear();
                return doubles;
            }
            public List<object[]> Get_Data_By_Querys(string Query)
            {
                return null;
            }
            public string[] Get_Data_By_Query(string Query)
            {
                stringA[0].Clear();

                stringA[0].Append(Query);

                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[5000000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    Value1[count] = stringD[0];

                    count++;

                }

                Array.Resize(ref Value1, count);

                cmd[0].Dispose();
                SqReader[0].Close();

                string[] _string = Array.ConvertAll<object, string>(Value1, Convert.ToString);


                stringA[0].Clear();
                return _string;
            }

            public Dictionary<string, object[]> Get_Data_By_Query_S4PD(string Query, string Chan)
            {
                stringA[0].Clear();


                stringA[0].Append("PRAGMA table_info(" + Chan + ")");

                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[5000000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    Value1[count] = values[1];

                    count++;

                }

                Array.Resize(ref Value1, count);



                stringA[0].Clear();
                stringA[0].Append("select Count(Freq) from " + Chan);

                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] data = new object[1];
                count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    data = new object[Convert.ToInt64(values[0])];

                }

                long length = Convert.ToInt64(data.Length);

                Dictionary<string, object[]> Test = new Dictionary<string, object[]>();

                for (int i = 1; i < Value1.Length - 6;  i++)
                {
                    stringA[0].Clear();
                    stringA[0].Append("select " + Value1[i] + " from " + Chan);

                    cmd[0] = new SQLiteCommand(conn[0]);
                    cmd[0].CommandText = stringA[0].ToString();
                    SqReader[0] = cmd[0].ExecuteReader();

                    data = new object[length];
                    count = 0;

                    while (SqReader[0].Read())
                    {
                        object[] values = new object[SqReader[0].FieldCount];
                        SqReader[0].GetValues(values);
                        data[count] = values[0];

                        count++;

                    }

                    Test.Add(Convert.ToString(Value1[i]), data);

                }


                cmd[0].Dispose();
                SqReader[0].Close();


                stringA[0].Clear();
                return Test;
            }

            public string[] Get_Data_By_Query(string Query, int DB)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);

                stringA[0].Clear();
                return _string;
            }

            public void Get_Defined_Para(object[,] DummyData, string key, Data_Class.Data_Editing.INT Data_InterFace)
            {


            }

            public void Get_Gross_Check_Para(Data_Class.Data_Editing.INT Data_Edit, string Select_Para, double Persent, string Selector, int SelectedBin)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                Get_Gross_Para = Select_Para;
                Get_Gross_Persent = Persent;
                //   Gross = ForGross_Fail_Unit;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = false;
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Gross_Check_Para_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
                double test = TestTime1.Elapsed.TotalMilliseconds;


                //  stringA[0].Append("Select id from data where id not like '%F%'");
                stringA[0].Append("Select id from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                ID = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    ID[count] = dr.ItemArray[0];
                    count++;
                }

                stringA[0].Clear();
                // List_Gross_Values.Add(Gross_Values1);
            }

            public void Get_Gross_Check_Para_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                int k = 0;
                for (k = 0; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    string[] Split_Dummy = Data.Reference_Header[Data.DB_Column_Limit * i + k].Split('_');
                    if (Split_Dummy.Length != 1)
                    {
                        if (Split_Dummy[1].ToUpper() == Get_Gross_Para.ToUpper())
                        {
                            ds[i] = new DataSet();

                            //   stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data where id not like '%F%'");
                            stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data");
                            cmd[i].CommandText = stringA[i].ToString();

                            sqlAdapter[i].SelectCommand = cmd[i];
                            sqlAdapter[i].Fill(ds[i]);

                            object[] DataValue = new object[ds[i].Tables[0].Rows.Count];

                            int count = 0;
                            foreach (DataRow dr in ds[i].Tables[0].Rows)
                            {
                                DataValue[count] = dr.ItemArray[0];
                                count++;
                            }


                            double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

                            double DataMin = doubles.Min();
                            double DataMax = doubles.Max();
                            double DataAve = doubles.Average();

                            double DataMinindex = doubles.ToList().IndexOf(DataMin);
                            double DataMaxindex = doubles.ToList().IndexOf(DataMax);

                            double Divide = DataMax / DataMin;

                            string[] test;
                            string _Substring = Get_Gross_Para.Substring(0, 1);

                            double MinSpec = 0f;
                            bool Define_Flag = false;

                            if (Get_Gross_Para.ToUpper().Contains("IBATT") || Get_Gross_Para.ToUpper().Contains("ICC") || Get_Gross_Para.ToUpper().Contains("IDD"))
                            {
                                Define_Flag = true;
                                test = Convert.ToString(Get_Gross_Persent).Split('.');
                                MinSpec = 1 - (Convert.ToDouble(test[1]) / 10);
                            }
                            else
                            {
                                Define_Flag = false;
                                MinSpec = Convert.ToDouble(Get_Gross_Persent) * -1;
                            }

                            if (Define_Flag)
                            {
                                for (int j = 0; j < doubles.Length; j++)
                                {
                                    if (DataAve / doubles[j] > Get_Gross_Persent || DataAve / doubles[j] < MinSpec)
                                    {
                                        if (!Gross.Contains(Convert.ToString(j + 1)))
                                        {
                                            //       Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                for (int j = 0; j < doubles.Length; j++)
                                {
                                    if (DataAve - doubles[j] > Get_Gross_Persent || doubles[j] - DataAve < MinSpec)
                                    {
                                        if (!Gross.Contains(Convert.ToString(j + 1)))
                                        {
                                            //         Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                                        }
                                    }
                                }
                            }

                            stringA[i].Clear();
                            cmd[i].CommandText = "";
                        }
                        //if (Get_Gross_Para == "POUT" && Split_Dummy.Length > 7 && Split_Dummy[6].ToUpper() == "FIXEDPOUT" && Split_Dummy[1].ToUpper() == "POUT")
                        //{
                        //    ds[i] = new DataSet();


                        //    //    stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data where id not like '%F%'");
                        //    stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data");
                        //    cmd[i].CommandText = stringA[i].ToString();

                        //    sqlAdapter[i].SelectCommand = cmd[i];
                        //    sqlAdapter[i].Fill(ds[i]);

                        //    object[] DataValue = new object[ds[i].Tables[0].Rows.Count];

                        //    int count = 0;

                        //    foreach (DataRow dr in ds[i].Tables[0].Rows)
                        //    {
                        //        DataValue[count] = dr.ItemArray[0];
                        //        count++;
                        //    }

                        //    string remove = Split_Dummy[7].Replace("dBm", "");

                        //    double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

                        //    double DataMin = Convert.ToDouble(remove) - Get_Gross_Persent;
                        //    double DataMax = Convert.ToDouble(remove) + Get_Gross_Persent;

                        //    for (int j = 0; j < doubles.Length; j++)
                        //    {
                        //        if (doubles[j] < DataMin)
                        //        {
                        //            if (!Gross.Contains(Convert.ToString(j + 1)))
                        //            {
                        //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                        //            }

                        //        }
                        //        else if (doubles[j] > DataMax)
                        //        {
                        //            if (!Gross.Contains(Convert.ToString(j + 1)))
                        //            {
                        //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                        //            }
                        //        }
                        //    }

                        //    stringA[i].Clear();
                        //    cmd[i].CommandText = "";
                        //}

                    }
                }
                ThreadFlags[i].Set();
            }
            public void Get_From_Db_Data_for_Anly(Data_Class.Data_Editing.INT Data_Edit)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;


                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data), i);


                //}
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Get_From_Db_Data(i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }

            public void Get_Current_Setting(Data_Class.Data_Editing.INT Data_Edit, int NB)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Get_From_Db_Data(i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Get_From_Db_Data(Object threadContext)
            {
                int i = (int)threadContext;

                CSV_Class.CSV.MERGE csv = new CSV_Class.CSV.MERGE();

                int Count = Data.Per_DB_Column_Count[i];
                this.Filename = "";


                if (_Flag)
                {

                    string[] name = conn[i].DataSource.ToString().Split('_');

                    for (int namei = 0; namei < 3; namei++)
                    {
                        if (namei == 2)
                        {
                            this.Filename += Lot_ID;
                        }
                        else
                        {
                            this.Filename += name[namei] + "_";
                        }

                    }
                    csv.Write_Open("C:\\Automation\\DB\\YIELD\\" + conn[i].DataSource.ToString().Substring(0, conn[i].DataSource.ToString().Length - 2) + ".csv\\" + this.Filename + "_" + i + ".csv");
                    string a = "C:\\Automation\\DB\\YIELD\\" + conn[i].DataSource.ToString().Substring(0, conn[i].DataSource.ToString().Length - 2) + ".csv\\" + this.Filename + ".csv";
                }
                else
                {

                    csv.Write_Open("C:\\Automation\\DB\\YIELD\\" + conn[i].DataSource.ToString().Substring(0, conn[i].DataSource.ToString().Length - 2) + ".csv\\" + conn[i].DataSource.ToString() + ".csv");
                }




                StringBuilder Apped = new StringBuilder();

                for (int Row = 0; Row < 3; Row++)
                {
                    if (Row == 0)
                    {

                        #region header
                        for (int j = 0; j < Count; j++)
                        {
                            if (j == 0)
                            {
                                if (i == 0)
                                {
                                    Apped.Append(Data.Reference_Header[0] + ",");
                                }
                                else
                                {
                                    Apped.Append(Data.Reference_Header[Data.DB_Column_Limit * i] + ",");
                                }

                            }
                            else
                            {
                                Apped.Append(Data.Reference_Header[Data.DB_Column_Limit * i + j] + ",");
                            }

                            if (j == Count - 1)
                            {
                                if (i == Data.Per_DB_Column_Count.Length - 1)
                                {
                                    Apped.Append("SubLot");
                                    csv.Write(Apped.ToString());
                                    Apped.Clear();
                                }
                                else
                                {
                                    Apped.Append("");
                                    csv.Write(Apped.ToString());
                                    Apped.Clear();
                                }

                            }

                        }
                        #endregion
                    }
                    else if (Row == 1)
                    {
                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count - 6;
                        }

                        #region Spec high
                        for (int j = 0; j < Count; j++)
                        {
                            if (i == 0)
                            {
                                if (j == 0)
                                {
                                    Apped.Append("HighL,");
                                }
                                else if (j < 10)
                                {
                                    Apped.Append(",");
                                }
                                else
                                {
                                    Apped.Append(DataSet_Value[i][1][j] + ",");
                                }
                            }
                            else if (i == Data.Per_DB_Column_Count.Length - 1)
                            {
                                if (j < Count)
                                {
                                    Apped.Append(DataSet_Value[i][1][j] + ",");
                                }
                                else
                                {
                                    int dummy_row = 0;
                                    for (dummy_row = 0; dummy_row < 6; j++)
                                    {
                                        Apped.Append(",");
                                    }

                                }
                            }
                            else
                            {
                                Apped.Append(DataSet_Value[i][1][j] + ",");
                            }

                            if (j == Count - 1)
                            {
                                Apped.Append("");
                                csv.Write(Apped.ToString());
                                Apped.Clear();
                            }
                        }

                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count + 6;
                        }
                        #endregion
                    }
                    else if (Row == 2)
                    {

                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count - 6;
                        }
                        #region Low
                        for (int j = 0; j < Count; j++)
                        {
                            if (i == 0)
                            {
                                if (j == 0)
                                {
                                    Apped.Append("LowL,");
                                }
                                else if (j < 10)
                                {
                                    Apped.Append(",");
                                }
                                else
                                {
                                    Apped.Append(DataSet_Value[i][0][j] + ",");
                                }
                            }
                            else if (i == Data.Per_DB_Column_Count.Length - 1)
                            {
                                if (j < Count)
                                {
                                    Apped.Append(DataSet_Value[i][0][j] + ",");
                                }
                                else
                                {
                                    int dummy_row = 0;
                                    for (dummy_row = 0; dummy_row < 6; j++)
                                    {
                                        Apped.Append(",");
                                    }

                                }
                            }
                            else
                            {
                                Apped.Append(DataSet_Value[i][0][j] + ",");
                            }

                            if (j == Count - 1)
                            {
                                Apped.Append("");
                                csv.Write(Apped.ToString());
                                Apped.Clear();
                            }
                        }

                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count + 6;
                        }
                        #endregion
                    }

                }

                if (i == Data.Per_DB_Column_Count.Length - 1)
                {
                    Count = Count + 1;
                }




                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {

                    Dictionary<string, List<string>> tests = key.Value;


                    if (!_Flag)
                    {
                        foreach (KeyValuePair<string, List<string>> ts in tests)
                        {

                            stringA[i].Clear();

                            if (_Flag == true)
                            {
                                stringA[i].Append("Select * from " + key.Key + " where Fail = '0'");
                            }
                            else
                            {
                                stringA[i].Append("Select * from  " + key.Key + " where Fail = '0'");
                            }

                            cmd[i] = new SQLiteCommand(conn[i]);
                            cmd[i].CommandText = stringA[i].ToString();
                            SqReader[i] = cmd[i].ExecuteReader();


                            while (SqReader[i].Read())
                            {

                                Stopwatch TestTime1 = new Stopwatch();
                                TestTime1.Restart();
                                TestTime1.Start();


                                object[] values = new object[SqReader[i].FieldCount];
                                SqReader[i].GetValues(values);

                                for (int j = 0; j < Count; j++)
                                {
                                    Apped.Append(values[j] + ",");
                                }

                                csv.Write(Apped.ToString());
                                Apped.Clear();

                            }
                            SqReader[i].Close();

                            stringA[i].Clear();
                            cmd[i].Dispose();

                            break;
                        }

                    }
                    else
                    {
                        foreach (KeyValuePair<string, List<string>> ts in tests)
                        {
                            if (ts.Key == Lot_ID)
                            {
                                stringA[i].Clear();

                                if (_Flag == true)
                                {
                                    stringA[i].Append("Select * from " + key.Key + " where Fail = '0'");
                                }
                                else
                                {
                                    stringA[i].Append("Select * from  " + key.Key + " where Fail = '0'");
                                }

                                cmd[i] = new SQLiteCommand(conn[i]);
                                cmd[i].CommandText = stringA[i].ToString();
                                SqReader[i] = cmd[i].ExecuteReader();


                                while (SqReader[i].Read())
                                {

                                    Stopwatch TestTime1 = new Stopwatch();
                                    TestTime1.Restart();
                                    TestTime1.Start();


                                    object[] values = new object[SqReader[i].FieldCount];
                                    SqReader[i].GetValues(values);

                                    for (int j = 0; j < Count; j++)
                                    {
                                        Apped.Append(values[j] + ",");
                                    }

                                    csv.Write(Apped.ToString());
                                    Apped.Clear();

                                }
                                SqReader[i].Close();

                                stringA[i].Clear();
                                cmd[i].Dispose();

                                break;
                            }
                        }
                    }




                }





                #region
                //for (int loop = 0; loop < Table_Count; loop++)
                //{
                //    stringA[i].Clear();

                //    if (_Flag == true)
                //    {
                //        stringA[i].Append("Select * from data" + loop + " where Fail = '0' and LOTID =" + "'" + Lot_ID + "'");
                //    }
                //    else
                //    {
                //        stringA[i].Append("Select * from data" + loop + " where Fail = '0'");
                //    }


                //    cmd[i].CommandText = stringA[i].ToString();
                //    SqReader[i] = cmd[i].ExecuteReader();


                //    while (SqReader[i].Read())
                //    {

                //        Stopwatch TestTime1 = new Stopwatch();
                //        TestTime1.Restart();
                //        TestTime1.Start();


                //        object[] values = new object[SqReader[i].FieldCount];
                //        SqReader[i].GetValues(values);

                //        for (int j = 0; j < Count; j++)
                //        {
                //            Apped.Append(values[j] + ",");
                //        }

                //        csv.Write(Apped.ToString());
                //        Apped.Clear();

                //    }
                //    SqReader[i].Close();

                //    stringA[i].Clear();
                //    cmd[i].CommandText = "";



                //}
                #endregion

                csv.Write_Close();

                ThreadFlags[i].Set();


            }
            public void Get_From_Db_Data_for_Anly_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public int Get_Sample_Count(int DB, string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);



                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[0].Tables[0].Rows)
                //{
                //    Value[count] = dr.ItemArray[0];
                //    count++;
                //}

                //   sqlAdapter[0].Dispose();
                //   cmd[0].Dispose();

                //   conn[0].Dispose();

                //sqlAdapter[0].Dispose();
                //stringA[0].Clear();

                //   cmd[0].Dispose();
                // conn[0].Close();


                //   int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Value.Length;
            }

            public void Get_From_Db_Ref_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;

                int count = 0;

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    count += Data.Per_DB_Column_Count[i];
                }


                this.Data.Reference_Header = new string[count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    Get_From_Db_Ref_Header_Thread(i);
                    //  ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data_for_Anly_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }



            }
            public void Get_From_Db_Ref_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                //  int Count_Data = 0;
                int count = 0;

                stringA[i].Clear();

                stringA[i].Append("Select * from REFHEADER");


                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                count = 0;

                while (SqReader[i].Read())
                {

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    int ForCount = 0;

                    ForCount = values.Length - 5;

                    for (int j = 0; j < ForCount; j++)
                    {
                        this.Data.Reference_Header[this.Data.DB_Column_Limit * i + j] = Convert.ToString(values[j]);

                    }



                    double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                    count++;
                }
                SqReader[i].Close();
                cmd[i].Dispose();
                stringA[i].Clear();
      


                ThreadFlags[i].Set();


            }

            public int Get_Column_Count(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                }

                sqlAdapter[0].Dispose();
                cmd[0].Dispose();

                int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Data_Count[0];
            }

            public void Close(Data_Class.Data_Editing.INT Data_Edit)
            {

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(close_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    tran[i].Dispose();
                    cmd[i].Dispose();
                    conn[i].Dispose();
                    // conn[i].Close();
                    sqlAdapter[i].Dispose();

                }
            }
            public void close_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                cmd[i].CommandText = "vacuum";
                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Read_Dispose(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }

            public void Set_Conn(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }

            public void trans(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;

                tran = new SQLiteTransaction[Data_Edit.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Tran_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Tran_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                cmd[i].Dispose();
                conn[i].Dispose();

                conn[i] = new SQLiteConnection(strConn[i]);
                cmd[i] = new SQLiteCommand(conn[i]);
                conn[i].Open();


                tran[i] = conn[i].BeginTransaction();
                cmd[i].Transaction = tran[i];

                ThreadFlags[i].Set();
            }

            public void Commit(Data_Class.Data_Editing.INT Data_Edit)
            {

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Commit_thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
            }
            public void Commit_thread(Object threadContext)
            {
                int i = (int)threadContext;
                tran[i].Commit();
                ThreadFlags[i].Set();
            }
            public void STDEVandMedian(List<double[]> Ds, int DB, int RowCount)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                double[][] ReturnValue = new double[Data.Per_DB_Column_Count[DB]][];

                for (int i = 0; i < Data.Per_DB_Column_Count[DB]; i++)
                {
                    ReturnValue[i] = new double[RowCount];
                }
                double dummytesttime1 = TestTime1.Elapsed.TotalMilliseconds;
                int j = 0;
                int k = 0;


                foreach (double[] o in Ds)
                {
                    var t = o;
                    for (int q = 0; q < t.Length - 2; q++)
                    {
                        ReturnValue[j][k] = t[q];
                        j++;
                    }

                    j = 0;
                    k++;
                }

                int Para_Count = 0;

                for (int i = 0; i < ReturnValue.Length; i++)
                {
                    double average = ReturnValue[i].Average();
                    double Median = 0f;

                    if (ReturnValue[i].Length % 2 == 0)
                    {
                        Array.Sort(ReturnValue[i]);

                        double dummyi = ReturnValue[i][(ReturnValue[i].Length / 2) - 1];
                        double dummyj = ReturnValue[i][ReturnValue[i].Length / 2];
                        Median = (dummyi + dummyj) / 2;
                    }
                    else
                    {
                        Array.Sort(ReturnValue[i]);
                        int GetMedian_i = (ReturnValue[i].Length) / 2;
                        Median = ReturnValue[i][GetMedian_i];
                    }

                    double minusSquareSummary = 0.0;

                    foreach (double source in ReturnValue[i])
                    {
                        minusSquareSummary += (source - average) * (source - average);
                    }

                    double stdev = Math.Sqrt(minusSquareSummary / (ReturnValue[i].Length - 1));

                    for (int q = 0; i < For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].CPK.Length; i++)
                    {

                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Std[q] = stdev;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Median_Data[q] = Median;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Min_Data[q] = ReturnValue[i].Min();
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Max_Data[q] = ReturnValue[i].Max();
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Avg[q] = ReturnValue[i].Average();

                    }
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_CPK = (average - ReturnValue[i].Min()) / (3 * stdev);
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_CPK = (average - ReturnValue[i].Max()) / (3 * stdev);

                    Para_Count++;

                }
                double dummytesttime2 = TestTime1.Elapsed.TotalMilliseconds;

            }
            static double[] STDEVandMedian(DataSet ds)
            {
                List<double> DataSet_Values = new List<double>();
                double[] ReturnValue = new double[2];

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    DataSet_Values.Add(Convert.ToDouble(dr.ItemArray[0]));
                }

                double average = DataSet_Values.Average();
                double Median = 0f;

                if (DataSet_Values.Count % 2 == 0)
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = DataSet_Values.Count / 2;
                    Median = DataSet_Values[GetMedian_i];
                }
                else
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = (DataSet_Values.Count + 1) / 2;
                    Median = DataSet_Values[GetMedian_i];
                }

                double minusSquareSummary = 0.0;

                foreach (double source in DataSet_Values)
                {
                    minusSquareSummary += (source - average) * (source - average);
                }

                double stdev = Math.Sqrt(minusSquareSummary / (DataSet_Values.Count - 1));

                ReturnValue[0] = stdev; ReturnValue[1] = Median;

                return ReturnValue;
            }

            public string Get_Data_From_Table(string Table, string header)
            {

                return "";
            }



        }

        public class MERGE_S4PD : INT
        {
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ReaderWriterLockSlim[] sqlitelock { get; set; }
            public string[] strConn { get; set; }
            public SQLiteConnection[] conn { get; set; }
            public SQLiteCommand[] cmd { get; set; }

            public SQLiteDataAdapter[] sqlAdapter { get; set; }
            public SQLiteCommandBuilder[] sqlcmdbuilder { get; set; }
            public SQLiteDataReader[] SqReader { get; set; }

            public DbDataReader[] DbReader { get; set; }
            public DataSet[] ds { get; set; }
            public DataTable dt_test { get; set; }
            public DataTable[] dt { get; set; }
            public SQLiteTransaction[] tran { get; set; }

            public ManualResetEvent[] ThreadFlags { get; set; }
            public ManualResetEvent[] Insert_ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }


            public string FilePath { get; set; }
            public string RefHeader { get; set; }
            public bool[] Wait { get; set; }

            public int Limit { get; set; }
            public int Limit_Count { get; set; }
            public int Table_Count { get; set; }
            public bool[] Insert_Thread_Wait { get; set; }
            public double[] Testtime { get; set; }

            public object[] ID { get; set; }
            public object[] Value { get; set; }
            public object[] WAFER_ID { get; set; }
            public object[] LOT_ID { get; set; }
            public object[] SITE_ID { get; set; }
            public Dictionary<string, double[]> Selected_Parameter_Distribution { get; set; }
            public double[][] test { get; set; }

            double[] Testtime1 { get; set; }
            double[] Testtime2 { get; set; }
            double[] Testtime3 { get; set; }
            public string[][] Teststring { get; set; }
            public double[][] Testdouble { get; set; }


            public Dictionary<string, IQR> DIC_IQR { get; set; }
            public object[] Variation { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test { get; set; }
            public List<List<RowAndPass>[]>[] Yield_Test_New_Spec { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent { get; set; }
            public List<List<int>>[] For_Any_Yield { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_Lot { get; set; }
            public List<List<List<int>>>[] For_Any_Yield_For_SITE { get; set; }

            public List<List<int>[]>[] ForCampare_Yield { get; set; }
            public List<List<int>[]>[] For_Any_Yield_Percent_For_New_Spec { get; set; }
            public List<List<int>>[] For_Any_Yield_For_New_Spec { get; set; }
            public List<List<int>[]>[] For_New_Spec_ForCampare_Yield { get; set; }
            public List<int[]>[] ForCampare_Yield_Fro_DB { get; set; }
            public List<List<int[]>>[] ForCampare_Yield_Fro_DB_List { get; set; }
            public List<List<List<List<int>[]>>>[] ForCampare_Yield_DB_LotVariation { get; set; }
            public List<List<int>>[] For_New_Spec_ForCampare_Yield2 { get; set; }
            public List<List<List<int[]>>>[] ForCampare_Yield_Fro_DB_List_LotVariation { get; set; }
            public Dictionary<string, int> Refer_Site_And_Num { get; set; }
            public Dictionary<string, int> Refer_Lot_And_Num { get; set; }
            public List<int>[] ForCampare_Yield_List { get; set; }
            public List<List<int>[]> ForCampare_Yield_List1 { get; set; }
            public List<List<int>[]>[] ForCampare_Yield_List2 { get; set; }
            public Dictionary<string, Values> Values { get; set; }
            public Dictionary<string, Data_Calculation> Cal_Value_by_rowsdata { get; set; }
            public Dictionary<string, Data_Calculation> For_New_Spec_Cal_Value_by_rowsdata { get; set; }
            public List<int>[] Check { get; set; }
            public List<List<int>[]> Test { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public List<double[]>[] DB_DataSet_Values { get; set; }

            public Dictionary<string, int> Lot_Dic { get; set; }
            public Dictionary<string, int> Site_Dic { get; set; }
            public Dictionary<string, int> Bin_Dic { get; set; }
            public Dictionary<string, Dictionary<string, List<string>>> Matching_Lots { get; set; }
            public Dictionary<string, List<string>> Matching_Lot { get; set; }
            public Stopwatch[] TestTime1 { get; set; }
            public Stopwatch[] TestTime2 { get; set; }
            public Stopwatch[] TestTime3 { get; set; }
            public Stopwatch[] TestTime4 { get; set; }
            public Stopwatch[] TestTime5 { get; set; }
            public long SampleCount { get; set; }
            public object Update_Data_ID { get; set; }
            public string[] Update_Datas_ID { get; set; }
            public string Get_Gross_Para { get; set; }

            public double Get_Gross_Persent { get; set; }
            public string Get_Gross_Selector { get; set; }
            public object[] Std_Value { get; set; }
            public double[] Std_Value_Convert { get; set; }
            public List<Dictionary<string, Gross>[]> List_Gross_Values { get; set; }
            public Dictionary<string, Gross>[] Gross_Values1 { get; set; }
            public long NB { get; set; }


            public Dictionary<string, CSV_Class.For_Box>[] Dic_Test { get; set; }
            public Dictionary<string, CSV_Class.For_Box> Dic_Test_For_Spec_Gen { get; set; }
            public string Table { get; set; }
            public string Filename { get; set; }

            public double[] Make_New_Spec_For_Yield_Min { get; set; }
            public double[] Make_New_Spec_For_Yield_Max { get; set; }
            public List<string> Gross { get; set; }
            public List<string[]>[] DataSet_Value { get; set; }
            public List<double[]>[] DataSet_Double_Value { get; set; }
            public string Lot_ID { get; set; }
            public string SubLot_ID { get; set; }
            public string Tester_ID { get; set; }
            public string Site { get; set; }
            public string Bin { get; set; }
            public string ID_Unit { get; set; }

            public int Bin_place { get; set; }

            public string Query { get; set; }
            public string Query2 { get; set; }
            public string CellID { get; set; }
            public bool _From_Db { get; set; }
            public int Spec_Table_Count { get; set; }
            public bool _Flag { get; set; }
            public bool _SUBLOT_Flag { get; set; }
            public bool Clotho_Spec_Flag { get; set; }
            public string Before_Lot_ID { get; set; }
            public string Changed_Lot_ID { get; set; }

            public int[] Each_Thread_Count { get; set; }

            public string[] No_Index { get; set; }
            public string[] Paraname { get; set; }
            public string[] SpecMin { get; set; }
            public string[] SpecMax { get; set; }
            public string[] DataMin { get; set; }
            public string[] DataMedian { get; set; }
            public string[] DataMax { get; set; }
            public string[] CPK { get; set; }
            public string[] STD { get; set; }
            public string[] Percent { get; set; }
            public string[] Fail { get; set; }

            public string[] Line { get; set; }

            public int Count_Current_Setting { get; set; }

            public void Open_DB(string FileName, Data_Class.Data_Editing.INT Data_Edit)
            {
                string Filename = FileName.Substring(FileName.LastIndexOf("\\") + 1);
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "\\" + Filename.Substring(0, FileName.Length - 4) + "_" + i + ".db";
                    //  strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    cmd[i].ExecuteNonQuery();

                }


                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                dt = new DataTable[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    dt[i] = new DataTable();
                    cmd[i].CommandText = "PRAGMA synchronous";
                    SqReader[i] = cmd[i].ExecuteReader();
                    dt[i].Load(SqReader[i]);
                }



            }

            public void Open_DB(string[] FileName, Data_Class.Data_Editing.INT Data_Edit)
            {
                Data_Edit.DB_Count = FileName.Length;
                strConn = new string[Data_Edit.DB_Count];
                conn = new SQLiteConnection[Data_Edit.DB_Count];
                cmd = new SQLiteCommand[Data_Edit.DB_Count];
                tran = new SQLiteTransaction[Data_Edit.DB_Count];
                stringA = new StringBuilder[Data_Edit.DB_Count];
                TestTime1 = new Stopwatch[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                sqlAdapter = new SQLiteDataAdapter[Data_Edit.DB_Count];
                SqReader = new SQLiteDataReader[Data_Edit.DB_Count];
                DbReader = new DbDataReader[Data_Edit.DB_Count];
                ds = new DataSet[Data_Edit.DB_Count];


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    string Filename = FileName[i].Substring(FileName[i].LastIndexOf("\\") + 1);

                    int length = Filename.Length;
                    Filename = Filename.Substring(0, length - 5);

                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "\\" + Filename + i + ".db";
                    strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + ".csv\\" + Filename.Substring(0, Filename.Length) + "_" + i + ".db";
                    //strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA TEMP_STORE = FILE; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SCHEMA.SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA SCHEMA.PAGE_SIZE = 4096; PRAGMA SCHEMA.MAX_PAGE_COUNT = 1073741823; PRAGMA SCHEMA.JOURNAL_MODE = WAL; PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = FALSE; PRAGMA CHECKPOINT_FULLFSYNC = FALSE;  PRAGMA SCHEMA.AUTO_VACCUM = 0; AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; Version = 3;";
                    // strConn[i] = @"Data Source = C:\\Automation\\DB\\YIELD\\" + Filename + "_" + i + ".db; PRAGMA threads = 7; PRAGMA LOCKING_MODE = RESERVED; DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = 2; Cache_size = 10000000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 10000000;PRAGMA journal_mode = WAL;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    // strConn[i] = @"Data Source = MEMORY" + i + ".db;  DEBUG = 1;Version = 3;cache = shared;strict = on;PRAGAM read_uncommitted = true; PRAGMA synchronous=off; PRAGMA temp_store = memory; Cache_size = 89810000;PRAGMA page_sige = 4096; PRAGMA default_cache_size = 100000;PRAGMA journal_mode = MEMORY;PRAGMA count_changes=OFF;PRAGMA Column = 2000;";
                    conn[i] = new SQLiteConnection(strConn[i]);
                    cmd[i] = new SQLiteCommand(conn[i]);
                    stringA[i] = new StringBuilder();
                    TestTime1[i] = new Stopwatch();
                    sqlAdapter[i] = new SQLiteDataAdapter();
                    ds[i] = new DataSet();
                    conn[i].Open();
                    //cmd[i].CommandText = "PRAGMA JOURNAL_MODE = PERSIST; PRAGMA JOURNAL_SIZE_LIMIT = -1; PRAGMA default_cache_size = 10000000; PRAGMA count_changes=OFF; PRAGMA TEMP_STORE = MEMORY; PRAGMA WAL_AUTOCHECKPOINT = 1000; PRAGMA synchronous = off; PRAGMA SECURE_DELETE = FALSE; PRAGMA RECURSIVE_TRIGGERS = FALSE; PRAGMA PAGE_SIZE = 4096; PRAGMA MAX_PAGE_COUNT = 1073741823;  PRAGMA IGNORE_CHECK_CONSTRAINTS = FALSE; PRAGMA FOREIGN_KEYS = true; PRAGMA CHECKPOINT_FULLFSYNC = FALSE; PRAGMA AUTO_VACCUM = 1; PRAGMA AUTOMATIC_INDEX = FALSE; PRAGMA LOCKING_MODE = EXCLUSIVE; PRAGMA Version = 3; ";
                    //cmd[i].ExecuteNonQuery();

                }

            }

            public void DropTable(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {

                try
                {
                    for (int i = 0; i < Data_Edit.DB_Count; i++)
                    {
                        cmd[i].CommandText = "";
                        cmd[i].CommandText = "drop TABLE " + Query;
                        cmd[i].ExecuteNonQuery();
                    }
                }
                catch { }
            }

            public void Insert_Header(Data_Class.Data_Editing.INT Data_Edit)
            {



                Lot_ID = Lot_ID.Replace('-', '_');
                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < 1; i++)
                {
                    stringA[i].Clear();
                    MakecolumnsThread(i);

                }

               
                //if (Clotho_Spec_Flag)
                //{

                //    Data = Data_Edit;
                //    ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                //    Wait = new bool[Data_Edit.DB_Count];
                //    Wait = new bool[Data_Edit.DB_Count];
                //    Testtime = new double[Data_Edit.DB_Count];

                //    Data.Data_Table = "Clotho_Spec";
                //    for (int i = 0; i < Data_Edit.DB_Count; i++)
                //    {
                //        stringA[i].Clear();
                //        ThreadFlags[i] = new ManualResetEvent(false);
                //        ThreadPool.QueueUserWorkItem(new WaitCallback(MakecolumnsThread1), i);
                //    }

                //    for (int i = 0; i < Data_Edit.DB_Count; i++)
                //    {
                //        Wait[i] = ThreadFlags[i].WaitOne();
                //        stringA[i] = new StringBuilder();
                //    }
                //}
         
            }

            public void MakecolumnsThread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Ref_New_Header.Length;
                cmd[i] = new SQLiteCommand(conn[i]);

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {

                        stringA[i].Append("CREATE TABLE " + Table + "(" + Data.Ref_New_Header[j] + " VARCAHR(20)");

                    }
                    else
                    {

                        stringA[i].Append(" " + Data.Ref_New_Header[j] + " VARCHAR(20)");

                    }

                    if (j == Count - 1)
                    {

                        stringA[i].Append(", SubLot VARCAHR(5), id VARCAHR(5), LOTID VARCAHR(20), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");

                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }

            }

            public void MakecolumnsThread1(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE " + Data.Data_Table + "(" + Data.New_Header[0] + " VARCAHR(20)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE " + Data.Data_Table + "(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(20)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(20)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        //    if (_SUBLOT_Flag == true)
                        //    {
                        //        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");

                        //    }
                        //    else
                        //    {
                        stringA[i].Append(", SubLot VARCAHR(5), id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), FAIL VARCHAR(20), BIN VARCHAR(20));");
                        //   }

                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }
            public void Insert_Current_Setting(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }
            public void Insert_Current_Setting_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {
                Data = Data_Edit;
                this.Table = Table;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    //  cmd[i].Reset();
                    //    ThreadFlags[i] = new ManualResetEvent(false);
                    Insert_Current_Setting_Data_Thread(i);
                    //  ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //       Wait[i] = ThreadFlags[i].WaitOne();
                }

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    cmd[i].Reset();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Clotho_Spec_Max_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}

            }

            public void Insert_Current_Setting_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();
                stringA[i].Clear();
                SampleCount = 1;

                cmd[i] = new SQLiteCommand(conn[i]);

                int Count = Data.Per_DB_Column_Count[i];


                int k = 0;


                if (Table.ToUpper() == "CLOTHO_SPEC")
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Min[Spec_Count] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Min[Spec_Count] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Min[Spec_Count] + "',");

                            }


                            stringA[i].Append("'0','" + Spec_Count + "','0','0', '0', '0');");


                            cmd[i].CommandText = stringA[i].ToString();

                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                    }




                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;
                    for (int Spec_Count = 0; Spec_Count < Data.Clotho_Spcc_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Clotho_Spcc_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }
                else
                {
                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Min[0] + "',");

                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }

                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Min[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Min[0] + "',");

                            }

                            stringA[i].Append("'1', '" + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                    Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                    stringA[i].Clear();
                    cmd[i].Reset();
                    k = 0;
                    SampleCount = 2;

                    for (int Spec_Count = 0; Spec_Count < Data.Customor_Clotho_List[0].Min.Length; Spec_Count++)
                    {
                        if (i == 0)
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[0].Max[0] + "',");
                            for (int p = 0; p < 9; p++)
                            {
                                stringA[i].Append("'" + p + "',");
                            }
                            for (k = 10; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }
                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }
                        else
                        {
                            stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i - 9].Max[0] + "',");

                            for (k = 1; k < Count; k++)
                            {

                                stringA[i].Append("'" + Data.Customor_Clotho_List[Data.DB_Column_Limit * i + k - 9].Max[0] + "',");

                            }

                            string Test = Convert.ToString(Spec_Count) + Convert.ToString(Spec_Count);

                            stringA[i].Append("'1', '" + Data.Clotho_Spcc_List[0].Min.Length + Spec_Count + "', '1', '1', '1', '1');");

                            cmd[i].CommandText = stringA[i].ToString();
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                }




                //   ThreadFlags[i].Set();
            }
            public void Insert_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE spec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE spec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY );");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_New_Spec_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data = Data_Edit;
                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_New_Spec_Header_Thread), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void Insert_New_Spec_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE newspec(" + Data.New_Header[0] + " VARCAHR(5)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE newspec(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(5)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(5)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", id VARCAHR(5) PRIMARY KEY );");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Insert_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                // sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];
                sqlitelock = new ReaderWriterLockSlim[Data.DB_Count];
                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                //Testdouble = new double[7][];

                //Testdouble[0] = new double[Data.DB_Column_Limit];
                //Testdouble[1] = new double[Data.DB_Column_Limit];
                //Testdouble[2] = new double[Data.DB_Column_Limit];
                //Testdouble[3] = new double[Data.DB_Column_Limit];
                //Testdouble[4] = new double[Data.DB_Column_Limit];
                //Testdouble[5] = new double[Data.DB_Column_Limit];
                //Testdouble[6] = new double[Data.Per_DB_Column_Count[6]];
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder(100000);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder(100000);
                    Testtime[i] = TestTime1.Elapsed.TotalMilliseconds;
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);
            }
            public void Insert_Ref_Header_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                Data.Data_Table = "REFHEADER";

                ThreadFlags = new ManualResetEvent[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                stringA = new StringBuilder[Data.DB_Count];
                // sqlAdapter = new SQLiteDataAdapter[Data.DB_Count];
                Wait = new bool[Data.DB_Count];
                Testtime = new double[Data.DB_Count];
                sqlitelock = new ReaderWriterLockSlim[Data.DB_Count];


                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //sqlAdapter[i] = new SQLiteDataAdapter();
                    stringA[i] = new StringBuilder(100000);
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Ref_Header_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder(100000);
                    Testtime[i] = TestTime1.Elapsed.TotalMilliseconds;
                }

            }

            public void Insert_Ref_Header_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Reference_Header[0] + "',");
                    // stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[0].Replace("PID-", "") + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Reference_Header[(Data.DB_Column_Limit * i)] + "',");

                }

                for (k = 1; k < Count; k++)
                {
                    stringA[i].Append("'" + Data.Reference_Header[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.Reference_Header[(Data.DB_Column_Limit * i) + k] + "', '" + SubLot_ID + "', '" + SampleCount + "' , '" + Lot_ID + "' , '" + Site + "' ,'0');");

                // stringA[i].Append("'" + SubLot_ID + "', '" + SampleCount + "' , '0');");

                // stringA[i].Append("'1, '" + SubLot_ID + "', '" + SampleCount + "' , '0');");
                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void Insert_Data(long Sample)
            {
                SampleCount = Sample;

                Lot_ID = Lot_ID.Replace('-', '_');
                //ForCampare_Yield_List = new List<int>[Data.DB_Count];

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    ForCampare_Yield_List[i] = new List<int>();
                //}

                //for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                //{
                //    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                //    {
                //        ForCampare_Yield_List[i].Add(0);
                //    }
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    Insert_Data_NoThread(i);

                //}

                for (int i = 0; i < 1; i++)
                {
                    stringA[i].Clear();
                  //  ThreadFlags[i] = new ManualResetEvent(false);

                    Insert_Data_Thread(i);

                 //   ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Data_Thread), i);
                }

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}

                //   ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Insert_Data_NoThread(int DB)
            {
                int i = DB;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[0].Replace("PID-", "") + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Data.Data_Table + " VALUES ('" + Data.Getstring[(Data.DB_Column_Limit * i)] + "',");

                }

                for (k = 1; k < Count; k++)
                {
                    stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.Getstring[(Data.DB_Column_Limit * i) + k] + "', '" + SubLot_ID + "', '" + SampleCount + "' , '0');");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Getstring.Length;
                TestTime1[i].Restart();
                TestTime1[i].Start();

                stringA[i].Clear();
                int k = 0;

                cmd[i] = new SQLiteCommand(conn[i]);



                stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.Getstring[0] + "',");

                for (k = 1; k < Count; k++)
                {
                    stringA[i].Append("'" + Data.Getstring[k] + "',");

                }

                stringA[i].Append("'"  + SubLot_ID + "', '" + ID_Unit + "' , '" + Lot_ID + "' , '" + Site + "' ,'0', '" + Bin + "');");
                // stringA[i].Append("'" + Data.Getstring[k] + "', '" + SubLot_ID + "', '" + ID_Unit + "' , '" + Lot_ID + "' , '" + Site + "' ,'0', '" + Bin + "');");

                // stringA[i].Append("'" + SubLot_ID + "', '" + SampleCount + "' , '0');");

                // stringA[i].Append("'1, '" + SubLot_ID + "', '" + SampleCount + "' , '0');");
                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
               // ThreadFlags[i].Set();
            }

            public void Insert_Data_Get_From_DB(int Sample)
            {

            }
            public void Insert_Spec_Get_From_DB(Data_Class.Data_Editing.INT Data_Edit)
            {


                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Get_From_DB_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }

            public void Insert_Spec_Get_From_DB_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }
            public void Insert_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Insert_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;
                int k = 0;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[0] + "',");

                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_LowSpec[Data.DB_Column_Limit * i] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_LowSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_LowSpec[Data.DB_Column_Limit * i + k] + "', '0', '0' ,'0','0', '0', '0');");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                stringA[i].Clear();
                cmd[i].Reset();
                k = 0;
                SampleCount = 2;

                if (i == 0)
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[0] + "',");
                }
                else
                {
                    stringA[i].Append("INSERT INTO " + Table + " VALUES ('" + Data.New_HighSpec[Data.DB_Column_Limit * i] + "',");
                }

                for (k = 1; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    stringA[i].Append("'" + Data.New_HighSpec[(Data.DB_Column_Limit * i) + k] + "',");

                }

                stringA[i].Append("'" + Data.New_HighSpec[Data.DB_Column_Limit * i + k] + "', '1', '1' , '1', '1', '1', '1');");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Insert_Spec_Data(Data_Class.Data_Editing.INT Data_Edit, string Table)
            {

            }


            public void Insert_Files_Name(string Filename)
            {

                this.Filename = Filename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(thread_Insert_File_Name), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }

            public void thread_Insert_File_Name(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();


                stringA[i].Append("INSERT INTO Files VALUES ('" + this.Filename + "');");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Make_table(string Tablename)
            {
                stringA[0].Clear();
                stringA[0].Append("CREATE TABLE " + Tablename + "( FIRST VARCAHR(5), END VARCAHR(5), DBCOUNT VARCHAR(5), COLUMNCOUNT VARCHAR(5) );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";

                stringA[0].Clear();
                stringA[0].Append("INSERT INTO INF VALUES ('" + TheFirst_Trashes_Header_Count + "' , '" + TheEnd_Trashes_Header_Count + "' , '" + Data.Per_DB_Column_Count.Length + "' , '" + Data.Per_DB_Column_Count[Data.Per_DB_Column_Count.Length - 1] + "' );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";
            }

            public void Make_table2(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {
                Data = Data_Edit;
                Table = Tablename;

                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(_Make_Table), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void _Make_Table(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                for (int j = 0; j < Count; j++)
                {
                    if (j == 0)
                    {
                        if (i == 0)
                        {
                            //stringA[i].Append("CREATE TABLE data(" + Data.New_Header[0] + " real");
                            stringA[i].Append("CREATE TABLE " + Table + "(" + Data.New_Header[0] + " VARCAHR(20)");
                            // Teststring[i][0] = Data.New_Header[0];
                        }
                        else
                        {
                            // stringA[i].Append("CREATE TABLE data(" + Data.New_Header[Data.DB_Column_Limit * i] + " real");
                            stringA[i].Append("CREATE TABLE " + Table + "(" + Data.New_Header[Data.DB_Column_Limit * i] + " VARCAHR(20)");
                            //  Teststring[i][0] = Data.New_Header[Data.DB_Column_Limit * i];
                        }

                    }
                    else
                    {
                        // stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " real");
                        stringA[i].Append(" " + Data.New_Header[Data.DB_Column_Limit * i + j] + " VARCHAR(20)");
                        // Teststring[i][j] = Data.New_Header[Data.DB_Column_Limit * i + j];
                    }

                    if (j == Count - 1)
                    {
                        stringA[i].Append(", SubLot VARCAHR(5), id VARCAHR(5) PRIMARY KEY, LOTID VARCAHR(5), SITEID VARCAHR(5), Fail VARCHAR(20));");
                        //  stringA[i].Append(", id INTEGER PRIMARY KEY AUTOINCREMENT);");
                        cmd[i].CommandText = stringA[i].ToString();
                        cmd[i].ExecuteNonQuery();
                        cmd[i].CommandText = "";
                    }
                    stringA[i].Append(",");
                }
                ThreadFlags[i].Set();
            }

            public void Make_table_For_Filename(Data_Class.Data_Editing.INT Data_Edit, string Tablename)
            {
                Data = Data_Edit;
                Table = Tablename;

                ThreadFlags = new ManualResetEvent[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Wait = new bool[Data_Edit.DB_Count];
                Testtime = new double[Data_Edit.DB_Count];

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(_Make_Table_For_Filename), i);
                }

                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                    stringA[i] = new StringBuilder();

                }
            }

            public void _Make_Table_For_Filename(Object threadContext)
            {
                int i = (int)threadContext;

                stringA[i].Append("CREATE TABLE " + Table + "(FIle VARCAHR(20))");


                cmd[i].CommandText = stringA[i].ToString();
                cmd[i].ExecuteNonQuery();
                cmd[i].CommandText = "";

                stringA[i].Append(",");

                ThreadFlags[i].Set();
            }

            public void Make_table_For_Trace(string Tablename,string Chan, bool Flag)
            {
                stringA[0].Clear();
                stringA[0].Append("CREATE TABLE " + Tablename + "( FIRST VARCAHR(5), END VARCAHR(5), DBCOUNT VARCHAR(5), COLUMNCOUNT VARCHAR(5) );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";

                stringA[0].Clear();
                stringA[0].Append("INSERT INTO INF VALUES ('" + TheFirst_Trashes_Header_Count + "' , '" + TheEnd_Trashes_Header_Count + "' , '" + Data.Per_DB_Column_Count.Length + "' , '" + Data.Per_DB_Column_Count[Data.Per_DB_Column_Count.Length - 1] + "' );");
                cmd[0].CommandText = stringA[0].ToString();
                cmd[0].ExecuteNonQuery();
                cmd[0].CommandText = "";
            }

            public void Delete_Spec_Data(string Tablename)
            {

                Table = Tablename;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Delete_Spec_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Delete_Spec_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                SampleCount = 1;
                int k = 0;


                stringA[i].Append("Delete from " + Table + " where id = 0");


                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;


                stringA[i].Clear();

                stringA[i].Append("Delete from " + Table + " where id = 1");

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Delete_Lot_Data(string Query)
            {
                this.Query = Query;

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    cmd[i].Reset();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Delete_Lot_Data_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
            }
            public void Delete_Lot_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                stringA[i].Append(this.Query);

                cmd[i].CommandText = stringA[i].ToString();

                cmd[i].ExecuteNonQuery();

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                ThreadFlags[i].Set();
            }

            public void Save_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                //Update_Data_ID = data;

                //if (data != null)
                //{
                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        ThreadFlags[i] = new ManualResetEvent(false);
                //        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                //    }

                //    for (int i = 0; i < Data.DB_Count; i++)
                //    {
                //        Wait[i] = ThreadFlags[i].WaitOne();
                //    }
                //}

            }
            public void Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    //  Insert_table_Data_Thread(i);
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Insert_table_Data_Thread), i);
                //}

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}


            }

            public void Road_Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit)
            {
                // SampleCount = Sample;

                ForCampare_Yield_List = new List<int>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ForCampare_Yield_List[i] = new List<int>();
                }

                for (int i = 0; i < ForCampare_Yield_List.Length; i++)
                {
                    for (int j = 0; j < Data.Per_DB_Column_Count[i]; j++)
                    {
                        ForCampare_Yield_List[i].Add(0);
                    }
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Road_Save_Customer_Spec_table_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

                ForCampare_Yield_List1.Add(ForCampare_Yield_List);

                Insert_ThreadFlags[0].Set();
            }
            public void Road_Save_Customer_Spec_table_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i] - 1;
                TestTime1[i].Restart();
                TestTime1[i].Start();


                int k = 0;

                if (i == 0)
                {
                    ForCampare_Yield_List[0][0] = 0;
                }
                else
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i] < Convert.ToDouble(DataSet_Value[i][0][0]) || Data.New_LowSpec[Data.DB_Column_Limit * i] > Convert.ToDouble(DataSet_Value[i][0][0]))
                    {
                        ForCampare_Yield_List[i][0] = 1;
                    }
                }

                for (k = 1; k < Count; k++)
                {
                    if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][k]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][k]))
                    {
                        ForCampare_Yield_List[i][k] = 1;
                    }

                }

                if (Data.New_HighSpec[Data.DB_Column_Limit * i + k] < Convert.ToDouble(DataSet_Value[i][0][Count]) || Data.New_LowSpec[Data.DB_Column_Limit * i + k] > Convert.ToDouble(DataSet_Value[i][0][Count]))
                {
                    ForCampare_Yield_List[i][Data.Per_DB_Column_Count[i] - 1] = 1;
                }


                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;

                stringA[i].Clear();
                ThreadFlags[i].Set();
            }

            public void LOTID_Update(string Query, string Query2, string CellID)
            {

                this.Query = Query;
                this.Query2 = Query2;
                this.CellID = CellID;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ThreadFlags[i] = new ManualResetEvent(false);
                    //   LOTID_Update_Thread(i);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(LOTID_Update_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }
            public void LOTID_Update_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                if (this.CellID == "DIE_X")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }

                }
                else if (this.CellID == "DIE_Y")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "TIME")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "TOTAL_TESTS")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "WAFER_ID")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();
                    }
                }
                else if (this.CellID == "LOTID")
                {
                    if (i == 0)
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();

                        if (this.Query2 != null)
                        {
                            cmd[i].CommandText = this.Query2;
                            cmd[i].ExecuteNonQuery();
                            stringA[i].Clear();
                        }

                    }
                    else
                    {
                        cmd[i].CommandText = Query;
                        cmd[i].ExecuteNonQuery();
                        stringA[i].Clear();

                    }
                }




                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }


            public void Gross_Update_Data(object data)
            {
                Update_Data_ID = data;

                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Data_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }

            }
            public void Gross_Update_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (object o in (Array)Update_Data_ID)
                {
                    cmd[i].CommandText = "Update data set FAIL = '1'  where id = " + o.ToString();
                    cmd[i].ExecuteNonQuery();
                    stringA[i].Clear();
                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Gross_Update_Datas(List<string> data)
            {
                Update_Datas_ID = data.ToArray();
                if (data != null)
                {
                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        ThreadFlags[i] = new ManualResetEvent(false);
                        ThreadPool.QueueUserWorkItem(new WaitCallback(Gross_Update_Datas_Thread), i);
                    }

                    for (int i = 0; i < Data.DB_Count; i++)
                    {
                        Wait[i] = ThreadFlags[i].WaitOne();
                    }
                }
            }
            public void Gross_Update_Datas_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                TestTime1[i].Restart();
                TestTime1[i].Start();

                foreach (object o in (Array)Update_Datas_ID)
                {
                    cmd[i].CommandText = "Update data set FAIL = '1'  where id = " + o.ToString();
                    cmd[i].ExecuteNonQuery();
                    cmd[i].Reset();
                }

                Testtime[i] = TestTime1[i].Elapsed.TotalMilliseconds;
                ThreadFlags[i].Set();
            }
            public void Chnaged_Spec_Update_Data(int DB, int Index, string Parameter, double Spec, int GetId)
            {
                stringA[DB].Clear();
                stringA[DB].Append("Update newspec set " + Parameter + " = " + Spec + " where id = " + GetId);

                cmd[DB].CommandText = stringA[DB].ToString();

                cmd[DB].ExecuteNonQuery();
                cmd[DB].Reset();

                stringA[DB].Clear();
            }
            public Dictionary<string, double[]> Chnaged_Spec_Anl_Yield(int DB, int Index, string Parameter)
            {
                stringA[DB].Clear();
                Dictionary<string, double[]> Dic_Change_Spec = new Dictionary<string, double[]>();


                stringA[DB].Append("Select " + Parameter + " from newspec");

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                object[] GetSpec = new object[ds[DB].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    GetSpec[count] = dr.ItemArray[0];
                    count++;
                }

                double[] Toduble_Spec = Array.ConvertAll<object, double>(GetSpec, Convert.ToDouble);

                Dic_Change_Spec.Add("SPEC", Toduble_Spec);
                stringA[DB].Clear();

                stringA[DB].Append("Select " + Parameter + " from data where Fail not like '1'");

                cmd[DB].CommandText = stringA[DB].ToString();
                ds[DB] = new DataSet();

                sqlAdapter[DB].SelectCommand = cmd[DB];
                sqlAdapter[DB].Fill(ds[DB]);

                object[] GetData = new object[ds[DB].Tables[0].Rows.Count];
                count = 0;
                foreach (DataRow dr in ds[DB].Tables[0].Rows)
                {
                    GetData[count] = dr.ItemArray[0];
                    count++;
                }

                double[] Toduble_Data = Array.ConvertAll<object, double>(GetData, Convert.ToDouble);

                Dic_Change_Spec.Add("DATA", Toduble_Data);

                stringA[DB].Clear();

                return Dic_Change_Spec;
            }
            public void Get_Ave_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

                test = new double[Data.DB_Count][];
                double[] test1 = new double[Data.DB_Count];

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    test[i] = new double[10000];
                    Testtime[i] = new double();
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                    //ThreadFlags[i] = new ManualResetEvent(false);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //Wait[i] = ThreadFlags[i].WaitOne();
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                }


            }

            public void Get_Ave_Data_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                test = new double[Data.DB_Count][];
                double[] test1 = new double[Data.DB_Count];

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    test[i] = new double[10000];
                    Testtime[i] = new double();
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                    //ThreadFlags[i] = new ManualResetEvent(false);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //Wait[i] = ThreadFlags[i].WaitOne();
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                }


            }
            public void Set_Refer_for_Anlyzer(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_Ave_Data2(Data_Class.Data_Editing.INT Data_Edit)
            {

                test = new double[Data.DB_Count][];
                double[] test1 = new double[Data.DB_Count];

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    test[i] = new double[10000];
                    Testtime[i] = new double();
                    stringA[i].Clear();
                    Get_Ave_Data_Thread(i);
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                    //ThreadFlags[i] = new ManualResetEvent(false);
                    //ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Ave_Data_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    //Wait[i] = ThreadFlags[i].WaitOne();
                    test1[i] = TestTime1.Elapsed.TotalMilliseconds;
                }


            }
            public void Get_Ave_Data_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                stringA[i].Append("Select * from data where Fail not like '1'");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                List<double[]> DataSet_Values = new List<double[]>();
                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    double[] doubles = Array.ConvertAll<object, double>(values, Convert.ToDouble);
                    DataSet_Values.Add(doubles);

                    count++;

                }
                SqReader[i].Close();

                STDEVandMedian(DataSet_Values, i, count);

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].CommandText = "";
                ThreadFlags[i].Set();
            }

            public void Get_Saved_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                DataSet_Value = new List<string[]>[Data.DB_Count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    DataSet_Value[i] = new List<string[]>();
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Saved_Spec_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }

            }

            public void Get_Saved_Spec_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                cmd[i] = new SQLiteCommand(conn[i]);
                stringA[i].Append("Select * from Clotho_Spec");
                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                int count = 0;

                while (SqReader[i].Read())
                {
                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    DataSet_Value[i].Add(stringD);

                    count++;

                }
                SqReader[i].Close();

                double testtime = TestTime1.Elapsed.TotalMilliseconds;
                stringA[i].Clear();
                cmd[i].Dispose();
                ThreadFlags[i].Set();

            }

            public void Get_Rows_Data(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface)
            {
                //stringA[DB].Clear();
                //stringA[DB].Append("Select id, " + Select_Para + " from data");

                //cmd[DB].CommandText = stringA[DB].ToString();
                //ds[DB] = new DataSet();

                //sqlAdapter[DB].SelectCommand = cmd[DB];
                //sqlAdapter[DB].Fill(ds[DB]);

                //ID = new object[ds[DB].Tables[0].Rows.Count];
                //Value = new object[ds[DB].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //{
                //    ID[count] = dr.ItemArray[0];
                //    Value[count] = dr.ItemArray[1];

                //    count++;
                //}

                //double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);


                //stringA[DB].Clear();
            }

            public void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface, DataTable dt)
            {
                //stringA[DB].Clear();
                //stringA[DB].Append("Select id, " + Select_Para + " from data");

                //cmd[DB].CommandText = stringA[DB].ToString();
                //ds[DB] = new DataSet();

                //sqlAdapter[DB].SelectCommand = cmd[DB];
                //sqlAdapter[DB].Fill(ds[DB]);

                //ID = new object[ds[DB].Tables[0].Rows.Count];
                //Value = new object[ds[DB].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[DB].Tables[0].Rows)
                //{
                //    ID[count] = dr.ItemArray[0];
                //    Value[count] = dr.ItemArray[1];

                //    count++;
                //}

                //double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);


                //stringA[DB].Clear();
            }
            public void Get_Selected_Para(int DB, string Select_Para, bool Flag, string Selector)
            {


            }
            public double[] Get_Find_Bin(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                double[] doubles = Array.ConvertAll<object, double>(Value, Convert.ToDouble);
                sqlAdapter[0].Dispose();
                cmd[0].Dispose();
                stringA[0].Clear();
                return doubles;
            }

            public List<object[]> Get_Data_By_Querys(string Query)
            {
                return null;
            }
            public string[] Get_Data_By_Query(string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                // cmd[0].CommandText = stringA[0].ToString();
                // ds[0] = new DataSet();

                // sqlAdapter[0].SelectCommand = cmd[0];
                // sqlAdapter[0].Fill(ds[0]);

                // Value = new object[ds[0].Tables[0].Rows.Count];

                //// int count = 0;
                // foreach (DataRow dr in ds[0].Tables[0].Rows)
                // {
                //     Value[count] = dr.ItemArray[0];
                //     count++;
                // }

                //  string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);
                // SqReader[0] = cmd[0].ExecuteReader();
                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[500000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    Value1[count] = stringD[0];

                    count++;

                }

                Array.Resize(ref Value1, count);

                cmd[0].Dispose();
                SqReader[0].Close();

                string[] _string = Array.ConvertAll<object, string>(Value1, Convert.ToString);


                stringA[0].Clear();
                return _string;
            }


            public Dictionary<string,object[]> Get_Data_By_Query_S4PD(string Query, string Chan)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                // cmd[0].CommandText = stringA[0].ToString();
                // ds[0] = new DataSet();

                // sqlAdapter[0].SelectCommand = cmd[0];
                // sqlAdapter[0].Fill(ds[0]);

                // Value = new object[ds[0].Tables[0].Rows.Count];

                //// int count = 0;
                // foreach (DataRow dr in ds[0].Tables[0].Rows)
                // {
                //     Value[count] = dr.ItemArray[0];
                //     count++;
                // }

                //  string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);
                // SqReader[0] = cmd[0].ExecuteReader();
                cmd[0] = new SQLiteCommand(conn[0]);
                cmd[0].CommandText = stringA[0].ToString();
                SqReader[0] = cmd[0].ExecuteReader();

                object[] Value1 = new object[500000];
                int count = 0;

                while (SqReader[0].Read())
                {
                    object[] values = new object[SqReader[0].FieldCount];
                    SqReader[0].GetValues(values);
                    string[] stringD = Array.ConvertAll<object, string>(values, Convert.ToString);
                    Value1[count] = stringD[0];

                    count++;

                }

                Array.Resize(ref Value1, count);

                cmd[0].Dispose();
                SqReader[0].Close();

                string[] _string = Array.ConvertAll<object, string>(Value1, Convert.ToString);


                stringA[0].Clear();
                return null;
            }
            public string[] Get_Data_By_Query(string Query, int DB)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                    count++;
                }

                string[] _string = Array.ConvertAll<object, string>(Value, Convert.ToString);

                stringA[0].Clear();
                return _string;
            }

            public void Get_Defined_Para(object[,] DummyData, string key, Data_Class.Data_Editing.INT Data_InterFace)
            {


            }

            public void Get_Gross_Check_Para(Data_Class.Data_Editing.INT Data_Edit, string Select_Para, double Persent, string Selector, int SelectedBin)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                Get_Gross_Para = Select_Para;
                Get_Gross_Persent = Persent;
                //   Gross = ForGross_Fail_Unit;
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = false;
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_Gross_Check_Para_Thread), i);
                }

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }
                double test = TestTime1.Elapsed.TotalMilliseconds;


                //  stringA[0].Append("Select id from data where id not like '%F%'");
                stringA[0].Append("Select id from data");

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                ID = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    ID[count] = dr.ItemArray[0];
                    count++;
                }

                stringA[0].Clear();
                // List_Gross_Values.Add(Gross_Values1);
            }

            public void Get_Gross_Check_Para_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                int k = 0;
                for (k = 0; k < Data.Per_DB_Column_Count[i] - 1; k++)
                {
                    string[] Split_Dummy = Data.Reference_Header[Data.DB_Column_Limit * i + k].Split('_');
                    if (Split_Dummy.Length != 1)
                    {
                        if (Split_Dummy[1].ToUpper() == Get_Gross_Para.ToUpper())
                        {
                            ds[i] = new DataSet();

                            //   stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data where id not like '%F%'");
                            stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data");
                            cmd[i].CommandText = stringA[i].ToString();

                            sqlAdapter[i].SelectCommand = cmd[i];
                            sqlAdapter[i].Fill(ds[i]);

                            object[] DataValue = new object[ds[i].Tables[0].Rows.Count];

                            int count = 0;
                            foreach (DataRow dr in ds[i].Tables[0].Rows)
                            {
                                DataValue[count] = dr.ItemArray[0];
                                count++;
                            }


                            double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

                            double DataMin = doubles.Min();
                            double DataMax = doubles.Max();
                            double DataAve = doubles.Average();

                            double DataMinindex = doubles.ToList().IndexOf(DataMin);
                            double DataMaxindex = doubles.ToList().IndexOf(DataMax);

                            double Divide = DataMax / DataMin;

                            string[] test;
                            string _Substring = Get_Gross_Para.Substring(0, 1);

                            double MinSpec = 0f;
                            bool Define_Flag = false;

                            if (Get_Gross_Para.ToUpper().Contains("IBATT") || Get_Gross_Para.ToUpper().Contains("ICC") || Get_Gross_Para.ToUpper().Contains("IDD"))
                            {
                                Define_Flag = true;
                                test = Convert.ToString(Get_Gross_Persent).Split('.');
                                MinSpec = 1 - (Convert.ToDouble(test[1]) / 10);
                            }
                            else
                            {
                                Define_Flag = false;
                                MinSpec = Convert.ToDouble(Get_Gross_Persent) * -1;
                            }

                            if (Define_Flag)
                            {
                                for (int j = 0; j < doubles.Length; j++)
                                {
                                    if (DataAve / doubles[j] > Get_Gross_Persent || DataAve / doubles[j] < MinSpec)
                                    {
                                        if (!Gross.Contains(Convert.ToString(j + 1)))
                                        {
                                            //       Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                for (int j = 0; j < doubles.Length; j++)
                                {
                                    if (DataAve - doubles[j] > Get_Gross_Persent || doubles[j] - DataAve < MinSpec)
                                    {
                                        if (!Gross.Contains(Convert.ToString(j + 1)))
                                        {
                                            //         Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                                        }
                                    }
                                }
                            }

                            stringA[i].Clear();
                            cmd[i].CommandText = "";
                        }
                        //if (Get_Gross_Para == "POUT" && Split_Dummy.Length > 7 && Split_Dummy[6].ToUpper() == "FIXEDPOUT" && Split_Dummy[1].ToUpper() == "POUT")
                        //{
                        //    ds[i] = new DataSet();


                        //    //    stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data where id not like '%F%'");
                        //    stringA[i].Append("Select " + Data.New_Header[Data.DB_Column_Limit * i + k] + " from data");
                        //    cmd[i].CommandText = stringA[i].ToString();

                        //    sqlAdapter[i].SelectCommand = cmd[i];
                        //    sqlAdapter[i].Fill(ds[i]);

                        //    object[] DataValue = new object[ds[i].Tables[0].Rows.Count];

                        //    int count = 0;

                        //    foreach (DataRow dr in ds[i].Tables[0].Rows)
                        //    {
                        //        DataValue[count] = dr.ItemArray[0];
                        //        count++;
                        //    }

                        //    string remove = Split_Dummy[7].Replace("dBm", "");

                        //    double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

                        //    double DataMin = Convert.ToDouble(remove) - Get_Gross_Persent;
                        //    double DataMax = Convert.ToDouble(remove) + Get_Gross_Persent;

                        //    for (int j = 0; j < doubles.Length; j++)
                        //    {
                        //        if (doubles[j] < DataMin)
                        //        {
                        //            if (!Gross.Contains(Convert.ToString(j + 1)))
                        //            {
                        //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                        //            }

                        //        }
                        //        else if (doubles[j] > DataMax)
                        //        {
                        //            if (!Gross.Contains(Convert.ToString(j + 1)))
                        //            {
                        //                Gross_Values1[i].Add(Data.Reference_Header[Data.DB_Column_Limit * i + k], doubles); break;
                        //            }
                        //        }
                        //    }

                        //    stringA[i].Clear();
                        //    cmd[i].CommandText = "";
                        //}

                    }
                }
                ThreadFlags[i].Set();
            }
            public void Get_From_Db_Data_for_Anly(Data_Class.Data_Editing.INT Data_Edit)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;


                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    stringA[i].Clear();
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data), i);


                //}
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Get_From_Db_Data(i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }

            public void Get_Current_Setting(Data_Class.Data_Editing.INT Data_Edit, int NB)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;


                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Get_From_Db_Data(i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


            }
            public void Get_From_Db_Data(Object threadContext)
            {
                int i = (int)threadContext;

                CSV_Class.CSV.MERGE csv = new CSV_Class.CSV.MERGE();

                int Count = Data.Per_DB_Column_Count[i];
                this.Filename = "";


                if (_Flag)
                {

                    string[] name = conn[i].DataSource.ToString().Split('_');

                    for (int namei = 0; namei < 3; namei++)
                    {
                        if (namei == 2)
                        {
                            this.Filename += Lot_ID;
                        }
                        else
                        {
                            this.Filename += name[namei] + "_";
                        }

                    }
                    csv.Write_Open("C:\\Automation\\DB\\YIELD\\" + conn[i].DataSource.ToString().Substring(0, conn[i].DataSource.ToString().Length - 2) + ".csv\\" + this.Filename + "_" + i + ".csv");
                    string a = "C:\\Automation\\DB\\YIELD\\" + conn[i].DataSource.ToString().Substring(0, conn[i].DataSource.ToString().Length - 2) + ".csv\\" + this.Filename + ".csv";
                }
                else
                {

                    csv.Write_Open("C:\\Automation\\DB\\YIELD\\" + conn[i].DataSource.ToString().Substring(0, conn[i].DataSource.ToString().Length - 2) + ".csv\\" + conn[i].DataSource.ToString() + ".csv");
                }




                StringBuilder Apped = new StringBuilder();

                for (int Row = 0; Row < 3; Row++)
                {
                    if (Row == 0)
                    {

                        #region header
                        for (int j = 0; j < Count; j++)
                        {
                            if (j == 0)
                            {
                                if (i == 0)
                                {
                                    Apped.Append(Data.Reference_Header[0] + ",");
                                }
                                else
                                {
                                    Apped.Append(Data.Reference_Header[Data.DB_Column_Limit * i] + ",");
                                }

                            }
                            else
                            {
                                Apped.Append(Data.Reference_Header[Data.DB_Column_Limit * i + j] + ",");
                            }

                            if (j == Count - 1)
                            {
                                if (i == Data.Per_DB_Column_Count.Length - 1)
                                {
                                    Apped.Append("SubLot");
                                    csv.Write(Apped.ToString());
                                    Apped.Clear();
                                }
                                else
                                {
                                    Apped.Append("");
                                    csv.Write(Apped.ToString());
                                    Apped.Clear();
                                }

                            }

                        }
                        #endregion
                    }
                    else if (Row == 1)
                    {
                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count - 6;
                        }

                        #region Spec high
                        for (int j = 0; j < Count; j++)
                        {
                            if (i == 0)
                            {
                                if (j == 0)
                                {
                                    Apped.Append("HighL,");
                                }
                                else if (j < 10)
                                {
                                    Apped.Append(",");
                                }
                                else
                                {
                                    Apped.Append(DataSet_Value[i][1][j] + ",");
                                }
                            }
                            else if (i == Data.Per_DB_Column_Count.Length - 1)
                            {
                                if (j < Count)
                                {
                                    Apped.Append(DataSet_Value[i][1][j] + ",");
                                }
                                else
                                {
                                    int dummy_row = 0;
                                    for (dummy_row = 0; dummy_row < 6; j++)
                                    {
                                        Apped.Append(",");
                                    }

                                }
                            }
                            else
                            {
                                Apped.Append(DataSet_Value[i][1][j] + ",");
                            }

                            if (j == Count - 1)
                            {
                                Apped.Append("");
                                csv.Write(Apped.ToString());
                                Apped.Clear();
                            }
                        }

                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count + 6;
                        }
                        #endregion
                    }
                    else if (Row == 2)
                    {

                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count - 6;
                        }
                        #region Low
                        for (int j = 0; j < Count; j++)
                        {
                            if (i == 0)
                            {
                                if (j == 0)
                                {
                                    Apped.Append("LowL,");
                                }
                                else if (j < 10)
                                {
                                    Apped.Append(",");
                                }
                                else
                                {
                                    Apped.Append(DataSet_Value[i][0][j] + ",");
                                }
                            }
                            else if (i == Data.Per_DB_Column_Count.Length - 1)
                            {
                                if (j < Count)
                                {
                                    Apped.Append(DataSet_Value[i][0][j] + ",");
                                }
                                else
                                {
                                    int dummy_row = 0;
                                    for (dummy_row = 0; dummy_row < 6; j++)
                                    {
                                        Apped.Append(",");
                                    }

                                }
                            }
                            else
                            {
                                Apped.Append(DataSet_Value[i][0][j] + ",");
                            }

                            if (j == Count - 1)
                            {
                                Apped.Append("");
                                csv.Write(Apped.ToString());
                                Apped.Clear();
                            }
                        }

                        if (i == Data.Per_DB_Column_Count.Length - 1)
                        {
                            Count = Count + 6;
                        }
                        #endregion
                    }

                }

                if (i == Data.Per_DB_Column_Count.Length - 1)
                {
                    Count = Count + 1;
                }




                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {

                    Dictionary<string, List<string>> tests = key.Value;


                    if (!_Flag)
                    {
                        foreach (KeyValuePair<string, List<string>> ts in tests)
                        {

                            stringA[i].Clear();

                            if (_Flag == true)
                            {
                                stringA[i].Append("Select * from " + key.Key + " where Fail = '0'");
                            }
                            else
                            {
                                stringA[i].Append("Select * from  " + key.Key + " where Fail = '0'");
                            }

                            cmd[i] = new SQLiteCommand(conn[i]);
                            cmd[i].CommandText = stringA[i].ToString();
                            SqReader[i] = cmd[i].ExecuteReader();


                            while (SqReader[i].Read())
                            {

                                Stopwatch TestTime1 = new Stopwatch();
                                TestTime1.Restart();
                                TestTime1.Start();


                                object[] values = new object[SqReader[i].FieldCount];
                                SqReader[i].GetValues(values);

                                for (int j = 0; j < Count; j++)
                                {
                                    Apped.Append(values[j] + ",");
                                }

                                csv.Write(Apped.ToString());
                                Apped.Clear();

                            }
                            SqReader[i].Close();

                            stringA[i].Clear();
                            cmd[i].Dispose();

                            break;
                        }

                    }
                    else
                    {
                        foreach (KeyValuePair<string, List<string>> ts in tests)
                        {
                            if (ts.Key == Lot_ID)
                            {
                                stringA[i].Clear();

                                if (_Flag == true)
                                {
                                    stringA[i].Append("Select * from " + key.Key + " where Fail = '0'");
                                }
                                else
                                {
                                    stringA[i].Append("Select * from  " + key.Key + " where Fail = '0'");
                                }

                                cmd[i] = new SQLiteCommand(conn[i]);
                                cmd[i].CommandText = stringA[i].ToString();
                                SqReader[i] = cmd[i].ExecuteReader();


                                while (SqReader[i].Read())
                                {

                                    Stopwatch TestTime1 = new Stopwatch();
                                    TestTime1.Restart();
                                    TestTime1.Start();


                                    object[] values = new object[SqReader[i].FieldCount];
                                    SqReader[i].GetValues(values);

                                    for (int j = 0; j < Count; j++)
                                    {
                                        Apped.Append(values[j] + ",");
                                    }

                                    csv.Write(Apped.ToString());
                                    Apped.Clear();

                                }
                                SqReader[i].Close();

                                stringA[i].Clear();
                                cmd[i].Dispose();

                                break;
                            }
                        }
                    }




                }





                #region
                //for (int loop = 0; loop < Table_Count; loop++)
                //{
                //    stringA[i].Clear();

                //    if (_Flag == true)
                //    {
                //        stringA[i].Append("Select * from data" + loop + " where Fail = '0' and LOTID =" + "'" + Lot_ID + "'");
                //    }
                //    else
                //    {
                //        stringA[i].Append("Select * from data" + loop + " where Fail = '0'");
                //    }


                //    cmd[i].CommandText = stringA[i].ToString();
                //    SqReader[i] = cmd[i].ExecuteReader();


                //    while (SqReader[i].Read())
                //    {

                //        Stopwatch TestTime1 = new Stopwatch();
                //        TestTime1.Restart();
                //        TestTime1.Start();


                //        object[] values = new object[SqReader[i].FieldCount];
                //        SqReader[i].GetValues(values);

                //        for (int j = 0; j < Count; j++)
                //        {
                //            Apped.Append(values[j] + ",");
                //        }

                //        csv.Write(Apped.ToString());
                //        Apped.Clear();

                //    }
                //    SqReader[i].Close();

                //    stringA[i].Clear();
                //    cmd[i].CommandText = "";



                //}
                #endregion

                csv.Write_Close();

                ThreadFlags[i].Set();


            }
            public void Get_From_Db_Data_for_Anly_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit)
            {

            }
            public int Get_Sample_Count(int DB, string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);



                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                //int count = 0;
                //foreach (DataRow dr in ds[0].Tables[0].Rows)
                //{
                //    Value[count] = dr.ItemArray[0];
                //    count++;
                //}

                //   sqlAdapter[0].Dispose();
                //   cmd[0].Dispose();

                //   conn[0].Dispose();

                //sqlAdapter[0].Dispose();
                //stringA[0].Clear();

                //   cmd[0].Dispose();
                // conn[0].Close();


                //   int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Value.Length;
            }

            public void Get_From_Db_Ref_Header(Data_Class.Data_Editing.INT Data_Edit)
            {

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                this.Data = Data_Edit;

                int count = 0;

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    count += Data.Per_DB_Column_Count[i];
                }


                this.Data.Reference_Header = new string[count];

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    stringA[i].Clear();
                    ThreadFlags[i] = new ManualResetEvent(false);
                    Get_From_Db_Ref_Header_Thread(i);
                    //  ThreadPool.QueueUserWorkItem(new WaitCallback(Get_From_Db_Data_for_Anly_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }



            }
            public void Get_From_Db_Ref_Header_Thread(Object threadContext)
            {
                int i = (int)threadContext;

                //  int Count_Data = 0;
                int count = 0;

                stringA[i].Clear();

                stringA[i].Append("Select * from REFHEADER");


                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[i].CommandText = stringA[i].ToString();
                SqReader[i] = cmd[i].ExecuteReader();

                count = 0;

                while (SqReader[i].Read())
                {

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    object[] values = new object[SqReader[i].FieldCount];
                    SqReader[i].GetValues(values);
                    int ForCount = 0;

                    ForCount = values.Length - 5;

                    for (int j = 0; j < ForCount; j++)
                    {
                        this.Data.Reference_Header[this.Data.DB_Column_Limit * i + j] = Convert.ToString(values[j]);

                    }



                    double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                    count++;
                }
                SqReader[i].Close();
                cmd[i].Dispose();
                stringA[i].Clear();



                ThreadFlags[i].Set();


            }

            public int Get_Column_Count(Data_Class.Data_Editing.INT Data_Edit, string Query)
            {
                stringA[0].Clear();
                stringA[0].Append(Query);

                cmd[0] = new SQLiteCommand(conn[0]);
                sqlAdapter[0] = new SQLiteDataAdapter();

                cmd[0].CommandText = stringA[0].ToString();
                ds[0] = new DataSet();

                sqlAdapter[0].SelectCommand = cmd[0];
                sqlAdapter[0].Fill(ds[0]);

                Value = new object[ds[0].Tables[0].Rows.Count];

                int count = 0;
                foreach (DataRow dr in ds[0].Tables[0].Rows)
                {
                    Value[count] = dr.ItemArray[0];
                }

                sqlAdapter[0].Dispose();
                cmd[0].Dispose();

                int[] Data_Count = Array.ConvertAll<object, int>(Value, Convert.ToInt32);

                return Data_Count[0];
            }

            public void Close(Data_Class.Data_Editing.INT Data_Edit)
            {

                for (int i = 0; i < Data.DB_Count; i++)
                {
                    ThreadFlags[i] = new ManualResetEvent(false);
                    ThreadPool.QueueUserWorkItem(new WaitCallback(close_Thread), i);
                }
                for (int i = 0; i < Data.DB_Count; i++)
                {
                    Wait[i] = ThreadFlags[i].WaitOne();
                }


                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    tran[i].Dispose();
                    cmd[i].Dispose();
                    conn[i].Dispose();
                    // conn[i].Close();
                    sqlAdapter[i].Dispose();

                }
            }
            public void close_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                cmd[i].CommandText = "vacuum";
                cmd[i].ExecuteNonQuery();

                ThreadFlags[i].Set();
            }

            public void Read_Dispose(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }

            public void Set_Conn(Data_Class.Data_Editing.INT Data_Edit)
            {
                for (int i = 0; i < Data_Edit.DB_Count; i++)
                {
                    cmd[i].Dispose();


                }
            }

            public void trans(Data_Class.Data_Editing.INT Data_Edit)
            {
                Data = Data_Edit;


              //  tran[0] = conn[0].BeginTransaction();
               // cmd[0].Transaction = tran[0];

                Tran_Thread(0);
                //tran = new SQLiteTransaction[Data_Edit.DB_Count];

                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    ThreadFlags[i] = new ManualResetEvent(false);
                //    ThreadPool.QueueUserWorkItem(new WaitCallback(Tran_Thread), i);
                //}
                //for (int i = 0; i < Data.DB_Count; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}


            }
            public void Tran_Thread(Object threadContext)
            {
                int i = (int)threadContext;
                cmd[0].Dispose();
                conn[0].Dispose();

                conn[0] = new SQLiteConnection(strConn[i]);
                cmd[0] = new SQLiteCommand(conn[i]);
                conn[0].Open();


                tran[0] = conn[i].BeginTransaction();
                cmd[0].Transaction = tran[i];

               // ThreadFlags[i].Set();
            }

            public void Commit(Data_Class.Data_Editing.INT Data_Edit)
            {

                //for (int i = 0; i < 1; i++)
                //{
                    Commit_thread(0);
            //           ThreadFlags[i] = new ManualResetEvent(false);
           //         ThreadPool.QueueUserWorkItem(new WaitCallback(Commit_thread), i);
                //}
                //for (int i = 0; i < 1; i++)
                //{
                //    Wait[i] = ThreadFlags[i].WaitOne();
                //}
            }
            public void Commit_thread(Object threadContext)
            {
                int i = (int)threadContext;
                tran[i].Commit();
                ThreadFlags[i].Set();
            }
            public void STDEVandMedian(List<double[]> Ds, int DB, int RowCount)
            {
                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                double[][] ReturnValue = new double[Data.Per_DB_Column_Count[DB]][];

                for (int i = 0; i < Data.Per_DB_Column_Count[DB]; i++)
                {
                    ReturnValue[i] = new double[RowCount];
                }
                double dummytesttime1 = TestTime1.Elapsed.TotalMilliseconds;
                int j = 0;
                int k = 0;


                foreach (double[] o in Ds)
                {
                    var t = o;
                    for (int q = 0; q < t.Length - 2; q++)
                    {
                        ReturnValue[j][k] = t[q];
                        j++;
                    }

                    j = 0;
                    k++;
                }

                int Para_Count = 0;

                for (int i = 0; i < ReturnValue.Length; i++)
                {
                    double average = ReturnValue[i].Average();
                    double Median = 0f;

                    if (ReturnValue[i].Length % 2 == 0)
                    {
                        Array.Sort(ReturnValue[i]);

                        double dummyi = ReturnValue[i][(ReturnValue[i].Length / 2) - 1];
                        double dummyj = ReturnValue[i][ReturnValue[i].Length / 2];
                        Median = (dummyi + dummyj) / 2;
                    }
                    else
                    {
                        Array.Sort(ReturnValue[i]);
                        int GetMedian_i = (ReturnValue[i].Length) / 2;
                        Median = ReturnValue[i][GetMedian_i];
                    }

                    double minusSquareSummary = 0.0;

                    foreach (double source in ReturnValue[i])
                    {
                        minusSquareSummary += (source - average) * (source - average);
                    }

                    double stdev = Math.Sqrt(minusSquareSummary / (ReturnValue[i].Length - 1));

                    for (int q = 0; i < For_New_Spec_Cal_Value_by_rowsdata[Data.Reference_Header[0]].CPK.Length; i++)
                    {

                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Std[q] = stdev;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Median_Data[q] = Median;
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Min_Data[q] = ReturnValue[i].Min();
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Max_Data[q] = ReturnValue[i].Max();
                        Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].Avg[q] = ReturnValue[i].Average();

                    }
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].L_CPK = (average - ReturnValue[i].Min()) / (3 * stdev);
                    //Cal_Value_by_rowsdata[Data.Reference_Header[Data.DB_Column_Limit * DB + Para_Count]].H_CPK = (average - ReturnValue[i].Max()) / (3 * stdev);

                    Para_Count++;

                }
                double dummytesttime2 = TestTime1.Elapsed.TotalMilliseconds;

            }
            static double[] STDEVandMedian(DataSet ds)
            {
                List<double> DataSet_Values = new List<double>();
                double[] ReturnValue = new double[2];

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    DataSet_Values.Add(Convert.ToDouble(dr.ItemArray[0]));
                }

                double average = DataSet_Values.Average();
                double Median = 0f;

                if (DataSet_Values.Count % 2 == 0)
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = DataSet_Values.Count / 2;
                    Median = DataSet_Values[GetMedian_i];
                }
                else
                {
                    DataSet_Values.Sort();
                    int GetMedian_i = (DataSet_Values.Count + 1) / 2;
                    Median = DataSet_Values[GetMedian_i];
                }

                double minusSquareSummary = 0.0;

                foreach (double source in DataSet_Values)
                {
                    minusSquareSummary += (source - average) * (source - average);
                }

                double stdev = Math.Sqrt(minusSquareSummary / (DataSet_Values.Count - 1));

                ReturnValue[0] = stdev; ReturnValue[1] = Median;

                return ReturnValue;
            }

            public string Get_Data_From_Table(string Table, string header)
            {

                return "";
            }



        }

        public INT Open(string Key)
        {
            INT Int = null;
            switch (Key)
            {
                case "YIELD":
                    Int = new Yield_DB();
                    Int.Limit = 10000;

                    break;
                case "FCM":
                    Int = new FCM_Automation_EXCEL();

                    break;

                case "MERGE":
                    Int = new MERGE();
                    Int.Limit = 10000;
                    break;
                case "MERGE_S4PD":
                    Int = new MERGE_S4PD();
                    Int.Limit = 10000;
                    break;

                case "BOXPLOT":
                    Int = new BOXPLOT();
                    break;
            }
            return Int;
        }

        public interface INT
        {
            Data_Class.Data_Editing.INT Data { get; set; }
            ReaderWriterLockSlim[] sqlitelock { get; set; }
            string[] strConn { get; set; }
            SQLiteConnection[] conn { get; set; }
            SQLiteCommand[] cmd { get; set; }

            SQLiteDataAdapter[] sqlAdapter { get; set; }
            SQLiteCommandBuilder[] sqlcmdbuilder { get; set; }
            SQLiteDataReader[] SqReader { get; set; }

            DbDataReader[] DbReader { get; set; }
            DataSet[] ds { get; set; }

            DataTable dt_test { get; set; }
            DataTable[] dt { get; set; }
            SQLiteTransaction[] tran { get; set; }

            ManualResetEvent[] ThreadFlags { get; set; }

            ManualResetEvent[] Insert_ThreadFlags { get; set; }
            StringBuilder[] stringA { get; set; }
            bool[] Wait { get; set; }

            bool[] Insert_Thread_Wait { get; set; }
            double[] Testtime { get; set; }

            double[][] test { get; set; }
            string[][] Teststring { get; set; }
            double[][] Testdouble { get; set; }

            object[] ID { get; set; }
            object[] Value { get; set; }
            object[] WAFER_ID { get; set; }
            object[] LOT_ID { get; set; }
            object[] SITE_ID { get; set; }

            Dictionary<string, double[]> Selected_Parameter_Distribution { get; set; }

            object[] Variation { get; set; }

            int Limit { get; set; }
            int Limit_Count { get; set; }

            int Table_Count { get; set; }
            int Spec_Table_Count { get; set; }

            List<List<RowAndPass>[]>[] Yield_Test { get; set; }
            List<List<RowAndPass>[]>[] Yield_Test_New_Spec { get; set; }

            List<List<int>>[] For_Any_Yield { get; set; }
            List<List<List<int>>>[] For_Any_Yield_For_Lot { get; set; }
            List<List<List<int>>>[] For_Any_Yield_For_SITE { get; set; }
            List<List<int>[]>[] For_Any_Yield_Percent { get; set; }
            List<List<int>[]>[] ForCampare_Yield { get; set; }
            List<List<int>[]>[] For_Any_Yield_Percent_For_New_Spec { get; set; }
            List<List<int>>[] For_Any_Yield_For_New_Spec { get; set; }
            List<List<int>[]>[] For_New_Spec_ForCampare_Yield { get; set; }
            List<int[]>[] ForCampare_Yield_Fro_DB { get; set; }
            List<List<int[]>>[] ForCampare_Yield_Fro_DB_List { get; set; }
            List<List<int>>[] For_New_Spec_ForCampare_Yield2 { get; set; }
            List<List<List<List<int>[]>>>[] ForCampare_Yield_DB_LotVariation { get; set; }
            List<List<List<int[]>>>[] ForCampare_Yield_Fro_DB_List_LotVariation { get; set; }

            Dictionary<string, int> Refer_Site_And_Num { get; set; }
            Dictionary<string, int> Refer_Lot_And_Num { get; set; }
            List<int>[] ForCampare_Yield_List { get; set; }
            List<List<int>[]> ForCampare_Yield_List1 { get; set; }
            List<List<int>[]>[] ForCampare_Yield_List2 { get; set; }
            Dictionary<string, Values> Values { get; set; }
            Dictionary<string, Data_Calculation> Cal_Value_by_rowsdata { get; set; }
            Dictionary<string, Data_Calculation> For_New_Spec_Cal_Value_by_rowsdata { get; set; }

            List<double[]>[] DB_DataSet_Values { get; set; }

            Dictionary<string, int> Lot_Dic { get; set; }
            Dictionary<string, int> Site_Dic { get; set; }
            Dictionary<string, int> Bin_Dic { get; set; }

            Dictionary<string, Dictionary<string, List<string>>> Matching_Lots { get; set; }
            Dictionary<string, List<string>> Matching_Lot { get; set; }

            int TheFirst_Trashes_Header_Count { get; set; }
            int TheEnd_Trashes_Header_Count { get; set; }

            Dictionary<string, CSV_Class.For_Box>[] Dic_Test { get; set; }
            Dictionary<string, CSV_Class.For_Box> Dic_Test_For_Spec_Gen { get; set; }

            Stopwatch[] TestTime1 { get; set; }
            Stopwatch[] TestTime2 { get; set; }
            Stopwatch[] TestTime3 { get; set; }
            Stopwatch[] TestTime4 { get; set; }
            Stopwatch[] TestTime5 { get; set; }
            long SampleCount { get; set; }
            object Update_Data_ID { get; set; }
            string[] Update_Datas_ID { get; set; }
            string Get_Gross_Para { get; set; }
            string Get_Gross_Selector { get; set; }
            List<Dictionary<string, Gross>[]> List_Gross_Values { get; set; }
            Dictionary<string, Gross>[] Gross_Values1 { get; set; }
            string Table { get; set; }

            double[] Make_New_Spec_For_Yield_Min { get; set; }
            double[] Make_New_Spec_For_Yield_Max { get; set; }
            List<string> Gross { get; set; }

            List<string[]>[] DataSet_Value { get; set; }
            List<double[]>[] DataSet_Double_Value { get; set; }

            List<int>[] Check { get; set; }
            List<List<int>[]> Test { get; set; }
            string Lot_ID { get; set; }
            string SubLot_ID { get; set; }
            string Tester_ID { get; set; }
            string Site { get; set; }
            string Bin { get; set; }
            string ID_Unit { get; set; }
            int Bin_place { get; set; }
            string Filename { get; set; }
            object[] Std_Value { get; set; }
            double[] Std_Value_Convert { get; set; }

            Dictionary<string, IQR> DIC_IQR { get; set; }

            int Count_Current_Setting { get; set; }
            string Query { get; set; }
            long NB { get; set; }
            bool _From_Db { get; set; }

            bool _Flag { get; set; }
            bool Clotho_Spec_Flag { get; set; }

            bool _SUBLOT_Flag { get; set; }

            string Before_Lot_ID { get; set; }
            string Changed_Lot_ID { get; set; }

            int[] Each_Thread_Count { get; set; }

            string[] No_Index { get; set; }
            string[] Paraname { get; set; }
            string[] SpecMin { get; set; }
            string[] SpecMax { get; set; }
            string[] DataMin { get; set; }
            string[] DataMedian { get; set; }
            string[] DataMax { get; set; }
            string[] CPK { get; set; }
            string[] STD { get; set; }
            string[] Percent { get; set; }
            string[] Fail{ get; set; }

            string[] Line { get; set; }

            void Open_DB(string FileName, Data_Class.Data_Editing.INT Data_Edit);
            void Open_DB(string[] FileName, Data_Class.Data_Editing.INT Data_Edit);
            void DropTable(Data_Class.Data_Editing.INT Data_Edit, string Query);
            void Insert_Header(Data_Class.Data_Editing.INT Data_Edit);
            void Insert_Spec_Header(Data_Class.Data_Editing.INT Data_Edit);
            void Insert_Current_Setting(Data_Class.Data_Editing.INT Data_Edit);
            void Insert_New_Spec_Header(Data_Class.Data_Editing.INT Data_Edit);
            void Insert_Data(Data_Class.Data_Editing.INT Data_Edit);

            void Insert_Ref_Header_Data(Data_Class.Data_Editing.INT Data_Edit);
            void Insert_Data(long Sample);

            void Insert_Data_Get_From_DB(int Sample);
            void Insert_Spec_Get_From_DB(Data_Class.Data_Editing.INT Data_Edit);
            void Insert_Spec_Data(string Tablename);

            void Insert_Spec_Data(Data_Class.Data_Editing.INT Data_Edit, string Table);
            void Insert_Current_Setting_Data(Data_Class.Data_Editing.INT Data_Edit, string Table);

            void Insert_Files_Name(string FileName);

    
            void Make_table(string Tablename);
            void Make_table2(Data_Class.Data_Editing.INT Data_Edit, string Tablename);
            void Make_table_For_Filename(Data_Class.Data_Editing.INT Data_Edit, string Tablename);
            void Make_table_For_Trace(string Tablename, string Chan, bool Flag);

            void Delete_Spec_Data(string Tablename);
            void Delete_Lot_Data(string Query);
            void Gross_Update_Data(object data);

            void Save_table(Data_Class.Data_Editing.INT Data_Edit);
            void Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit);


            void Road_Save_Customer_Spec_table(Data_Class.Data_Editing.INT Data_Edit);

            void LOTID_Update(string Query, string Query2, string CellID);

            void Gross_Update_Datas(List<string> data);

            void Chnaged_Spec_Update_Data(int DB, int Index, string Parameter, double Spec, int GetId);
            Dictionary<string, double[]> Chnaged_Spec_Anl_Yield(int DB, int Index, string Parameter);
            void Get_Rows_Data(Data_Class.Data_Editing.INT Data_Edit);
            void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface);
            void Get_Selected_Para(Data_Class.Data_Editing.INT Data_Interface, DataTable dt);
            void Get_Selected_Para(int i, string Select_Para, bool Lot_Flag, string Selector);
            void Get_Defined_Para(object[,] DummyData, string Key, Data_Class.Data_Editing.INT Data);
            double[] Get_Find_Bin(string Query);

            List<object[]> Get_Data_By_Querys(string Query);

            string[] Get_Data_By_Query(string Query);

            Dictionary<string, object[]> Get_Data_By_Query_S4PD(string Query, string Chan);
            string[] Get_Data_By_Query(string Query, int DB);
            void Get_Ave_Data(Data_Class.Data_Editing.INT Data_Edit);
            void Get_Ave_Data_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit);
            void Get_Ave_Data2(Data_Class.Data_Editing.INT Data_Edit);

            void Set_Refer_for_Anlyzer(Data_Class.Data_Editing.INT Data_Edit);
            void Get_Saved_Spec(Data_Class.Data_Editing.INT Data_Edit);
            void Get_Gross_Check_Para(Data_Class.Data_Editing.INT Data_Edit, string Select_Para, double Persent, string Selector, int SelectedBin);

            void Get_From_Db_Data_for_Anly(Data_Class.Data_Editing.INT Data_Edit);
            void Get_From_Db_Data_for_Anly_For_New_Spec(Data_Class.Data_Editing.INT Data_Edit);

            void Get_Current_Setting(Data_Class.Data_Editing.INT Data_Edit, int NB);

            void Get_From_Db_Ref_Header(Data_Class.Data_Editing.INT Data_Edit);

            string Get_Data_From_Table(string Table, string header);

            int Get_Sample_Count(int DB, string Query);
            int Get_Column_Count(Data_Class.Data_Editing.INT Data_Edit, string Query);

            void Close(Data_Class.Data_Editing.INT Data_Edit);
            void Read_Dispose(Data_Class.Data_Editing.INT Data_Edit);
            void Set_Conn(Data_Class.Data_Editing.INT Data_Edit);
            void trans(Data_Class.Data_Editing.INT Data_Edit);
            void Commit(Data_Class.Data_Editing.INT Data_Edit);

        }

        public class Data_Information
        {
            public object[,] DummyData;
            public string Key;
            public int DB_NB;
            public string[] Ref_Header;
            public double[] High_Spec;
            public double[] Low_Spec;

            public Data_Information(object[,] DummyData, string Key, int DB_NB, string[] Ref_Header, double[] High_Spec, double[] Low_Spec)
            {
                this.DummyData = DummyData;
                this.Key = Key;
                this.DB_NB = DB_NB;
                this.Ref_Header = Ref_Header;
                this.High_Spec = High_Spec;
                this.Low_Spec = Low_Spec;
            }
        }
        public class Data_Calculation
        {
            public int[] No;
            public string[] Parameter;

            public string[] Min_Selector;
            public string[] Max_Selector;
            public double[] Min_Spec_Control;
            public double[] Max_Spec_Control;
            public double[] Min_Spec;
            public double[] Max_Spec;
            public double[] Min_Data;
            public double[] Median_Data;
            public double[] Max_Data;
            public double[] CPK;
            public double[] Std;
            public double[] Persent;
            public long[] Fail_Count;
            public long[] Outlier;

            public double[] Avg;
            public double[] L_CPK;
            public double[] H_CPK;
            public double[] L_Avg;
            public double[] H_Avg;
            public double[] L_Std;
            public double[] H_Std;

            public double[] L_IQR_Value;
            public double[] H_IQR_Value;
            public long[] Count;

            public Dictionary<double, double> L_IQR_Value_Array;
            public Dictionary<double, double> H_IQR_Value_Array;

            public Data_Calculation(int index)
            {

                No = new int[index];
                Parameter = new string[index];

                Min_Selector = new string[index];
                Max_Selector = new string[index];
                Min_Spec_Control = new double[index];
                Max_Spec_Control = new double[index];
                Min_Spec = new double[index];
                Max_Spec = new double[index];
                Min_Data = new double[index];
                Median_Data = new double[index];
                Max_Data = new double[index];
                CPK = new double[index];
                Std = new double[index];
                Persent = new double[index];
                Fail_Count = new long[index];
                Outlier = new long[index];

                Avg = new double[index];
                L_CPK = new double[index];
                H_CPK = new double[index];
                L_Avg = new double[index];
                H_Avg = new double[index];
                L_Std = new double[index];
                H_Std = new double[index];

                L_IQR_Value = new double[index];
                H_IQR_Value = new double[index];
                Count = new long[index];

                this.L_IQR_Value_Array = new Dictionary<double, double>();
                this.H_IQR_Value_Array = new Dictionary<double, double>();
            }
        }
        public class Values
        {
            public object[] Data;
            public double Low;
            public double High;
            public string Key;
            public Values(object[] Data, double Low, double High, string Key)
            {
                this.Data = Data;
                this.Low = Low;
                this.High = High;
                this.Key = Key;
            }
        }

        public class IQR
        {
            public double L_IQR;
            public double H_IQR;
            public string[] SN;

            public IQR(double L_IQR, double H_IQR, string[] SN)
            {
                this.L_IQR = L_IQR;
                this.H_IQR = H_IQR;
                this.SN = SN;

            }
        }

        public class RowAndPass
        {
            public long SN;
            public int Row;
            public int Pass;

            public RowAndPass(long SN, int Row, int Pass)
            {
                this.SN = SN;
                this.Row = Row;
                this.Pass = Pass;
            }

        }
        public class Gross
        {
            public double[] Data;
            public double STD;
            public double SpecL;
            public double SpecH;
            public Gross(double[] Data, double STD, double SpecL, double SpecH)
            {
                this.Data = Data;
                this.STD = STD;
                this.SpecL = SpecL;
                this.SpecH = SpecH;
            }

        }

        [SQLiteFunction(Arguments = 1, FuncType = FunctionType.Aggregate, Name = "AVERAGE")]
        public class AVERAGE : SQLiteFunction
        {
            //////////////////////////////////////////////////////////////////////////////////////////////////// Field
            ////////////////////////////////////////////////////////////////////////////////////////// Private

            #region Field

            /// <summary>
            /// 카운트
            /// </summary>
            private int count = 0;

            #endregion

            //////////////////////////////////////////////////////////////////////////////////////////////////// Method
            ////////////////////////////////////////////////////////////////////////////////////////// Public

            #region 단계별 처리하기 - Step(argumentArray, stepNumber, contextData)

            /// <summary>
            /// 단계별 처리하기
            /// </summary>
            /// <param name="argumentArray">인자 배열</param>
            /// <param name="stepNumber">단계 변호</param>
            /// <param name="contextData">컨텍스트 데이터</param>
            public override void Step(object[] argumentArray, int stepNumber, ref object contextData)
            {
                if (contextData == null)
                {
                    contextData = 0.0;

                    this.count = 0;
                }

                contextData = Convert.ToDouble(contextData) + Convert.ToDouble(argumentArray[0]);
                //contextData = Convert.ToDouble(contextData) + argumentArray[0];
                this.count++;
            }

            #endregion
            #region 최종 처리하기 - Final(contextData)

            /// <summary>
            /// 최종 처리하기
            /// </summary>
            /// <param name="contextData">컨텍스트 데이터</param>
            /// <returns>결과</returns>
            public override object Final(object contextData)
            {
                return (double)contextData / count;
            }

            #endregion
        }

        [SQLiteFunction(Arguments = 1, FuncType = FunctionType.Aggregate, Name = "STDEV")]
        public class STDEV : SQLiteFunction
        {
            //////////////////////////////////////////////////////////////////////////////////////////////////// Field
            ////////////////////////////////////////////////////////////////////////////////////////// Private

            #region Field

            /// <summary>
            /// 소스 리스트
            /// </summary>
            private List<double> sourceList = new List<double>();

            #endregion

            //////////////////////////////////////////////////////////////////////////////////////////////////// Method
            ////////////////////////////////////////////////////////////////////////////////////////// Public

            #region 단계별 처리하기 - Step(argumentArray, stepNumber, contextData)

            /// <summary>
            /// 단계별 처리하기
            /// </summary>
            /// <param name="argumentArray">인자 배열</param>
            /// <param name="stepNumber">단계 변호</param>
            /// <param name="contextData">컨텍스트 데이터</param>
            public override void Step(object[] argumentArray, int stepNumber, ref object contextData)
            {
                if (contextData == null)
                {
                    contextData = 0;

                    this.sourceList.Clear();
                }

                this.sourceList.Add(Convert.ToDouble(argumentArray[0]));
            }

            #endregion
            #region 최종 처리하기 - Final(contextData)

            /// <summary>
            /// 최종 처리하기
            /// </summary>
            /// <param name="contextData">컨텍스트 데이터</param>
            /// <returns>결과</returns>
            public override object Final(object contextData)
            {
                if (this.sourceList.Count == 1)
                {
                    return double.NaN;
                }

                double average = this.sourceList.Average();
                double Median = 0f;

                if (this.sourceList.Count % 2 == 0)
                {
                    this.sourceList.Sort();
                    int GetMedian_i = this.sourceList.Count / 2;
                    Median = this.sourceList[GetMedian_i];
                }
                else
                {
                    this.sourceList.Sort();
                    int GetMedian_i = (this.sourceList.Count + 1) / 2;
                    Median = this.sourceList[GetMedian_i];
                }

                double minusSquareSummary = 0.0;

                foreach (double source in sourceList)
                {
                    minusSquareSummary += (source - average) * (source - average);
                }


                double stdev = Math.Sqrt(minusSquareSummary / (sourceList.Count - 1));

                return stdev;
            }

            #endregion
        }


    }
    public static class LinqExtensions
    {
        public static void DisposeItems<T>(this IEnumerable<T> source) where T : IDisposable
        {
            foreach (var item in source)
            {
                item.Dispose();
            }
        }
    }

}


