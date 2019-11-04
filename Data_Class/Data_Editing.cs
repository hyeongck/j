using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Data_Class
{
    public class Data_Editing
    {
       // public static string[] New_Header;
       // public static double[] New_HighSpec;
       // public static double[] New_LowSpec;
       // public static string[] ForAnl_NewMinSpec { get; set; }
       // public static string[] ForAnl_NewMaxSpec { get; set; }

        public class FCM_Automation_EXCEL : INT
        {
            Data_Editing Edit = new Data_Editing();
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ManualResetEvent[] ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public string[] Getstring { get; set; }
            public List<string> Reference_Header_List { get; set; }
            public string[] Reference_Header { get; set; }

            public string Data_Table { get; set; }

            public double[] New_HighSpec { get; set; }
            public double[] New_LowSpec { get; set; }
            public string[] New_Header { get; set; }

            public string[] For_GetSpec_Header { get; set; }
            public string[] Customer_Clotho_Spec_Data { get; set; }
            public string[] Clotho_Spec_Data { get; set; }

            public double[] New_Data { get; set; }
            public double[] For_Thread_New_Data { get; set; }
            public string Defined_Spec_Min { get; set; }
            public string Defined_Spec_Max { get; set; }
            public string Defined_Spec_Typical { get; set; }
            public int Defined_Convert_Index { get; set; }
            public string Defined_Convert { get; set; }
            public string Defined_Complience { get; set; }
            public int Defined_Both { get; set; }
            public string Spec_Num { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public int DB_Count { get; set; }
            public int DB_Column_Limit { get; set; }

            public int[] Per_DB_Column_Count { get; set; }
            public int[] Per_DB_Column_Count_Start { get; set; }
            public int[] Per_DB_Column_Count_End { get; set; }

            public bool Flag { get; set; }
            public bool SUBLOT_Falg { get; set; }
            public Dictionary<string, string> Dummy_Spec_Band { get; set; }
            public Dictionary<string, Dictionary<string, string>> Spec_Band { get; set; }

            public Dictionary<string, Spec> Dic_Spec { get; set; }
            public Dictionary<string, SWBIN> SWBIN_Dic { get; set; }
            public List<Clotho_Spec> Clotho_List { get; set; }
            public List<Clotho_Spec> Customor_Clotho_List { get; set; }
            public List<Clotho_Spec> New_Clotho_List { get; set; }
            public List<Clotho_Spec> Clotho_Spcc_List { get; set; }
            public string[] Para { get; set; }
            public string[] Band { get; set; }
            public int ConditionCount { get; set; }
            public int Set_ID { get; set; }
            public int Set_FAIL { get; set; }

            public bool _From_DB { get; set; }

            public string[] Ref_New_Header { get; set; }
            public double[] Ref_New_HighSpec { get; set; }
            public double[] Ref_New_LowSpec { get; set; }
            public string[] Ref_ForAnl_NewMinSpec { get; set; }
            public string[] Ref_ForAnl_NewMaxSpec { get; set; }

            public bool Find_First_Row(string[] Getstring)
            {
                try
                {
                    if (Getstring[0].ToUpper().Contains("PARAMETER") || Getstring[0].ToUpper().Contains("PARAMETERNAME"))
                    {
                        for (int i = 1; i < 10; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        for (int i = Getstring.Length - 6; i < Getstring.Length; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        int j = 1;
                        Reference_Header = new string[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                        Reference_Header[0] = Getstring[0];
                        for (int i = TheFirst_Trashes_Header_Count + 1; i < Getstring.Length - TheEnd_Trashes_Header_Count; i++)
                        {
                            string Dummy = Getstring[i].Replace('.', '_');
                            Dummy = Dummy.Replace('-', '_');

                            Reference_Header[j] = Getstring[i];
                            j++;
                        }

                        Flag = true;
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            public bool Find_Spec_Row(string[] Getstring, bool Flag)
            {
                try
                {
                    Flag = false;
                    if (Getstring[0].ToUpper().Contains("HIGH") || Getstring[0].ToUpper().Contains("LOW"))
                    {
                        int j = 1;
                        if (Getstring[0].ToUpper().Contains("HIGH"))
                        {
                            New_HighSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_HighSpec[0] = Convert.ToDouble(0);
                            for (int i = TheFirst_Trashes_Header_Count + 1; i < Getstring.Length - TheEnd_Trashes_Header_Count; i++)
                            {
                                New_HighSpec[j] = Convert.ToDouble(Getstring[i]);
                                j++;
                            }

                            Ref_New_HighSpec = New_HighSpec;
                            Flag = true;
                        }
                        if (Getstring[0].ToUpper().Contains("LOW"))
                        {
                            New_LowSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_LowSpec[0] = Convert.ToDouble(0);
                            for (int i = TheFirst_Trashes_Header_Count + 1; i < Getstring.Length - TheEnd_Trashes_Header_Count; i++)
                            {
                                New_LowSpec[j] = Convert.ToDouble(Getstring[i]);
                                j++;
                            }
                            Ref_New_LowSpec = New_LowSpec;
                            Flag = true;
                        }
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }
            public void Find_Cloth_DataFile(string[] Getstring)
            {

                SWBIN_Dic = new Dictionary<string, SWBIN>();

                string[] Values = Getstring[0].Split(',');
                int Values_length = Values.Length;

                int Bin_Count = (Values_length - 5) / 2;

                Clotho_List = new List<Clotho_Spec>();
                Clotho_Spcc_List = new List<Clotho_Spec>();
                Customor_Clotho_List = new List<Clotho_Spec>();

                for (int i = 0; i < Getstring.Length; i++)
                {
                    Values = Getstring[i].Split(',');

                    if (Values[0].ToUpper().Contains("SWBIN"))
                    {

                        for (int k = 0; k < 10000; k++)
                        {
                            i++;
                            if (Values[0].Contains("#END"))
                            {
                                break;
                            }
                            else
                            {
                                Values = Getstring[i].Split(',');
                                SWBIN SW = new SWBIN(Values[1], Values[3], false);

                                if (Values[1].ToUpper().Contains("PASS"))
                                {
                                    SWBIN_Dic.Add(Values[0], SW);
                                }


                                string[] Key = SWBIN_Dic.Keys.ToArray();
                                int index = 0;
                                int Find_Index = 0;

                                foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                                {
                                    if (Item.Value.Name.ToUpper().ToString().Contains("PASS"))
                                    {
                                        Item.Value.Flag = true;
                                        Find_Index = index;
                                    }
                                    index++;
                                }
                            }
                        }
                    }
                    else if (Values[0].ToUpper().Contains("TESTNUM"))
                    {
                        int Start_Index = 5;
                        int l = 1;
                        int j = 0;
                        var Para = new string[40000];
                        Para[0] = "PARAMETER";


                        double[] Min = new double[Bin_Count];
                        double[] Max = new double[Bin_Count];

                        int n = 0;
                        Start_Index = 5;

                        for (j = Start_Index; j < Values.Length; j++)
                        {
                            Min[n] = -9999f;
                            Max[n] = 9999f;
                            j++;
                            n++;
                        }
                        Clotho_Spec Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Clotho_List.Add(Clotho_Spec_Data);



                        n = 0;
                        int Index = 0;
                        string[] Array = SWBIN_Dic.Keys.ToArray();
                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Index++;
                            }
                        }

                        Min = new double[Index];
                        Max = new double[Index];

                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Min[n] = -9999f;
                                Max[n] = 9999f;
                                Start_Index++;
                                n++;
                            }
                        }

                        Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Clotho_Spcc_List.Add(Clotho_Spec_Data);


                        for (int h = i; h < Getstring.Length - 1; h++)
                        {

                            i++;
                            Values = Getstring[i].Split(',');

                            Min = new double[Bin_Count];
                            Max = new double[Bin_Count];

                            n = 0;
                            Start_Index = 5;

                            for (j = Start_Index; j < Values.Length; j++)
                            {
                                Min[n] = Convert.ToDouble(Values[j]);
                                Max[n] = Convert.ToDouble(Values[j + 1]);
                                j++;
                                n++;
                            }
                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Clotho_List.Add(Clotho_Spec_Data);


                            n = 0;
                            Index = 0;
                            Array = SWBIN_Dic.Keys.ToArray();
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Index++;
                                }
                            }

                            Min = new double[Index];
                            Max = new double[Index];
                            Start_Index = 5;
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Min[n] = Convert.ToDouble(Values[Start_Index]);
                                    Max[n] = Convert.ToDouble(Values[Start_Index + 1]);
                                    Start_Index++;
                                    Start_Index++;
                                    n++;
                                }
                            }

                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Clotho_Spcc_List.Add(Clotho_Spec_Data);

                            Para[l] = Values[1];
                            l++;
                        }


                        System.Array.Resize(ref Para, l);
                        Ref_New_Header = Para;

                        Para = null;
                    }
                }



            }
            public void Find_Cloth_DataFile_For_New_Spec(string[] Customer)
            {

            }
            public void Define_DB_Count(string[] Getstring)
            {

                if (Reference_Header.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count < 2000)
                {
                    DB_Count = 1;
                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count[0] = DB_Column_Limit;
                    Per_DB_Column_Count_Start[0] = DB_Column_Limit;
                    Per_DB_Column_Count_End[0] = DB_Column_Limit;

                }
                else
                {
                    double length = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(DB_Column_Limit);
                    double Temp = Math.Truncate(length);

                    int Dummy_DB_Count = 0;

                    if (length > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
                    else DB_Count = Convert.ToInt16(Temp);

                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count_Start = new int[DB_Count];
                    Per_DB_Column_Count_End = new int[DB_Count];

                    int dummy = 0;



                    for (int i = 0; i < Per_DB_Column_Count.Length; i++)
                    {


                        if (i == Per_DB_Column_Count.Length - 1)
                        {
                            //     Per_DB_Column_Count[i] = Getstring.Length - (dummy) - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count;
                            Per_DB_Column_Count[i] = Getstring.Length - (dummy) + 9;
                            Dummy_DB_Count++;
                        }
                        else
                        {
                            Per_DB_Column_Count[i] = DB_Column_Limit;
                            dummy += DB_Column_Limit;
                            Dummy_DB_Count++;
                        }


                        if (i == 0)
                        {
                            Per_DB_Column_Count_Start[i] = 11;
                            Per_DB_Column_Count_End[i] = DB_Column_Limit + TheFirst_Trashes_Header_Count - 1;
                        }
                        else if (i == Per_DB_Column_Count.Length - 1)
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count;
                            Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i] + 9;
                        }
                        else
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count - DB_Column_Limit;
                            Per_DB_Column_Count_End[i] = dummy + TheFirst_Trashes_Header_Count - 1;
                        }


                    }
                }
                //double length = Convert.ToDouble(Clotho_List.Count) / Convert.ToDouble(DB_Column_Limit);
                //double Temp = Math.Truncate(length);

                //int Dummy_DB_Count = 0;

                //if (length > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
                //else DB_Count = Convert.ToInt16(Temp);

                //Per_DB_Column_Count = new int[DB_Count];
                //Per_DB_Column_Count_Start = new int[DB_Count];
                //Per_DB_Column_Count_End = new int[DB_Count];

                //int dummy = 0;

                //for (int i = 0; i < Per_DB_Column_Count.Length; i++)
                //{
                //    if (i == Per_DB_Column_Count.Length - 1)
                //    {
                //        Per_DB_Column_Count[i] = Clotho_List.Count - dummy;
                //        Dummy_DB_Count++;
                //    }
                //    else
                //    {
                //        Per_DB_Column_Count[i] = DB_Column_Limit;
                //        dummy += DB_Column_Limit;
                //        Dummy_DB_Count++;
                //    }

                //    if (i == 0)
                //    {
                //        Per_DB_Column_Count_Start[i] = 0;
                //        Per_DB_Column_Count_End[i] = DB_Column_Limit - 1;
                //    }
                //    else if (i == Per_DB_Column_Count.Length - 1)
                //    {
                //        Per_DB_Column_Count_Start[i] = dummy;
                //        Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i];
                //    }
                //    else
                //    {
                //        Per_DB_Column_Count_Start[i] = dummy - DB_Column_Limit;
                //        Per_DB_Column_Count_End[i] = dummy - 1;
                //    }
                //}

            }

            public void Make_New_header()
            {
                New_Header = new string[Reference_Header.Length];
                for (int j = 0; j < Reference_Header.Length; j++)
                {
                    string Dummy = Reference_Header[j].Replace('.', '_');
                    Dummy = Dummy.Replace('-', '_');

                    New_Header[j] = Dummy;
                }
            }
            public void Edit_Data(string[] GetString, int Data_Row, Data_Class.Data_Editing.INT Data)
            {
                try
                {


                    New_Data = new double[Reference_Header.Length];
                    string Dummy = GetString[0].Replace("PID-", "");
                    New_Data[0] = Convert.ToDouble(Dummy);

                    int j = 1;
                    for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                    {
                        New_Data[j] = Convert.ToDouble(GetString[i]);
                        j++;
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }
            public void Edit_Data(string[] GetString)
            {
                try
                {


                    New_Data = new double[Reference_Header.Length];
                    string Dummy = GetString[0].Replace("PID-", "");
                    New_Data[0] = Convert.ToDouble(Dummy);

                    int j = 1;
                    for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                    {
                        New_Data[j] = Convert.ToDouble(GetString[i]);
                        j++;
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }

            public void Find_Trash_Header(string Value)
            {
                switch (Value.ToUpper().Trim())
                {
                    case "PASSFAIL":
                    case "TIMESTAMP":
                    case "INDEXTIME":
                    case "PARTSN":
                    case "SWBINNAME":
                    case "HWBINNAME":
                        TheEnd_Trashes_Header_Count++;
                        break;
                    case "SBIN":
                    case "HBIN":
                    case "DIE_X":
                    case "DIE_Y":
                    case "SITE":
                    case "TIME":
                    case "TOTAL_TESTS":
                    case "LOT_ID":
                    case "WAFER_ID":
                        TheFirst_Trashes_Header_Count++;
                        break;
                }
            }

            public void TestPlanAddDic(object[,] Data, int Row)
            {
                Dummy_Spec_Band = new Dictionary<string, string>();

                if (Row == 1)
                {
                    for (int i = 0; i < Data.GetLength(1); i++)
                    {
                        Band[i] = Data[1, i + 1].ToString().ToUpper();
                        ConditionCount++;
                    }

                    var tmp = Band;
                    Array.Resize(ref tmp, ConditionCount);
                    Band = tmp;
                }
                else
                {
                    for (int i = 0; i <= ConditionCount - 1; i++)
                    {
                        string Value = "";
                        Value = Data[1, i + 1].ToString();
                        Dummy_Spec_Band.Add(Band[i], Value);
                    }

                    Spec_Band.Add(Dummy_Spec_Band["DEFINE_SPEC"], Dummy_Spec_Band);
                }
            }

            public void Find_Para_by_Defined(string Data, string Spec_Min, string Spec_Max, string Typical, string Convert, string Complience, int index, int Both)
            {

            }
        }
        public class Yield : INT
        {
            Data_Editing Edit = new Data_Editing();
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ManualResetEvent[] ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public string[] Getstring { get; set; }
            public List<string> Reference_Header_List { get; set; }
            public string[] Reference_Header { get; set; }
            public string Data_Table { get; set; }

            public double[] New_HighSpec { get; set; }
            public double[] New_LowSpec { get; set; }
            public string[] New_Header { get; set; }
            public double[] New_Data { get; set; }
            public double[] For_Thread_New_Data { get; set; }
            public string[] For_GetSpec_Header { get; set; }
            public string[] Customer_Clotho_Spec_Data { get; set; }
            public string[] Clotho_Spec_Data { get; set; }
            public string Defined_Spec_Min { get; set; }
            public string Defined_Spec_Max { get; set; }
            public string Defined_Spec_Typical { get; set; }

            public int Defined_Convert_Index { get; set; }
            public string Defined_Convert { get; set; }
            public string Defined_Complience { get; set; }
            public int Defined_Both { get; set; }
            public string Spec_Num { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public int DB_Count { get; set; }
            public int DB_Column_Limit { get; set; }

            public int[] Per_DB_Column_Count { get; set; }
            public int[] Per_DB_Column_Count_Start { get; set; }
            public int[] Per_DB_Column_Count_End { get; set; }

            public bool Flag { get; set; }
            public bool SUBLOT_Falg { get; set; }
            public Dictionary<string, string> Dummy_Spec_Band { get; set; }
            public Dictionary<string, Dictionary<string, string>> Spec_Band { get; set; }

            public Dictionary<string, Spec> Dic_Spec { get; set; }

            public Dictionary<string, SWBIN> SWBIN_Dic { get; set; }
            public List<Clotho_Spec> Clotho_List { get; set; }
            public List<Clotho_Spec> Customor_Clotho_List { get; set; }
            public List<Clotho_Spec> New_Clotho_List { get; set; }
            public List<Clotho_Spec> Clotho_Spcc_List { get; set; }
            public string[] Para { get; set; }
            public string[] Band { get; set; }
            public int ConditionCount { get; set; }

            public List<string> GetTCFDefineSpecNum { get; set; }
            public Dictionary<string, List<string>> GetTCFDefineSpecNum1 { get; set; }
            public Dictionary<string, string> Excel_Combobox { get; set; }
            public int Set_ID { get; set; }
            public int Set_FAIL { get; set; }
            public bool _From_DB { get; set; }

            public string[] Ref_New_Header { get; set; }
            public double[] Ref_New_HighSpec { get; set; }
            public double[] Ref_New_LowSpec { get; set; }
            public string[] Ref_ForAnl_NewMinSpec { get; set; }
            public string[] Ref_ForAnl_NewMaxSpec { get; set; }

            public bool Find_First_Row(string[] Getstring)
            {
                try
                {
                    Flag = false;
                    TheFirst_Trashes_Header_Count = 0;
                    TheEnd_Trashes_Header_Count = 0;
                    if (Getstring[0].ToUpper().Contains("PARAMETER") || Getstring[0].ToUpper().Contains("PARAMETERNAME"))
                    {
                        for (int i = 1; i < 10; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        for (int i = Getstring.Length - 6; i < Getstring.Length; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        int j = 1;
                        Reference_Header = new string[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                        Reference_Header[0] = Getstring[0];
                        int TheFrist = TheFirst_Trashes_Header_Count;
                        int TheEnd = TheEnd_Trashes_Header_Count;
                        for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                        {
                            string Dummy = Getstring[i].Replace('.', '_');
                            Dummy = Dummy.Replace('-', '_');

                            Reference_Header[j] = Getstring[i];
                            j++;
                        }
                        Ref_New_Header = Reference_Header;
                        Flag = true;
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            public bool Find_Spec_Row(string[] Getstring, bool Flag_NewSpec)
            {
                try
                {
                    Flag = false;

                    if (_From_DB)
                    {
                        if (Flag_NewSpec)
                        {
                            int TheFrist = TheFirst_Trashes_Header_Count;
                            int TheEnd = TheEnd_Trashes_Header_Count;
                            int j = 1;
                            for (int i = TheFrist + 1; i < Ref_ForAnl_NewMaxSpec.Length - TheEnd; i++)
                            {
                                New_HighSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMaxSpec[i]);
                                New_LowSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMinSpec[i]);
                                j++;
                            }

                            Ref_New_HighSpec = New_HighSpec;
                            Ref_New_LowSpec = New_LowSpec;
                        }
                    }
                    else
                    {

                        if (Getstring[0].ToUpper().Contains("HIGH") || Getstring[0].ToUpper().Contains("LOW"))
                        {
                            int j = 1;
                            if (Getstring[0].ToUpper().Contains("HIGH"))
                            {
                                New_HighSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                                New_HighSpec[0] = Convert.ToDouble(0);
                                int TheFrist = TheFirst_Trashes_Header_Count;
                                int TheEnd = TheEnd_Trashes_Header_Count;
                                if (Flag_NewSpec)
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        New_HighSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMaxSpec[i]);
                                        j++;
                                    }
                                }
                                else
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        if (Getstring[i] == "") New_HighSpec[j] = 999;
                                        else if (Getstring[i].Contains("G"))
                                        {
                                            string[] split = Getstring[i].ToUpper().Trim().Split('G');
                                            New_HighSpec[j] = Convert.ToDouble(split[1]);
                                        }
                                        else New_HighSpec[j] = Convert.ToDouble(Getstring[i]);
                                        j++;
                                    }
                                }

                                Ref_New_HighSpec = New_HighSpec;
                                Flag = true;
                            }
                            if (Getstring[0].ToUpper().Contains("LOW"))
                            {
                                New_LowSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                                New_LowSpec[0] = Convert.ToDouble(0);
                                int TheFrist = TheFirst_Trashes_Header_Count;
                                int TheEnd = TheEnd_Trashes_Header_Count;

                                if (Flag_NewSpec)
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        New_LowSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMinSpec[i]);
                                        j++;
                                    }
                                }
                                else
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {

                                        if (Getstring[i] == "") New_LowSpec[j] = -999;
                                        else if (Getstring[i].Contains("G"))
                                        {
                                            string[] split = Getstring[i].ToUpper().Trim().Split('G');
                                            New_LowSpec[j] = Convert.ToDouble(split[1]);
                                        }
                                        else New_LowSpec[j] = Convert.ToDouble(Getstring[i]);
                                        j++;
                                    }
                                }

                                Ref_New_LowSpec = New_LowSpec;
                                Flag = true;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }
            public void Find_Cloth_DataFile(string[] Getstring)
            {

                SWBIN_Dic = new Dictionary<string, SWBIN>();

                string[] Values = Getstring[0].Split(',');
                int Values_length = Values.Length;

                int Bin_Count = (Values_length - 5) / 2;

                Clotho_List = new List<Clotho_Spec>();
                Clotho_Spcc_List = new List<Clotho_Spec>();
             //   Customor_Clotho_List = new List<Clotho_Spec>();

                for (int i = 0; i < Getstring.Length; i++)
                {
                    Values = Getstring[i].Split(',');

                    if (Values[0].ToUpper().Contains("SWBIN"))
                    {

                        for (int k = 0; k < 10000; k++)
                        {
                            i++;
                            if (Values[0].Contains("#END"))
                            {
                                break;
                            }
                            else
                            {
                                Values = Getstring[i].Split(',');

                                if (Values[0].Contains("#END"))
                                {
                                    break;
                                }

                                SWBIN SW = new SWBIN(Values[1], Values[3], false);

                                //     if(Values[1].ToUpper().Contains("PASS"))
                                //    {
                                SWBIN_Dic.Add(Values[0], SW);
                                //     }


                                string[] Key = SWBIN_Dic.Keys.ToArray();
                                int index = 0;
                                int Find_Index = 0;

                                foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                                {
                                    //    if (Item.Value.Name.ToUpper().ToString().Contains("PASS"))
                                    //     {
                                    Item.Value.Flag = true;
                                    Find_Index = index;
                                    //      }
                                    index++;
                                }
                            }
                        }
                    }
                    else if (Values[0].ToUpper().Contains("TESTNUM"))
                    {
                        int Start_Index = 5;
                        int l = 1;
                        int j = 0;
                        var Para = new string[40000];
                        Para[0] = "PARAMETER";


                        double[] Min = new double[Bin_Count];
                        double[] Max = new double[Bin_Count];

                        int n = 0;
                        Start_Index = 5;

                        for (j = Start_Index; j < Values.Length; j++)
                        {
                            Min[n] = -9999f;
                            Max[n] = 9999f;
                            j++;
                            n++;
                        }
                        Clotho_Spec Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Clotho_List.Add(Clotho_Spec_Data);



                        n = 0;
                        int Index = 0;
                        string[] Array = SWBIN_Dic.Keys.ToArray();
                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Index++;
                            }
                        }

                        Min = new double[Index];
                        Max = new double[Index];

                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Min[n] = -9999f;
                                Max[n] = 9999f;
                                Start_Index++;
                                n++;
                            }
                        }

                        Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Clotho_Spcc_List.Add(Clotho_Spec_Data);


                        for (int h = i; h < Getstring.Length - 1; h++)
                        {

                            i++;
                            Values = Getstring[i].Split(',');

                            Min = new double[Bin_Count];
                            Max = new double[Bin_Count];

                            n = 0;
                            Start_Index = 5;

                            for (j = Start_Index; j < Values.Length; j++)
                            {
                                Min[n] = Convert.ToDouble(Values[j]);
                                Max[n] = Convert.ToDouble(Values[j + 1]);
                                j++;
                                n++;
                            }
                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Clotho_List.Add(Clotho_Spec_Data);


                            n = 0;
                            Index = 0;
                            Array = SWBIN_Dic.Keys.ToArray();
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Index++;
                                }
                            }

                            Min = new double[Index];
                            Max = new double[Index];
                            Start_Index = 5;
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Min[n] = Convert.ToDouble(Values[Start_Index]);
                                    Max[n] = Convert.ToDouble(Values[Start_Index + 1]);
                                    Start_Index++;
                                    Start_Index++;
                                    n++;
                                }
                            }

                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Clotho_Spcc_List.Add(Clotho_Spec_Data);

                            Para[l] = Values[1];
                            l++;
                        }


                        System.Array.Resize(ref Para, l);
                        Ref_New_Header = Para;

                        Para = null;
                    }
                }



            }
            public void Find_Cloth_DataFile_For_New_Spec(string[] Getstring)
            {


                SWBIN_Dic = new Dictionary<string, SWBIN>();

                string[] Values = Getstring[0].Split(',');
                int Values_length = Values.Length;

                int Bin_Count = (Values_length - 5) / 2;

                Clotho_List = new List<Clotho_Spec>();
             //   Clotho_Spcc_List = new List<Clotho_Spec>();
                Customor_Clotho_List = new List<Clotho_Spec>();

                for (int i = 0; i < Getstring.Length; i++)
                {
                    Values = Getstring[i].Split(',');

                    if (Values[0].ToUpper().Contains("SWBIN"))
                    {

                        for (int k = 0; k < 10000; k++)
                        {
                            i++;
                            if (Values[0].Contains("#END"))
                            {
                                break;
                            }
                            else
                            {
                                Values = Getstring[i].Split(',');

                                if (Values[0].Contains("#END"))
                                {
                                    break;
                                }

                                SWBIN SW = new SWBIN(Values[1], Values[3], false);

                                //     if(Values[1].ToUpper().Contains("PASS"))
                                //    {
                                SWBIN_Dic.Add(Values[0], SW);
                                //     }


                                string[] Key = SWBIN_Dic.Keys.ToArray();
                                int index = 0;
                                int Find_Index = 0;

                                foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                                {
                                    //    if (Item.Value.Name.ToUpper().ToString().Contains("PASS"))
                                    //     {
                                    Item.Value.Flag = true;
                                    Find_Index = index;
                                    //      }
                                    index++;
                                }
                            }
                        }
                    }
                    else if (Values[0].ToUpper().Contains("TESTNUM"))
                    {
                        int Start_Index = 5;
                        int l = 1;
                        int j = 0;
                        var Para = new string[40000];
                        Para[0] = "PARAMETER";


                        double[] Min = new double[Bin_Count];
                        double[] Max = new double[Bin_Count];

                        int n = 0;
                        Start_Index = 5;

                        for (j = Start_Index; j < Values.Length; j++)
                        {
                            Min[n] = -9999f;
                            Max[n] = 9999f;
                            j++;
                            n++;
                        }
                        Clotho_Spec Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Clotho_List.Add(Clotho_Spec_Data);



                        n = 0;
                        int Index = 0;
                        string[] Array = SWBIN_Dic.Keys.ToArray();
                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Index++;
                            }
                        }

                        Min = new double[Index];
                        Max = new double[Index];

                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Min[n] = -9999f;
                                Max[n] = 9999f;
                                Start_Index++;
                                n++;
                            }
                        }

                        Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Customor_Clotho_List.Add(Clotho_Spec_Data);


                        for (int h = i; h < Getstring.Length - 1; h++)
                        {

                            i++;
                            Values = Getstring[i].Split(',');

                            Min = new double[Bin_Count];
                            Max = new double[Bin_Count];

                            n = 0;
                            Start_Index = 5;

                            for (j = Start_Index; j < Values.Length; j++)
                            {
                                Min[n] = Convert.ToDouble(Values[j]);
                                Max[n] = Convert.ToDouble(Values[j + 1]);
                                j++;
                                n++;
                            }
                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Clotho_List.Add(Clotho_Spec_Data);


                            n = 0;
                            Index = 0;
                            Array = SWBIN_Dic.Keys.ToArray();
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Index++;
                                }
                            }

                            Min = new double[Index];
                            Max = new double[Index];
                            Start_Index = 5;
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Min[n] = Convert.ToDouble(Values[Start_Index]);
                                    Max[n] = Convert.ToDouble(Values[Start_Index + 1]);
                                    Start_Index++;
                                    Start_Index++;
                                    n++;
                                }
                            }

                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Customor_Clotho_List.Add(Clotho_Spec_Data);

                            Para[l] = Values[1];
                            l++;
                        }


                        System.Array.Resize(ref Para, l);
                        Ref_New_Header = Para;

                        Para = null;
                    }
                }



            }


            public void Define_DB_Count(string[] Getstring)
            {
                if (Reference_Header.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count < 2000)
                {
                    DB_Count = 1;
                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count[0] = DB_Column_Limit;
                    Per_DB_Column_Count_Start[0] = DB_Column_Limit;
                    Per_DB_Column_Count_End[0] = DB_Column_Limit;

                }
                else
                {
                    double length = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(DB_Column_Limit);
                    double Temp = Math.Truncate(length);

                    int Dummy_DB_Count = 0;

                    if (length > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
                    else DB_Count = Convert.ToInt16(Temp);

                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count_Start = new int[DB_Count];
                    Per_DB_Column_Count_End = new int[DB_Count];

                    int dummy = 0;



                    for (int i = 0; i < Per_DB_Column_Count.Length; i++)
                    {


                        if (i == Per_DB_Column_Count.Length - 1)
                        {
                            //     Per_DB_Column_Count[i] = Getstring.Length - (dummy) - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count;
                            Per_DB_Column_Count[i] = Getstring.Length - (dummy) + 9;
                            Dummy_DB_Count++;
                        }
                        else
                        {
                            Per_DB_Column_Count[i] = DB_Column_Limit;
                            dummy += DB_Column_Limit;
                            Dummy_DB_Count++;
                        }


                        if (i == 0)
                        {
                            Per_DB_Column_Count_Start[i] = 11;
                            Per_DB_Column_Count_End[i] = DB_Column_Limit + TheFirst_Trashes_Header_Count - 1;
                        }
                        else if (i == Per_DB_Column_Count.Length - 1)
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count;
                            Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i] + 9;
                        }
                        else
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count - DB_Column_Limit;
                            Per_DB_Column_Count_End[i] = dummy + TheFirst_Trashes_Header_Count - 1;
                        }


                    }
                }
            }
            public void Make_New_header()
            {
                New_Header = new string[Ref_New_Header.Length];
                for (int j = 0; j < Ref_New_Header.Length; j++)
                {
                    string Dummy = Ref_New_Header[j].Replace('.', '_');
                    Dummy = Dummy.Replace('-', '_');

                    New_Header[j] = Dummy;
                }
            }

            public void Edit_Data(string[] GetString, int Data_Row, Data_Class.Data_Editing.INT Data_Int)
            {
                try
                {

                    if (Data_Row != 0)
                    {
                        Data = Data_Int;
                        Data.Getstring = GetString;
                        ThreadFlags = new ManualResetEvent[Data.DB_Count];
                        Wait = new bool[Data.DB_Count];

                        For_Thread_New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        For_Thread_New_Data[0] = Convert.ToDouble(Dummy);

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            ThreadFlags[i] = new ManualResetEvent(false);
                            ThreadPool.QueueUserWorkItem(new WaitCallback(Thread_Edit_Data), i);
                        }

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            Wait[i] = ThreadFlags[i].WaitOne();
                        }

                    }
                    else if (Data_Row == 0)
                    {
                        New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        New_Data[0] = Convert.ToDouble(Dummy);

                        int j = 1;
                        for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                        {
                            New_Data[j] = Convert.ToDouble(GetString[i]);
                            j++;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }

            public void Thread_Edit_Data(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                int j = 0;
                int q = 0;
                if (i == 0)
                {
                    for (int k = 1; k < DB_Column_Limit; k++)
                    {
                        For_Thread_New_Data[k] = Convert.ToDouble(Getstring[TheFirst_Trashes_Header_Count + 1 + q]);
                        q++;
                    }
                }
                else
                {

                    for (int k = 0; k < Per_DB_Column_Count[i]; k++)
                    {
                        For_Thread_New_Data[(DB_Column_Limit * i) + k] = Convert.ToDouble(Getstring[Per_DB_Column_Count_Start[i] + k]);
                        j++;
                    }
                }
                ThreadFlags[i].Set();
            }

            public void Edit_Data(string[] GetString)
            {
                try
                {


                    New_Data = new double[Reference_Header.Length];
                    string Dummy = GetString[0].Replace("PID-", "");
                    New_Data[0] = Convert.ToDouble(Dummy);

                    int j = 1;
                    for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                    {
                        New_Data[j] = Convert.ToDouble(GetString[i]);
                        j++;
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }
            public void Find_Trash_Header(string Value)
            {
                switch (Value.ToUpper().Trim())
                {
                    case "PASSFAIL":
                    case "TIMESTAMP":
                    case "INDEXTIME":
                    case "PARTSN":
                    case "SWBINNAME":
                    case "HWBINNAME":
                        TheEnd_Trashes_Header_Count++;
                        break;
                    case "SBIN":
                    case "HBIN":
                    case "DIE_X":
                    case "DIE_Y":
                    case "SITE":
                    case "TIME":
                    case "TOTAL_TESTS":
                    case "LOT_ID":
                    case "WAFER_ID":
                        TheFirst_Trashes_Header_Count++;
                        break;
                }
            }

            public void TestPlanAddDic(object[,] Data, int Row)
            {

            }

            public void Find_Para_by_Defined(string Data, string Spec_Min, string Spec_Max, string Typical, string Convert, string Complience, int index, int Both)
            {

            }

        }
        public class BOXPLOT : INT
        {
            Data_Editing Edit = new Data_Editing();
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ManualResetEvent[] ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public string Data_Table { get; set; }

            public string[] Getstring { get; set; }
            public List<string> Reference_Header_List { get; set; }
            public string[] Reference_Header { get; set; }

            public double[] New_HighSpec { get; set; }
            public double[] New_LowSpec { get; set; }
            public string[] New_Header { get; set; }
            public double[] New_Data { get; set; }
            public double[] For_Thread_New_Data { get; set; }
            public string[] For_GetSpec_Header { get; set; }
            public string[] Customer_Clotho_Spec_Data { get; set; }
            public string[] Clotho_Spec_Data { get; set; }
            public string Defined_Spec_Min { get; set; }
            public string Defined_Spec_Max { get; set; }
            public string Defined_Spec_Typical { get; set; }

            public int Defined_Convert_Index { get; set; }
            public string Defined_Convert { get; set; }
            public string Defined_Complience { get; set; }
            public int Defined_Both { get; set; }
            public string Spec_Num { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public int DB_Count { get; set; }
            public int DB_Column_Limit { get; set; }

            public int[] Per_DB_Column_Count { get; set; }
            public int[] Per_DB_Column_Count_Start { get; set; }
            public int[] Per_DB_Column_Count_End { get; set; }

            public bool Flag { get; set; }
            public bool SUBLOT_Falg { get; set; }
            public Dictionary<string, string> Dummy_Spec_Band { get; set; }
            public Dictionary<string, Dictionary<string, string>> Spec_Band { get; set; }

            public Dictionary<string, Spec> Dic_Spec { get; set; }
            public Dictionary<string, SWBIN> SWBIN_Dic { get; set; }
            public List<Clotho_Spec> Clotho_List { get; set; }
            public List<Clotho_Spec> Customor_Clotho_List { get; set; }
            public List<Clotho_Spec> New_Clotho_List { get; set; }
            public List<Clotho_Spec> Clotho_Spcc_List { get; set; }
            public string[] Para { get; set; }
            public string[] Band { get; set; }
            public int ConditionCount { get; set; }

            public List<string> GetTCFDefineSpecNum { get; set; }
            public Dictionary<string, List<string>> GetTCFDefineSpecNum1 { get; set; }
            public Dictionary<string, string> Excel_Combobox { get; set; }
            public int Set_ID { get; set; }
            public int Set_FAIL { get; set; }
            public bool _From_DB { get; set; }

            public string[] Ref_New_Header { get; set; }
            public double[] Ref_New_HighSpec { get; set; }
            public double[] Ref_New_LowSpec { get; set; }
            public string[] Ref_ForAnl_NewMinSpec { get; set; }
            public string[] Ref_ForAnl_NewMaxSpec { get; set; }

            public bool Find_First_Row(string[] Getstring)
            {
                try
                {
                    Flag = false;
                    TheFirst_Trashes_Header_Count = 0;
                    TheEnd_Trashes_Header_Count = 0;
                    if (Getstring[0].ToUpper().Contains("PARAMETER") || Getstring[0].ToUpper().Contains("PARAMETERNAME"))
                    {
                        for (int i = 1; i < 10; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        for (int i = Getstring.Length - 6; i < Getstring.Length; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        int j = 1;
                        Reference_Header = new string[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                        Reference_Header[0] = Getstring[0];
                        int TheFrist = TheFirst_Trashes_Header_Count;
                        int TheEnd = TheEnd_Trashes_Header_Count;
                        for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                        {
                            string Dummy = Getstring[i].Replace('.', '_');
                            Dummy = Dummy.Replace('-', '_');

                            Reference_Header[j] = Getstring[i];
                            j++;
                        }
                        Ref_New_Header = Reference_Header;
                        Flag = true;
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            public bool Find_Spec_Row(string[] Getstring, bool Flag_NewSpec)
            {
                try
                {
                    Flag = false;

                    if (_From_DB)
                    {
                        if (Flag_NewSpec)
                        {
                            int TheFrist = TheFirst_Trashes_Header_Count;
                            int TheEnd = TheEnd_Trashes_Header_Count;
                            int j = 1;
                            for (int i = TheFrist + 1; i < Ref_ForAnl_NewMaxSpec.Length - TheEnd; i++)
                            {
                                New_HighSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMaxSpec[i]);
                                New_LowSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMinSpec[i]);
                                j++;
                            }

                            Ref_New_HighSpec = New_HighSpec;
                            Ref_New_LowSpec = New_LowSpec;
                        }
                    }
                    else
                    {

                        if (Getstring[0].ToUpper().Contains("HIGH") || Getstring[0].ToUpper().Contains("LOW"))
                        {
                            int j = 1;
                            if (Getstring[0].ToUpper().Contains("HIGH"))
                            {
                                New_HighSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                                New_HighSpec[0] = Convert.ToDouble(0);
                                int TheFrist = TheFirst_Trashes_Header_Count;
                                int TheEnd = TheEnd_Trashes_Header_Count;
                                if (Flag_NewSpec)
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        New_HighSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMaxSpec[i]);
                                        j++;
                                    }
                                }
                                else
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        if (Getstring[i] == "") New_HighSpec[j] = 999;
                                        else if (Getstring[i].Contains("G"))
                                        {
                                            string[] split = Getstring[i].ToUpper().Trim().Split('G');
                                            New_HighSpec[j] = Convert.ToDouble(split[1]);
                                        }
                                        else New_HighSpec[j] = Convert.ToDouble(Getstring[i]);
                                        j++;
                                    }
                                }

                                Ref_New_HighSpec = New_HighSpec;
                                Flag = true;
                            }
                            if (Getstring[0].ToUpper().Contains("LOW"))
                            {
                                New_LowSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                                New_LowSpec[0] = Convert.ToDouble(0);
                                int TheFrist = TheFirst_Trashes_Header_Count;
                                int TheEnd = TheEnd_Trashes_Header_Count;

                                if (Flag_NewSpec)
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        New_LowSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMinSpec[i]);
                                        j++;
                                    }
                                }
                                else
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {

                                        if (Getstring[i] == "") New_LowSpec[j] = -999;
                                        else if (Getstring[i].Contains("G"))
                                        {
                                            string[] split = Getstring[i].ToUpper().Trim().Split('G');
                                            New_LowSpec[j] = Convert.ToDouble(split[1]);
                                        }
                                        else New_LowSpec[j] = Convert.ToDouble(Getstring[i]);
                                        j++;
                                    }
                                }

                                Ref_New_LowSpec = New_LowSpec;
                                Flag = true;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }
            public void Find_Cloth_DataFile(string[] Getstring)
            {

            }
            public void Find_Cloth_DataFile_For_New_Spec(string[] Customer)
            {

            }
            //public void Define_DB_Count(string[] Getstring)
            //{
            //    if (Reference_Header.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count < 2000)
            //    {
            //        DB_Count = 1;
            //        Per_DB_Column_Count = new int[DB_Count];
            //        Per_DB_Column_Count[0] = DB_Column_Limit;
            //        Per_DB_Column_Count_Start[0] = DB_Column_Limit;
            //        Per_DB_Column_Count_End[0] = DB_Column_Limit;

            //    }
            //    else
            //    {
            //        double length = 0f;
            //        for (int j = 0; j < 20; j++)
            //        {
            //            length = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(j);
            //            if (length < 2000)
            //            {
            //                break;
            //            }

            //        }
            //        double Get_Count = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(length);
            //        double Temp = Math.Truncate(Get_Count);

            //        int Dummy_DB_Count = 0;

            //        if (Get_Count > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
            //        else DB_Count = Convert.ToInt16(Temp);

            //        Per_DB_Column_Count = new int[DB_Count];
            //        Per_DB_Column_Count_Start = new int[DB_Count];
            //        Per_DB_Column_Count_End = new int[DB_Count];

            //        int dummy = 0;


            //        for (int i = 0; i < Per_DB_Column_Count.Length; i++)
            //        {
            //            if (i == Per_DB_Column_Count.Length - 1)
            //            {
            //                Per_DB_Column_Count[i] = Getstring.Length - (dummy) - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count;
            //                Dummy_DB_Count++;
            //            }
            //            else
            //            {
            //                Per_DB_Column_Count[i] = Convert.ToInt16(Math.Truncate(length)) + 1;
            //                dummy += Convert.ToInt16(Math.Truncate(length)) + 1;
            //                Dummy_DB_Count++;
            //            }

            //            if (i == 0)
            //            {
            //                Per_DB_Column_Count_Start[i] = TheFirst_Trashes_Header_Count + 1;
            //                Per_DB_Column_Count_End[i] = Convert.ToInt16(Math.Truncate(length)) + 1 + TheFirst_Trashes_Header_Count - 1;
            //            }
            //            else if (i == Per_DB_Column_Count.Length - 1)
            //            {
            //                Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count;
            //                Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i] + TheFirst_Trashes_Header_Count - 1;
            //            }
            //            else
            //            {
            //                Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count - Convert.ToInt16(Math.Truncate(length)) + 1;
            //                Per_DB_Column_Count_End[i] = dummy + TheFirst_Trashes_Header_Count - 1;
            //            }
            //        }

            //        DB_Column_Limit = Per_DB_Column_Count[0];
            //    }
            //}

            public void Define_DB_Count(string[] Getstring)
            {
                if (Reference_Header.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count < 2000)
                {
                    DB_Count = 1;
                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count[0] = DB_Column_Limit;
                    Per_DB_Column_Count_Start[0] = DB_Column_Limit;
                    Per_DB_Column_Count_End[0] = DB_Column_Limit;

                }
                else
                {
                    double length = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(DB_Column_Limit);
                    double Temp = Math.Truncate(length);

                    int Dummy_DB_Count = 0;

                    if (length > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
                    else DB_Count = Convert.ToInt16(Temp);

                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count_Start = new int[DB_Count];
                    Per_DB_Column_Count_End = new int[DB_Count];

                    int dummy = 0;

                    for (int i = 0; i < Per_DB_Column_Count.Length; i++)
                    {
                        if (i == Per_DB_Column_Count.Length - 1)
                        {
                            Per_DB_Column_Count[i] = Getstring.Length - (dummy) - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count;
                            Dummy_DB_Count++;
                        }
                        else
                        {
                            Per_DB_Column_Count[i] = DB_Column_Limit;
                            dummy += DB_Column_Limit;
                            Dummy_DB_Count++;
                        }

                        if (i == 0)
                        {
                            Per_DB_Column_Count_Start[i] = TheFirst_Trashes_Header_Count + 1;
                            Per_DB_Column_Count_End[i] = DB_Column_Limit + TheFirst_Trashes_Header_Count - 1;
                        }
                        else if (i == Per_DB_Column_Count.Length - 1)
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count;
                            Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i] + TheFirst_Trashes_Header_Count - 1;
                        }
                        else
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count - DB_Column_Limit;
                            Per_DB_Column_Count_End[i] = dummy + TheFirst_Trashes_Header_Count - 1;
                        }
                    }
                }
            }

            public void Make_New_header()
            {
                New_Header = new string[Reference_Header.Length];
                for (int j = 0; j < Reference_Header.Length; j++)
                {
                    string Dummy = Reference_Header[j].Replace('.', '_');
                    Dummy = Dummy.Replace('-', '_');

                    New_Header[j] = Dummy;
                }
            }

            public void Edit_Data(string[] GetString, int Data_Row, Data_Class.Data_Editing.INT Data_Int)
            {
                try
                {

                    if (Data_Row != 0)
                    {
                        Data = Data_Int;
                        Data.Getstring = GetString;
                        ThreadFlags = new ManualResetEvent[Data.DB_Count];
                        Wait = new bool[Data.DB_Count];

                        For_Thread_New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        For_Thread_New_Data[0] = Convert.ToDouble(Dummy);

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            ThreadFlags[i] = new ManualResetEvent(false);
                            ThreadPool.QueueUserWorkItem(new WaitCallback(Thread_Edit_Data), i);
                        }

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            Wait[i] = ThreadFlags[i].WaitOne();
                        }

                    }
                    else if (Data_Row == 0)
                    {
                        New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        New_Data[0] = Convert.ToDouble(Dummy);

                        int j = 1;
                        for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                        {
                            New_Data[j] = Convert.ToDouble(GetString[i]);
                            j++;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }

            public void Thread_Edit_Data(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                int j = 0;
                int q = 0;
                if (i == 0)
                {
                    for (int k = 1; k < DB_Column_Limit; k++)
                    {
                        For_Thread_New_Data[k] = Convert.ToDouble(Getstring[TheFirst_Trashes_Header_Count + 1 + q]);
                        q++;
                    }
                }
                else
                {

                    for (int k = 0; k < Per_DB_Column_Count[i]; k++)
                    {
                        For_Thread_New_Data[(DB_Column_Limit * i) + k] = Convert.ToDouble(Getstring[Per_DB_Column_Count_Start[i] + k]);
                        j++;
                    }
                }
                ThreadFlags[i].Set();
            }

            public void Edit_Data(string[] GetString)
            {
                try
                {


                    New_Data = new double[Reference_Header.Length];
                    string Dummy = GetString[0].Replace("PID-", "");
                    New_Data[0] = Convert.ToDouble(Dummy);

                    int j = 1;
                    for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                    {
                        New_Data[j] = Convert.ToDouble(GetString[i]);
                        j++;
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }
            public void Find_Trash_Header(string Value)
            {
                switch (Value.ToUpper().Trim())
                {
                    case "PASSFAIL":
                    case "TIMESTAMP":
                    case "INDEXTIME":
                    case "PARTSN":
                    case "SWBINNAME":
                    case "HWBINNAME":
                        TheEnd_Trashes_Header_Count++;
                        break;
                    case "SBIN":
                    case "HBIN":
                    case "DIE_X":
                    case "DIE_Y":
                    case "SITE":
                    case "TIME":
                    case "TOTAL_TESTS":
                    case "LOT_ID":
                    case "WAFER_ID":
                        TheFirst_Trashes_Header_Count++;
                        break;
                }
            }

            public void TestPlanAddDic(object[,] Data, int Row)
            {

            }

            public void Find_Para_by_Defined(string Data, string Spec_Min, string Spec_Max, string Typical, string Convert, string Complience, int index, int Both)
            {

            }

        }
        public class GETSPEC : INT
        {
            Data_Editing Edit = new Data_Editing();
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ManualResetEvent[] ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public string[] Getstring { get; set; }
            public List<string> Reference_Header_List { get; set; }
            public string[] Reference_Header { get; set; }

            public string Data_Table { get; set; }

            public double[] New_HighSpec { get; set; }
            public double[] New_LowSpec { get; set; }
            public string[] New_Header { get; set; }
            public double[] For_Thread_New_Data { get; set; }
            public double[] New_Data { get; set; }

            public string[] For_GetSpec_Header { get; set; }
            public string[] Customer_Clotho_Spec_Data { get; set; }
            public string[] Clotho_Spec_Data { get; set; }
            public string Defined_Spec_Min { get; set; }
            public string Defined_Spec_Max { get; set; }
            public string Defined_Spec_Typical { get; set; }
            public int Defined_Convert_Index { get; set; }
            public string Defined_Convert { get; set; }
            public string Defined_Complience { get; set; }
            public int Defined_Both { get; set; }
            public string Spec_Num { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public int DB_Count { get; set; }
            public int DB_Column_Limit { get; set; }

            public int[] Per_DB_Column_Count { get; set; }
            public int[] Per_DB_Column_Count_Start { get; set; }
            public int[] Per_DB_Column_Count_End { get; set; }

            public bool Flag { get; set; }
            public bool SUBLOT_Falg { get; set; }
            public Dictionary<string, string> Dummy_Spec_Band { get; set; }
            public Dictionary<string, Dictionary<string, string>> Spec_Band { get; set; }

            public Dictionary<string, Spec> Dic_Spec { get; set; }
            public Dictionary<string, SWBIN> SWBIN_Dic { get; set; }

            public List<Clotho_Spec> Clotho_List { get; set; }
            public List<Clotho_Spec> Customor_Clotho_List { get; set; }
            public List<Clotho_Spec> New_Clotho_List { get; set; }
            public List<Clotho_Spec> Clotho_Spcc_List { get; set; }
            public string[] Para { get; set; }
            public string[] Band { get; set; }
            public int ConditionCount { get; set; }

            public List<string> GetTCFDefineSpecNum { get; set; }
            public Dictionary<string, List<string>> GetTCFDefineSpecNum1 { get; set; }
            public Dictionary<string, string> Excel_Combobox { get; set; }
            public int Set_ID { get; set; }
            public int Set_FAIL { get; set; }
            public bool _From_DB { get; set; }

            public string[] Ref_New_Header { get; set; }
            public double[] Ref_New_HighSpec { get; set; }
            public double[] Ref_New_LowSpec { get; set; }
            public string[] Ref_ForAnl_NewMinSpec { get; set; }
            public string[] Ref_ForAnl_NewMaxSpec { get; set; }

            public bool Find_First_Row(string[] Getstring)
            {
                try
                {
                    if (Getstring[0].ToUpper().Contains("TESTNUMBER") || Getstring[0].ToUpper().Contains("PARAMETERNAME"))
                    {
                        for (int i = 1; i < 10; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        for (int i = Getstring.Length - 6; i < Getstring.Length; i++)
                        {
                            Find_Trash_Header(Getstring[i]);
                        }

                        int j = 1;
                        Reference_Header = new string[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                        Reference_Header[0] = Getstring[0];
                        for (int i = TheFirst_Trashes_Header_Count + 1; i < Getstring.Length - TheEnd_Trashes_Header_Count; i++)
                        {
                            string Dummy = Getstring[i].Replace('.', '_');
                            Dummy = Dummy.Replace('-', '_');

                            Reference_Header[j] = Getstring[i];
                            j++;
                        }
                        Ref_New_Header = Reference_Header;
                        Flag = true;
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            public bool Find_Spec_Row(string[] Getstring, bool Flag)
            {
                try
                {
                    Flag = false;
                    if (Getstring[0].ToUpper().Contains("HIGH") || Getstring[0].ToUpper().Contains("LOW"))
                    {
                        int j = 1;
                        if (Getstring[0].ToUpper().Contains("HIGH"))
                        {
                            New_HighSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_HighSpec[0] = Convert.ToDouble(0);
                            for (int i = TheFirst_Trashes_Header_Count + 1; i < Getstring.Length - TheEnd_Trashes_Header_Count; i++)
                            {
                                New_HighSpec[j] = Convert.ToDouble(Getstring[i]);
                                j++;
                            }
                            Ref_New_HighSpec = New_HighSpec;
                            Flag = true;
                        }
                        if (Getstring[0].ToUpper().Contains("LOW"))
                        {
                            New_LowSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_LowSpec[0] = Convert.ToDouble(0);
                            for (int i = TheFirst_Trashes_Header_Count + 1; i < Getstring.Length - TheEnd_Trashes_Header_Count; i++)
                            {
                                New_LowSpec[j] = Convert.ToDouble(Getstring[i]);
                                j++;
                            }
                            Ref_New_LowSpec = New_LowSpec;
                            Flag = true;
                        }
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }
            public void Find_Cloth_DataFile(string[] Getstring)
            {
                SWBIN_Dic = new Dictionary<string, SWBIN>();

                string[] Values = Getstring[0].Split(',');
                int Values_length = Values.Length;

                int Bin_Count = (Values_length - 5) / 2;

                Clotho_List = new List<Clotho_Spec>();
                Clotho_Spcc_List = new List<Clotho_Spec>();


                for (int i = 0; i < Getstring.Length; i++)
                {
                    Values = Getstring[i].Split(',');

                    if (Values[0].ToUpper().Contains("SWBIN"))
                    {

                        for (int k = 0; k < 10000; k++)
                        {
                            i++;
                            if (Values[0].Contains("#END"))
                            {
                                break;
                            }
                            else
                            {
                                Values = Getstring[i].Split(',');
                                SWBIN SW = new SWBIN(Values[1], Values[3], false);

                                if (Values[1].ToUpper().Contains("PASS"))
                                {
                                    SWBIN_Dic.Add(Values[0], SW);
                                }


                                string[] Key = SWBIN_Dic.Keys.ToArray();
                                int index = 0;
                                int Find_Index = 0;

                                foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                                {
                                    if (Item.Value.Name.ToUpper().ToString().Contains("PASS"))
                                    {
                                        Item.Value.Flag = true;
                                        Find_Index = index;
                                    }
                                    index++;
                                }
                            }
                        }
                    }
                    else if (Values[0].ToUpper().Contains("TESTNUM"))
                    {
                        int Start_Index = 5;
                        int l = 1;
                        int j = 0;
                        var Para = new string[40000];
                        Para[0] = "PARAMETER";


                        double[] Min = new double[Bin_Count];
                        double[] Max = new double[Bin_Count];

                        int n = 0;
                        Start_Index = 5;

                        for (j = Start_Index; j < Values.Length; j++)
                        {
                            Min[n] = -9999f;
                            Max[n] = 9999f;
                            j++;
                            n++;
                        }
                        Clotho_Spec Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Clotho_List.Add(Clotho_Spec_Data);

                        n = 0;
                        int Index = 0;
                        string[] Array = SWBIN_Dic.Keys.ToArray();
                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Index++;
                            }
                        }

                        Min = new double[Index];
                        Max = new double[Index];

                        foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                        {
                            if (Item.Value.Flag == true)
                            {
                                Min[n] = -9999f;
                                Max[n] = 9999f;
                                Start_Index++;
                                n++;
                            }
                        }

                        Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                        Clotho_Spcc_List.Add(Clotho_Spec_Data);


                        for (int h = i; h < Getstring.Length - 1; h++)
                        {

                            i++;
                            Values = Getstring[i].Split(',');

                            Min = new double[Bin_Count];
                            Max = new double[Bin_Count];

                            n = 0;
                            Start_Index = 5;

                            for (j = Start_Index; j < Values.Length; j++)
                            {
                                Min[n] = Convert.ToDouble(Values[j]);
                                Max[n] = Convert.ToDouble(Values[j + 1]);
                                j++;
                                n++;
                            }
                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Clotho_List.Add(Clotho_Spec_Data);

                            n = 0;
                            Index = 0;
                            Array = SWBIN_Dic.Keys.ToArray();
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Index++;
                                }
                            }

                            Min = new double[Index];
                            Max = new double[Index];
                            Start_Index = 5;
                            foreach (KeyValuePair<string, SWBIN> Item in SWBIN_Dic)
                            {
                                if (Item.Value.Flag == true)
                                {
                                    Min[n] = Convert.ToDouble(Values[Start_Index]);
                                    Max[n] = Convert.ToDouble(Values[Start_Index + 1]);
                                    Start_Index++;
                                    Start_Index++;
                                    n++;
                                }
                            }

                            Clotho_Spec_Data = new Clotho_Spec(Min, Max);
                            Clotho_Spcc_List.Add(Clotho_Spec_Data);

                            Para[l] = Values[1];
                            l++;
                        }


                        System.Array.Resize(ref Para, l);
                        Ref_New_Header = Para;

                        For_GetSpec_Header = new string[Ref_New_Header.Length];

                        for (int s = 0; s < Ref_New_Header.Length; s++)
                        {
                            string[] split = Ref_New_Header[s].Split('_');
                            For_GetSpec_Header[s] = split[split.Length - 1].Replace('-', '_');
                        }

                        Para = null;
                    }
                }


            }
            public void Find_Cloth_DataFile_For_New_Spec(string[] Customer)
            {

            }
            public void Define_DB_Count(string[] Getstring)
            {
                double length = Convert.ToDouble(Clotho_List.Count) / Convert.ToDouble(DB_Column_Limit);
                double Temp = Math.Truncate(length);

                int Dummy_DB_Count = 0;

                if (length > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
                else DB_Count = Convert.ToInt16(Temp);

                Per_DB_Column_Count = new int[DB_Count];
                Per_DB_Column_Count_Start = new int[DB_Count];
                Per_DB_Column_Count_End = new int[DB_Count];

                int dummy = 0;

                for (int i = 0; i < Per_DB_Column_Count.Length; i++)
                {
                    if (i == Per_DB_Column_Count.Length - 1)
                    {
                        Per_DB_Column_Count[i] = Clotho_List.Count - dummy;
                        Dummy_DB_Count++;
                    }
                    else
                    {
                        Per_DB_Column_Count[i] = DB_Column_Limit;
                        dummy += DB_Column_Limit;
                        Dummy_DB_Count++;
                    }

                    if (i == 0)
                    {
                        Per_DB_Column_Count_Start[i] = 0;
                        Per_DB_Column_Count_End[i] = DB_Column_Limit - 1;
                    }
                    else if (i == Per_DB_Column_Count.Length - 1)
                    {
                        Per_DB_Column_Count_Start[i] = dummy;
                        Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i];
                    }
                    else
                    {
                        Per_DB_Column_Count_Start[i] = dummy - DB_Column_Limit;
                        Per_DB_Column_Count_End[i] = dummy - 1;
                    }
                }
            }

            public void Make_New_header()
            {
                New_Header = new string[Reference_Header.Length];
                for (int j = 0; j < Reference_Header.Length; j++)
                {
                    string Dummy = Reference_Header[j].Replace('.', '_');
                    Dummy = Dummy.Replace('-', '_');

                    New_Header[j] = Dummy;
                }
            }

            public void Edit_Data(string[] GetString, int Data_Row, Data_Class.Data_Editing.INT Data)
            {

                New_Data = new double[Reference_Header.Length];
                string Dummy = GetString[0].Replace("PID-", "");
                New_Data[0] = Convert.ToDouble(Dummy);

                int j = 1;
                for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                {
                    //  New_Data[j] = Convert.ToDouble(GetString[i]);
                    j++;
                }


            }

            public void Edit_Data(string[] GetString)
            {

                New_Data = new double[Reference_Header.Length];
                string Dummy = GetString[0].Replace("PID-", "");
                New_Data[0] = Convert.ToDouble(Dummy);

                int j = 1;
                for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                {
                    //  New_Data[j] = Convert.ToDouble(GetString[i]);
                    j++;
                }


            }

            public void Find_Trash_Header(string Value)
            {
                switch (Value.ToUpper().Trim())
                {
                    case "PASSFAIL":
                    case "TIMESTAMP":
                    case "INDEXTIME":
                    case "PARTSN":
                    case "SWBINNAME":
                    case "HWBINNAME":
                        TheEnd_Trashes_Header_Count++;
                        break;
                    case "SBIN":
                    case "HBIN":
                    case "DIE_X":
                    case "DIE_Y":
                    case "SITE":
                    case "TIME":
                    case "TOTAL_TESTS":
                    case "LOT_ID":
                    case "WAFER_ID":
                        TheFirst_Trashes_Header_Count++;
                        break;
                }
            }

            public void TestPlanAddDic(object[,] Data, int Row)
            {
                Dummy_Spec_Band = new Dictionary<string, string>();

                if (Row == 1)
                {
                    for (int i = 0; i < Data.GetLength(1); i++)
                    {
                        Band[i] = Data[1, i + 1].ToString().ToUpper();
                        ConditionCount++;
                    }

                    var tmp = Band;
                    Array.Resize(ref tmp, ConditionCount);
                    Band = tmp;
                }
                else
                {
                    for (int i = 0; i <= ConditionCount - 1; i++)
                    {
                        string Value = "";
                        Value = Data[1, i + 1].ToString();
                        Dummy_Spec_Band.Add(Band[i], Value);
                    }

                    Spec_Band.Add(Dummy_Spec_Band["DEFINE_SPEC"], Dummy_Spec_Band);
                }
            }

            public void Find_Para_by_Defined(string Data, string Spec_Min, string Spec_Max, string Typical, string Convert, string Complience, int index, int Both)
            {
                Defined_Spec_Min = Spec_Min;
                Defined_Spec_Max = Spec_Max;
                Defined_Spec_Typical = Typical;
                Defined_Convert = Convert;
                Defined_Complience = Complience;
                Defined_Convert_Index = index;
                Defined_Both = Both;
                Spec_Num = Data.Replace('-', '_');

                ThreadFlags = new ManualResetEvent[DB_Count];
                Wait = new bool[DB_Count];


                for (int i = 0; i < DB_Count; i++)
                {
                    Find_Para_By_Defined_Thread(i);
                    ///  ThreadFlags[i] = new ManualResetEvent(false);
                    //  ThreadPool.QueueUserWorkItem(new WaitCallback(Find_Para_By_Defined_Thread), i);
                }
                for (int i = 0; i < DB_Count; i++)
                {
                    //   Wait[i] = ThreadFlags[i].WaitOne();
                }
            }

            public void Find_Para_By_Defined_Thread(Object threadContext)
            {
                int i = (int)threadContext;


                for (int j = 0; j < Per_DB_Column_Count[i]; j++)
                {
                    if (For_GetSpec_Header[DB_Column_Limit * i + j].ToUpper() != "X")
                    {
                        if (For_GetSpec_Header[DB_Column_Limit * i + j].ToUpper() == Spec_Num.ToUpper().Trim())
                        {

                            if (Defined_Spec_Min == "TBD")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Min = "-88888";
                            }
                            else if (Defined_Spec_Min != "")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Min = Defined_Spec_Min;
                            }

                            if (Defined_Spec_Max == "TBD")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Max = "88888";
                            }
                            else if (Defined_Spec_Max != "")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Max = Defined_Spec_Max;
                            }

                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].SpecNumber = Spec_Num;

                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Convert = Defined_Convert;
                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Index = Defined_Convert_Index;
                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Both = Defined_Both;
                        }
                    }
                }
                ThreadFlags[i].Set();
            }

            public void Find_Para_By_Defined_Thread(int i)
            {

                for (int j = 0; j < Per_DB_Column_Count[i]; j++)
                {
                    if (For_GetSpec_Header[DB_Column_Limit * i + j].ToUpper() != "X")
                    {
                        if (For_GetSpec_Header[DB_Column_Limit * i + j].ToUpper() == Spec_Num.ToUpper().Trim())
                        {

                            if (Defined_Spec_Min == "TBD")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Min = "-88888";
                            }
                            else if (Defined_Spec_Min != "")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Min = Defined_Spec_Min;
                            }

                            if (Defined_Spec_Max == "TBD")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Min = "88888";
                            }
                            else if (Defined_Spec_Max != "")
                            {
                                Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Max = Defined_Spec_Max;
                            }

                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Typical = Defined_Spec_Typical;

                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].SpecNumber = Spec_Num;

                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Convert = Defined_Convert;
                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Complience = Defined_Complience;

                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Index = Defined_Convert_Index;
                            Dic_Spec[Reference_Header[DB_Column_Limit * i + j]].Both = Defined_Both;
                        }
                    }
                }

            }


        }
        public class MERGE : INT
        {
            Data_Editing Edit = new Data_Editing();
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ManualResetEvent[] ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public string[] Getstring { get; set; }
            public List<string> Reference_Header_List { get; set; }
            public string[] Reference_Header { get; set; }

            public string Data_Table { get; set; }

            public double[] New_HighSpec { get; set; }
            public double[] New_LowSpec { get; set; }
            public string[] New_Header { get; set; }
            public double[] New_Data { get; set; }
            public double[] For_Thread_New_Data { get; set; }
            public string[] For_GetSpec_Header { get; set; }
            public string[] Customer_Clotho_Spec_Data { get; set; }
            public string[] Clotho_Spec_Data { get; set; }
            public string Defined_Spec_Min { get; set; }
            public string Defined_Spec_Max { get; set; }
            public string Defined_Spec_Typical { get; set; }
            public int Defined_Convert_Index { get; set; }
            public string Defined_Convert { get; set; }
            public string Defined_Complience { get; set; }
            public int Defined_Both { get; set; }
            public string Spec_Num { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public int DB_Count { get; set; }
            public int DB_Column_Limit { get; set; }

            public int[] Per_DB_Column_Count { get; set; }
            public int[] Per_DB_Column_Count_Start { get; set; }
            public int[] Per_DB_Column_Count_End { get; set; }

            public bool Flag { get; set; }
            public bool SUBLOT_Falg { get; set; }

            public Dictionary<string, string> Dummy_Spec_Band { get; set; }
            public Dictionary<string, Dictionary<string, string>> Spec_Band { get; set; }

            public Dictionary<string, Spec> Dic_Spec { get; set; }
            public Dictionary<string, SWBIN> SWBIN_Dic { get; set; }

            public List<Clotho_Spec> Clotho_List { get; set; }
            public List<Clotho_Spec> Customor_Clotho_List { get; set; }
            public List<Clotho_Spec> New_Clotho_List { get; set; }
            public List<Clotho_Spec> Clotho_Spcc_List { get; set; }
            public string[] Para { get; set; }
            public string[] Band { get; set; }
            public int ConditionCount { get; set; }

            public List<string> GetTCFDefineSpecNum { get; set; }
            public Dictionary<string, List<string>> GetTCFDefineSpecNum1 { get; set; }
            public Dictionary<string, string> Excel_Combobox { get; set; }
            public int Set_ID { get; set; }
            public int Set_FAIL { get; set; }
            public bool _From_DB { get; set; }

            public string[] Ref_New_Header { get; set; }
            public double[] Ref_New_HighSpec { get; set; }
            public  double[] Ref_New_LowSpec { get; set; }
            public  string[] Ref_ForAnl_NewMinSpec { get; set; }
            public string[] Ref_ForAnl_NewMaxSpec { get; set; }




            public bool Find_First_Row(string[] Getstring)
            {
                try
                {
                    Flag = false;
                    TheFirst_Trashes_Header_Count = 0;
                    TheEnd_Trashes_Header_Count = 0;
                    if (Getstring[0].ToUpper().Contains("PARAMETER") || Getstring[0].ToUpper().Contains("PARAMETERNAME"))
                    {
                        //for (int i = 1; i < 10; i++)
                        //{
                        //    Find_Trash_Header(Getstring[i]);
                        //}

                        //for (int i = Getstring.Length - 6; i < Getstring.Length; i++)
                        //{
                        //    Find_Trash_Header(Getstring[i]);
                        //}

                        int j = 1;
                        Reference_Header = new string[Getstring.Length];
                        Reference_Header[0] = Getstring[0];
                        int TheFrist = TheFirst_Trashes_Header_Count;
                        int TheEnd = TheEnd_Trashes_Header_Count;
                        for (int i = 1; i < Getstring.Length; i++)
                        {
                            string Dummy = Getstring[i].Replace('.', '_');
                            Dummy = Dummy.Replace('-', '_');

                            Reference_Header[j] = Getstring[i];
                            j++;
                        }
                        //  Reference_Header[j] = "SUBLOT";
                        Ref_New_Header = Reference_Header;
                        Flag = true;
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            public bool Find_Spec_Row(string[] Getstring, bool Flag_NewSpec)
            {
                try
                {
                    Flag = false;
                    if (Getstring[0].ToUpper().Contains("HIGH") || Getstring[0].ToUpper().Contains("LOW"))
                    {
                        int j = 1;
                        if (Getstring[0].ToUpper().Contains("HIGH"))
                        {
                            New_HighSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_HighSpec[0] = Convert.ToDouble(0);
                            int TheFrist = TheFirst_Trashes_Header_Count;
                            int TheEnd = TheEnd_Trashes_Header_Count;
                            if (Flag_NewSpec)
                            {
                                for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                {
                                    New_HighSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMaxSpec[i]);
                                    j++;
                                }
                            }
                            else
                            {
                                try
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        // if (Getstring[i] == "") New_HighSpec[j] = 999;
                                        if (Getstring[i] == "") { }
                                        else New_HighSpec[j] = Convert.ToDouble(Getstring[i]);
                                        j++;
                                    }
                                }
                                catch
                                {

                                }
                            }

                         //   Data_Editing.New_HighSpec = New_HighSpec;
                            Flag = true;
                        }
                        if (Getstring[0].ToUpper().Contains("LOW"))
                        {
                            New_LowSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_LowSpec[0] = Convert.ToDouble(0);
                            int TheFrist = TheFirst_Trashes_Header_Count;
                            int TheEnd = TheEnd_Trashes_Header_Count;

                            if (Flag_NewSpec)
                            {
                                for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                {
                                    New_LowSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMinSpec[i]);
                                    j++;
                                }
                            }
                            else
                            {
                                for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                {
                                    //if (Getstring[i] == "") New_LowSpec[j] = -999;
                                    if (Getstring[i] == "") { }
                                    else New_LowSpec[j] = Convert.ToDouble(Getstring[i]);
                                    j++;
                                }
                            }

                        //    Data_Editing.New_LowSpec = New_LowSpec;
                            Flag = true;
                        }
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            //public void Define_DB_Count(string[] Getstring)
            //{
            //    if (Reference_Header.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count < 2000)
            //    {
            //        DB_Count = 1;
            //        Per_DB_Column_Count = new int[DB_Count];
            //        Per_DB_Column_Count[0] = DB_Column_Limit;
            //        Per_DB_Column_Count_Start[0] = DB_Column_Limit;
            //        Per_DB_Column_Count_End[0] = DB_Column_Limit;

            //    }
            //    else
            //    {
            //        double length = 0f;
            //        for (int j = 0; j < 20; j++)
            //        {
            //            length = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(j);
            //            if (length < 2000)
            //            {
            //                break;
            //            }

            //        }
            //        double Get_Count = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(length);
            //        double Temp = Math.Truncate(Get_Count);

            //        int Dummy_DB_Count = 0;

            //        if (Get_Count > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
            //        else DB_Count = Convert.ToInt16(Temp);

            //        Per_DB_Column_Count = new int[DB_Count];
            //        Per_DB_Column_Count_Start = new int[DB_Count];
            //        Per_DB_Column_Count_End = new int[DB_Count];

            //        int dummy = 0;


            //        for (int i = 0; i < Per_DB_Column_Count.Length; i++)
            //        {
            //            if (i == Per_DB_Column_Count.Length - 1)
            //            {
            //                Per_DB_Column_Count[i] = Getstring.Length - (dummy) - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count;
            //                Dummy_DB_Count++;
            //            }
            //            else
            //            {
            //                Per_DB_Column_Count[i] = Convert.ToInt16(Math.Truncate(length)) + 1;
            //                dummy += Convert.ToInt16(Math.Truncate(length)) + 1;
            //                Dummy_DB_Count++;
            //            }

            //            if (i == 0)
            //            {
            //                Per_DB_Column_Count_Start[i] = TheFirst_Trashes_Header_Count + 1;
            //                Per_DB_Column_Count_End[i] = Convert.ToInt16(Math.Truncate(length)) + 1 + TheFirst_Trashes_Header_Count - 1;
            //            }
            //            else if (i == Per_DB_Column_Count.Length - 1)
            //            {
            //                Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count;
            //                Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i] + TheFirst_Trashes_Header_Count - 1;
            //            }
            //            else
            //            {
            //                Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count - Convert.ToInt16(Math.Truncate(length)) + 1;
            //                Per_DB_Column_Count_End[i] = dummy + TheFirst_Trashes_Header_Count - 1;
            //            }
            //        }

            //        DB_Column_Limit = Per_DB_Column_Count[0];
            //    }
            //}

            public void Find_Cloth_DataFile(string[] Getstring)
            {

            }
            public void Find_Cloth_DataFile_For_New_Spec(string[] Customer)
            {

            }
            public void Define_DB_Count(string[] Getstring)
            {
                if (Reference_Header.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count < 2000)
                {
                    DB_Count = 1;
                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count[0] = DB_Column_Limit;
                    Per_DB_Column_Count_Start[0] = DB_Column_Limit;
                    Per_DB_Column_Count_End[0] = DB_Column_Limit;

                }
                else
                {
                    double length = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(DB_Column_Limit);
                    double Temp = Math.Truncate(length);

                    int Dummy_DB_Count = 0;

                    if (length > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
                    else DB_Count = Convert.ToInt16(Temp);

                    Per_DB_Column_Count = new int[DB_Count];
                    Per_DB_Column_Count_Start = new int[DB_Count];
                    Per_DB_Column_Count_End = new int[DB_Count];

                    int dummy = 0;



                    for (int i = 0; i < Per_DB_Column_Count.Length; i++)
                    {

                        if (i == 6)
                        {

                        }

                        if (i == Per_DB_Column_Count.Length - 1)
                        {
                            //     Per_DB_Column_Count[i] = Getstring.Length - (dummy) - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count;
                            Per_DB_Column_Count[i] = Getstring.Length - (dummy);
                            Dummy_DB_Count++;
                        }
                        else
                        {
                            Per_DB_Column_Count[i] = DB_Column_Limit;
                            dummy += DB_Column_Limit;
                            Dummy_DB_Count++;
                        }


                        if (i == 0)
                        {
                            Per_DB_Column_Count_Start[i] = TheFirst_Trashes_Header_Count + 1;
                            Per_DB_Column_Count_End[i] = DB_Column_Limit + TheFirst_Trashes_Header_Count - 1;
                        }
                        else if (i == Per_DB_Column_Count.Length - 1)
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count;
                            Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i] + TheFirst_Trashes_Header_Count - 1;
                        }
                        else
                        {
                            Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count - DB_Column_Limit;
                            Per_DB_Column_Count_End[i] = dummy + TheFirst_Trashes_Header_Count - 1;
                        }


                    }
                }
            }

            public void Make_New_header()
            {
                New_Header = new string[Reference_Header.Length];
                for (int j = 0; j < Reference_Header.Length; j++)
                {
                    string Dummy = Reference_Header[j].Replace('.', '_');
                    Dummy = Dummy.Replace('-', '_');

                    New_Header[j] = Dummy;
                }

 
            
            }

            public void Edit_Data(string[] GetString, int Data_Row, Data_Class.Data_Editing.INT Data_Int)
            {
                try
                {

                    if (Data_Row != 0)
                    {
                        Data = Data_Int;
                        Data.Getstring = GetString;
                        ThreadFlags = new ManualResetEvent[Data.DB_Count];
                        Wait = new bool[Data.DB_Count];

                        For_Thread_New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        For_Thread_New_Data[0] = Convert.ToDouble(Dummy);

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            ThreadFlags[i] = new ManualResetEvent(false);
                            ThreadPool.QueueUserWorkItem(new WaitCallback(Thread_Edit_Data), i);
                        }

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            Wait[i] = ThreadFlags[i].WaitOne();
                        }

                    }
                    else if (Data_Row == 0)
                    {
                        New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        New_Data[0] = Convert.ToDouble(Dummy);

                        int j = 1;
                        for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                        {
                            New_Data[j] = Convert.ToDouble(GetString[i]);
                            j++;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }

            public void Thread_Edit_Data(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                int j = 0;
                int q = 0;
                if (i == 0)
                {
                    for (int k = 1; k < DB_Column_Limit; k++)
                    {
                        For_Thread_New_Data[k] = Convert.ToDouble(Getstring[TheFirst_Trashes_Header_Count + 1 + q]);
                        q++;
                    }
                }
                else
                {

                    for (int k = 0; k < Per_DB_Column_Count[i]; k++)
                    {
                        For_Thread_New_Data[(DB_Column_Limit * i) + k] = Convert.ToDouble(Getstring[Per_DB_Column_Count_Start[i] + k]);
                        j++;
                    }
                }
                ThreadFlags[i].Set();
            }

            public void Edit_Data(string[] GetString)
            {
                try
                {


                    New_Data = new double[Reference_Header.Length];
                    string Dummy = GetString[0].Replace("PID-", "");
                    New_Data[0] = Convert.ToDouble(Dummy);

                    int j = 1;
                    for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                    {
                        New_Data[j] = Convert.ToDouble(GetString[i]);
                        j++;
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }
            public void Find_Trash_Header(string Value)
            {
                switch (Value.ToUpper().Trim())
                {
                    case "PASSFAIL":
                    case "TIMESTAMP":
                    case "INDEXTIME":
                    case "PARTSN":
                    case "SWBINNAME":
                    case "HWBINNAME":
                        TheEnd_Trashes_Header_Count++;
                        break;
                    case "SBIN":
                    case "HBIN":
                    case "DIE_X":
                    case "DIE_Y":
                    case "SITE":
                    case "TIME":
                    case "TOTAL_TESTS":
                    case "LOT_ID":
                    case "WAFER_ID":
                        TheFirst_Trashes_Header_Count++;
                        break;
                }
            }

            public void TestPlanAddDic(object[,] Data, int Row)
            {

            }

            public void Find_Para_by_Defined(string Data, string Spec_Min, string Spec_Max, string Typical, string Convert, string Complience, int index, int Both)
            {

            }

        }

        public class MERGE_S4PD : INT
        {
            Data_Editing Edit = new Data_Editing();
            public Data_Class.Data_Editing.INT Data { get; set; }
            public ManualResetEvent[] ThreadFlags { get; set; }
            public StringBuilder[] stringA { get; set; }
            public bool[] Wait { get; set; }

            public string[] Getstring { get; set; }
            public List<string> Reference_Header_List { get; set; }
            public string[] Reference_Header { get; set; }

            public string Data_Table { get; set; }

            public double[] New_HighSpec { get; set; }
            public double[] New_LowSpec { get; set; }
            public string[] New_Header { get; set; }
            public double[] New_Data { get; set; }
            public double[] For_Thread_New_Data { get; set; }
            public string[] For_GetSpec_Header { get; set; }
            public string[] Customer_Clotho_Spec_Data { get; set; }
            public string[] Clotho_Spec_Data { get; set; }
            public string Defined_Spec_Min { get; set; }
            public string Defined_Spec_Max { get; set; }
            public string Defined_Spec_Typical { get; set; }
            public int Defined_Convert_Index { get; set; }
            public string Defined_Convert { get; set; }
            public string Defined_Complience { get; set; }
            public int Defined_Both { get; set; }
            public string Spec_Num { get; set; }
            public int TheFirst_Trashes_Header_Count { get; set; }
            public int TheEnd_Trashes_Header_Count { get; set; }

            public int DB_Count { get; set; }
            public int DB_Column_Limit { get; set; }

            public int[] Per_DB_Column_Count { get; set; }
            public int[] Per_DB_Column_Count_Start { get; set; }
            public int[] Per_DB_Column_Count_End { get; set; }

            public bool Flag { get; set; }
            public bool SUBLOT_Falg { get; set; }

            public Dictionary<string, string> Dummy_Spec_Band { get; set; }
            public Dictionary<string, Dictionary<string, string>> Spec_Band { get; set; }

            public Dictionary<string, Spec> Dic_Spec { get; set; }
            public Dictionary<string, SWBIN> SWBIN_Dic { get; set; }

            public List<Clotho_Spec> Clotho_List { get; set; }
            public List<Clotho_Spec> Customor_Clotho_List { get; set; }
            public List<Clotho_Spec> New_Clotho_List { get; set; }
            public List<Clotho_Spec> Clotho_Spcc_List { get; set; }
            public string[] Para { get; set; }
            public string[] Band { get; set; }
            public int ConditionCount { get; set; }

            public List<string> GetTCFDefineSpecNum { get; set; }
            public Dictionary<string, List<string>> GetTCFDefineSpecNum1 { get; set; }
            public Dictionary<string, string> Excel_Combobox { get; set; }
            public int Set_ID { get; set; }
            public int Set_FAIL { get; set; }
            public bool _From_DB { get; set; }

            public string[] Ref_New_Header { get; set; }
            public double[] Ref_New_HighSpec { get; set; }
            public double[] Ref_New_LowSpec { get; set; }
            public string[] Ref_ForAnl_NewMinSpec { get; set; }
            public string[] Ref_ForAnl_NewMaxSpec { get; set; }




            public bool Find_First_Row(string[] Getstring)
            {
                try
                {
                    Flag = false;
                    TheFirst_Trashes_Header_Count = 0;
                    TheEnd_Trashes_Header_Count = 0;
                    if (Getstring[0].ToUpper().Contains("FREQ"))
                    {

                        int j = 1;
                        Reference_Header = new string[Getstring.Length + 1];
                        Reference_Header[0] = Getstring[0];
                        string Name = "";
                        for (int i = 1; i < Getstring.Length + 1; i++)
                        {

                            Name = Getstring[i];
                            Reference_Header[i] = Name + "_dB";
                            i++;

                            Reference_Header[i] = Name + "_Phase";

                            j++;
                        }

                        Ref_New_Header = Reference_Header;
                        Flag = true;
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            public bool Find_Spec_Row(string[] Getstring, bool Flag_NewSpec)
            {
                try
                {
                    Flag = false;
                    if (Getstring[0].ToUpper().Contains("HIGH") || Getstring[0].ToUpper().Contains("LOW"))
                    {
                        int j = 1;
                        if (Getstring[0].ToUpper().Contains("HIGH"))
                        {
                            New_HighSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_HighSpec[0] = Convert.ToDouble(0);
                            int TheFrist = TheFirst_Trashes_Header_Count;
                            int TheEnd = TheEnd_Trashes_Header_Count;
                            if (Flag_NewSpec)
                            {
                                for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                {
                                    New_HighSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMaxSpec[i]);
                                    j++;
                                }
                            }
                            else
                            {
                                try
                                {
                                    for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                    {
                                        // if (Getstring[i] == "") New_HighSpec[j] = 999;
                                        if (Getstring[i] == "") { }
                                        else New_HighSpec[j] = Convert.ToDouble(Getstring[i]);
                                        j++;
                                    }
                                }
                                catch
                                {

                                }
                            }

                            //   Data_Editing.New_HighSpec = New_HighSpec;
                            Flag = true;
                        }
                        if (Getstring[0].ToUpper().Contains("LOW"))
                        {
                            New_LowSpec = new double[Getstring.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count];
                            New_LowSpec[0] = Convert.ToDouble(0);
                            int TheFrist = TheFirst_Trashes_Header_Count;
                            int TheEnd = TheEnd_Trashes_Header_Count;

                            if (Flag_NewSpec)
                            {
                                for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                {
                                    New_LowSpec[j] = Convert.ToDouble(Ref_ForAnl_NewMinSpec[i]);
                                    j++;
                                }
                            }
                            else
                            {
                                for (int i = TheFrist + 1; i < Getstring.Length - TheEnd; i++)
                                {
                                    //if (Getstring[i] == "") New_LowSpec[j] = -999;
                                    if (Getstring[i] == "") { }
                                    else New_LowSpec[j] = Convert.ToDouble(Getstring[i]);
                                    j++;
                                }
                            }

                            //    Data_Editing.New_LowSpec = New_LowSpec;
                            Flag = true;
                        }
                    }
                }
                catch (Exception e)
                {
                    Flag = false;
                    MessageBox.Show(e.ToString());
                }
                return Flag;
            }

            //public void Define_DB_Count(string[] Getstring)
            //{
            //    if (Reference_Header.Length - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count < 2000)
            //    {
            //        DB_Count = 1;
            //        Per_DB_Column_Count = new int[DB_Count];
            //        Per_DB_Column_Count[0] = DB_Column_Limit;
            //        Per_DB_Column_Count_Start[0] = DB_Column_Limit;
            //        Per_DB_Column_Count_End[0] = DB_Column_Limit;

            //    }
            //    else
            //    {
            //        double length = 0f;
            //        for (int j = 0; j < 20; j++)
            //        {
            //            length = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(j);
            //            if (length < 2000)
            //            {
            //                break;
            //            }

            //        }
            //        double Get_Count = Convert.ToDouble(Reference_Header.Length) / Convert.ToDouble(length);
            //        double Temp = Math.Truncate(Get_Count);

            //        int Dummy_DB_Count = 0;

            //        if (Get_Count > Temp) DB_Count = Convert.ToInt16(Temp) + 1;
            //        else DB_Count = Convert.ToInt16(Temp);

            //        Per_DB_Column_Count = new int[DB_Count];
            //        Per_DB_Column_Count_Start = new int[DB_Count];
            //        Per_DB_Column_Count_End = new int[DB_Count];

            //        int dummy = 0;


            //        for (int i = 0; i < Per_DB_Column_Count.Length; i++)
            //        {
            //            if (i == Per_DB_Column_Count.Length - 1)
            //            {
            //                Per_DB_Column_Count[i] = Getstring.Length - (dummy) - TheFirst_Trashes_Header_Count - TheEnd_Trashes_Header_Count;
            //                Dummy_DB_Count++;
            //            }
            //            else
            //            {
            //                Per_DB_Column_Count[i] = Convert.ToInt16(Math.Truncate(length)) + 1;
            //                dummy += Convert.ToInt16(Math.Truncate(length)) + 1;
            //                Dummy_DB_Count++;
            //            }

            //            if (i == 0)
            //            {
            //                Per_DB_Column_Count_Start[i] = TheFirst_Trashes_Header_Count + 1;
            //                Per_DB_Column_Count_End[i] = Convert.ToInt16(Math.Truncate(length)) + 1 + TheFirst_Trashes_Header_Count - 1;
            //            }
            //            else if (i == Per_DB_Column_Count.Length - 1)
            //            {
            //                Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count;
            //                Per_DB_Column_Count_End[i] = dummy + Per_DB_Column_Count[i] + TheFirst_Trashes_Header_Count - 1;
            //            }
            //            else
            //            {
            //                Per_DB_Column_Count_Start[i] = dummy + TheFirst_Trashes_Header_Count - Convert.ToInt16(Math.Truncate(length)) + 1;
            //                Per_DB_Column_Count_End[i] = dummy + TheFirst_Trashes_Header_Count - 1;
            //            }
            //        }

            //        DB_Column_Limit = Per_DB_Column_Count[0];
            //    }
            //}

            public void Find_Cloth_DataFile(string[] Getstring)
            {

            }
            public void Find_Cloth_DataFile_For_New_Spec(string[] Customer)
            {

            }
            public void Define_DB_Count(string[] Getstring)
            {
                DB_Count = 1;

            
            }

            public void Make_New_header()
            {
                New_Header = new string[Reference_Header.Length];
                for (int j = 0; j < Reference_Header.Length; j++)
                {
                    string Dummy = Reference_Header[j].Replace('.', '_');
                    Dummy = Dummy.Replace('-', '_');

                    New_Header[j] = Dummy;
                }



            }

            public void Edit_Data(string[] GetString, int Data_Row, Data_Class.Data_Editing.INT Data_Int)
            {
                try
                {

                    if (Data_Row != 0)
                    {
                        Data = Data_Int;
                        Data.Getstring = GetString;
                        ThreadFlags = new ManualResetEvent[Data.DB_Count];
                        Wait = new bool[Data.DB_Count];

                        For_Thread_New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        For_Thread_New_Data[0] = Convert.ToDouble(Dummy);

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            ThreadFlags[i] = new ManualResetEvent(false);
                            ThreadPool.QueueUserWorkItem(new WaitCallback(Thread_Edit_Data), i);
                        }

                        for (int i = 0; i < Data.DB_Count; i++)
                        {
                            Wait[i] = ThreadFlags[i].WaitOne();
                        }

                    }
                    else if (Data_Row == 0)
                    {
                        New_Data = new double[Reference_Header.Length];
                        string Dummy = GetString[0].Replace("PID-", "");
                        New_Data[0] = Convert.ToDouble(Dummy);

                        int j = 1;
                        for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                        {
                            New_Data[j] = Convert.ToDouble(GetString[i]);
                            j++;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }

            public void Thread_Edit_Data(Object threadContext)
            {
                int i = (int)threadContext;
                int Count = Data.Per_DB_Column_Count[i];

                int j = 0;
                int q = 0;
                if (i == 0)
                {
                    for (int k = 1; k < DB_Column_Limit; k++)
                    {
                        For_Thread_New_Data[k] = Convert.ToDouble(Getstring[TheFirst_Trashes_Header_Count + 1 + q]);
                        q++;
                    }
                }
                else
                {

                    for (int k = 0; k < Per_DB_Column_Count[i]; k++)
                    {
                        For_Thread_New_Data[(DB_Column_Limit * i) + k] = Convert.ToDouble(Getstring[Per_DB_Column_Count_Start[i] + k]);
                        j++;
                    }
                }
                ThreadFlags[i].Set();
            }

            public void Edit_Data(string[] GetString)
            {
                try
                {


                    New_Data = new double[Reference_Header.Length];
                    string Dummy = GetString[0].Replace("PID-", "");
                    New_Data[0] = Convert.ToDouble(Dummy);

                    int j = 1;
                    for (int i = TheFirst_Trashes_Header_Count + 1; i < GetString.Length - TheEnd_Trashes_Header_Count; i++)
                    {
                        New_Data[j] = Convert.ToDouble(GetString[i]);
                        j++;
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }
            public void Find_Trash_Header(string Value)
            {
                switch (Value.ToUpper().Trim())
                {
                    case "PASSFAIL":
                    case "TIMESTAMP":
                    case "INDEXTIME":
                    case "PARTSN":
                    case "SWBINNAME":
                    case "HWBINNAME":
                        TheEnd_Trashes_Header_Count++;
                        break;
                    case "SBIN":
                    case "HBIN":
                    case "DIE_X":
                    case "DIE_Y":
                    case "SITE":
                    case "TIME":
                    case "TOTAL_TESTS":
                    case "LOT_ID":
                    case "WAFER_ID":
                        TheFirst_Trashes_Header_Count++;
                        break;
                }
            }

            public void TestPlanAddDic(object[,] Data, int Row)
            {

            }

            public void Find_Para_by_Defined(string Data, string Spec_Min, string Spec_Max, string Typical, string Convert, string Complience, int index, int Both)
            {

            }

        }
        public interface INT
        {
            string[] Getstring { get; set; }
            Data_Class.Data_Editing.INT Data { get; set; }
            ManualResetEvent[] ThreadFlags { get; set; }
            StringBuilder[] stringA { get; set; }
            bool[] Wait { get; set; }

            List<string> Reference_Header_List { get; set; }
            string[] Reference_Header { get; set; }
            double[] New_HighSpec { get; set; }
            double[] New_LowSpec { get; set; }
            string[] New_Header { get; set; }
            double[] New_Data { get; set; }
            double[] For_Thread_New_Data { get; set; }
            string[] For_GetSpec_Header { get; set; }

            string[] Clotho_Spec_Data { get; set; }
            string[] Customer_Clotho_Spec_Data { get; set; }


            string Defined_Spec_Min { get; set; }
            string Defined_Spec_Max { get; set; }
            string Defined_Spec_Typical { get; set; }
            int Defined_Convert_Index { get; set; }
            string Defined_Convert { get; set; }
            string Defined_Complience { get; set; }
            int Defined_Both { get; set; }
            string Spec_Num { get; set; }

            int TheFirst_Trashes_Header_Count { get; set; }
            int TheEnd_Trashes_Header_Count { get; set; }

            int DB_Count { get; set; }
            int DB_Column_Limit { get; set; }

            int[] Per_DB_Column_Count { get; set; }
            int[] Per_DB_Column_Count_Start { get; set; }
            int[] Per_DB_Column_Count_End { get; set; }

            bool Flag { get; set; }
            string Data_Table { get; set; }

            Dictionary<string, string> Dummy_Spec_Band { get; set; }
            Dictionary<string, Dictionary<string, string>> Spec_Band { get; set; }

            Dictionary<string, Spec> Dic_Spec { get; set; }

            Dictionary<string, SWBIN> SWBIN_Dic { get; set; }
            List<Clotho_Spec> Clotho_List { get; set; }
            List<Clotho_Spec> Customor_Clotho_List { get; set; }
            List<Clotho_Spec> New_Clotho_List { get; set; }
            List<Clotho_Spec> Clotho_Spcc_List { get; set; }



            string[] Para { get; set; }
            string[] Band { get; set; }
            int ConditionCount { get; set; }
            int Set_ID { get; set; }
            int Set_FAIL { get; set; }

            bool _From_DB { get; set; }
            bool SUBLOT_Falg { get; set; }
        
            string[] Ref_New_Header { get; set; }
            double[] Ref_New_HighSpec { get; set; }
            double[] Ref_New_LowSpec { get; set; }
            string[] Ref_ForAnl_NewMinSpec { get; set; }
            string[] Ref_ForAnl_NewMaxSpec { get; set; }

            bool Find_First_Row(string[] Getstring);
            bool Find_Spec_Row(string[] Getstring, bool Flag);

            void Find_Cloth_DataFile(string[] Getstring);
            void Find_Cloth_DataFile_For_New_Spec(string[] Customer);

            void Define_DB_Count(string[] Getstring);
            void Make_New_header();
            void Edit_Data(string[] Getstring, int Data_Row, Data_Class.Data_Editing.INT Data);

            void Edit_Data(string[] Getstring);
            void TestPlanAddDic(object[,] Data, int i);
            void Find_Para_by_Defined(string Data, string Spec_Min, string Spec_Max, string Typical, string Convert, string Complience, int index, int Both);


        }
        public INT Open(string Key)
        {
            INT Int = null;
            switch (Key)
            {
                case "YIELD":
                    Int = new Yield();
                    Int.New_Header = new string[1];
                    Int.DB_Column_Limit = 1993;
                    Int.TheEnd_Trashes_Header_Count = 0;
                    Int.TheFirst_Trashes_Header_Count = 0;

                    break;

                case "BOXPLOT":
                    Int = new BOXPLOT();
                    Int.New_Header = new string[1];
                    Int.DB_Column_Limit = 1993;
                    Int.TheEnd_Trashes_Header_Count = 0;
                    Int.TheFirst_Trashes_Header_Count = 0;

                    break;

                case "FCM":
                    Int = new FCM_Automation_EXCEL();
                    Int.DB_Column_Limit = 1993;

                    break;

                case "GETSPEC":
                    Int = new GETSPEC();
                    Int.DB_Column_Limit = 1993;

                    break;

                case "MERGE":
                    Int = new MERGE();
                    Int.New_Header = new string[1];
                    Int.DB_Column_Limit = 1993;
                    Int.TheEnd_Trashes_Header_Count = 0;
                    Int.TheFirst_Trashes_Header_Count = 0;

                    break;
                case "MERGE_S4PD":
                    Int = new MERGE_S4PD();
                    Int.New_Header = new string[1];
                    Int.DB_Column_Limit = 1993;
                    Int.TheEnd_Trashes_Header_Count = 0;
                    Int.TheFirst_Trashes_Header_Count = 0;

                    break;
            }
            return Int;
        }
        public class Spec
        {
            public string Min;
            public string Max;
            public string Typical;
            public string SpecNumber;
            public string Convert;
            public string Complience;
            public int Index;
            public int Both;
            public Spec(string Min, string Max, string Typical, string SpecNumber, string Convert, string Complience, int Index, int Both)
            {
                this.Min = Min;
                this.Max = Max;
                this.Typical = Typical;
                this.SpecNumber = SpecNumber;
                this.Convert = Convert;
                this.Complience = Complience;
                this.Index = Index;
                this.Both = Both;
            }
        }

        public class SWBIN
        {
            public string Name;
            public string Bin;
            public bool Flag;
            public SWBIN(string Name, string Bin, bool Flag)
            {
                this.Name = Name;
                this.Bin = Bin;
                this.Flag = Flag;
            }
        }

        public class Clotho_Spec
        {
            public double[] Min;
            public double[] Max;

            public Clotho_Spec(double[] Min, double[] Max)
            {
                this.Min = Min;
                this.Max = Max;

            }
        }



    }
}
