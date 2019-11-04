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
using System.IO.Compression;
using System.Threading;
using System.Diagnostics;
using System.Reflection;



namespace TestApplication
{
    public partial class Merge_Form : Form
    {
        string Files_path;
        string[] Files;
        string[] Files_Full_Name;
        string Key = "MERGE";
        Dictionary<string, double[]> Bin;

        CSV_Class.CSV CSV = new CSV_Class.CSV();
        CSV_Class.CSV.INT Csv_Interface;

        Data_Class.Data_Editing Data_Edit = new Data_Class.Data_Editing();
        Data_Class.Data_Editing.INT Data_Interface;

        DB_Class.DB_Editing DB = new DB_Class.DB_Editing();
        DB_Class.DB_Editing.INT DB_Interface;

        JMP_Class.JMP_Editing.INT JMP_Interface;
        JMP_Class.JMP_Editing JMP = new JMP_Class.JMP_Editing();


        Dictionary<string, List<string>> Lot_Information;
        Dictionary<string, Dictionary<string, List<string>>> Matching_Lots;


        Dictionary<string, Dictionary<string, List<string>>> information;
        Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>> Matching_Lots_Test;

        List<string> SubLot;
        TreeNode[] _Node;
        List<string> _Lot_Information_Dummy;

        List<List<string>> Node_List;
        List<string> Node_Find;

        List<string> Edit_data;
        string Node1;
        string Node2;
        string Node3;

        int Find_Scroll = 0;


        Dictionary<string, List<string>> FindNode = new Dictionary<string, List<string>>();
        List<string> FindNode_List = new List<string>();

        BindingSource Databind = new BindingSource();
        DataTable Dt = new DataTable();

        BindingSource Databind2 = new BindingSource();
        DataTable Dt2 = new DataTable();


        List<string> List_Files;

        int table_Count = 0;

        string[] Lot = new string[0];

        ContextMenuStrip m = new ContextMenuStrip();
        int x_Text = 0;
        int y_Text = 0;

        string Selected_Lot = "";

        bool Matching_Lot_Flag;

        FolderBrowserDialog Dialog = new FolderBrowserDialog();
        Dir.Dir_Directory Dir;

        OpenFileDialog Dialog1 = new OpenFileDialog();
        List<string> Selected_data_type;
        List<string> data_type;

        List<FileInfo> Filedata;
        Dictionary<string, List<FileInfo>> Dic_File;

        public Merge_Form()
        {
            InitializeComponent();
            dataGridView1.Visible = false;

            Databind.DataSource = Dt;
            dataGridView1.DataSource = Databind;

            ExtensionMethod.DoubleBuffered(dataGridView1, true);

            dataGridView2.Visible = false;

            Databind2.DataSource = Dt2;
            dataGridView2.DataSource = Databind2;

            ExtensionMethod.DoubleBuffered(dataGridView2, true);

            dataGridView3.Visible = false;

            Selected_data_type = new List<string>();
            Dic_File = new Dictionary<string, List<FileInfo>>();

            listBox2.Hide();
            Node_List = new List<List<string>>();
            Node_Find = new List<string>();
            Edit_data = new List<string>();

            Dir.Dir_Directory Dir = new Dir.Dir_Directory("C:\\temp\\dummy");
        }


        private void button1_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog Dialog = new FolderBrowserDialog();
            Dir.Dir_Directory Dir;

            Dialog.ShowDialog();
            string selected = Dialog.SelectedPath;

            DirectoryInfo di = new DirectoryInfo(selected);

            List<FileInfo> Filedata = DirSeach(selected);

            int data_Count = 0;
            for (int k = 0; k < Filedata.Count; k++)
            {
                int index = Filedata[k].Name.Length;
                string ss = Filedata[k].Name;
                string Dumy = Filedata[k].Name.ToString().ToUpper().Substring(index - 4, 4);

                int a = 0;
                if (Dumy == ".ZIP")
                {
                    data_Count++;
                }

            }



            Progress_Form Progress = new Progress_Form("UNZIP", 0);
            Progress.Show();
            Progress.Merge_Unzip_Init(data_Count);

            int Count = 0;
            foreach (var item in Filedata)
            {

                if (item.FullName.Substring(item.Name.Length - 4, 4).ToUpper() != ".CSV" && item.FullName.Substring(item.Name.Length - 4, 4).ToUpper() != ".TXT")
                {
                    if (item.DirectoryName.ToUpper().ToString().Contains("EXTRACT"))
                    {

                    }
                    else
                    {
                        Dir = new Dir.Dir_Directory(item.DirectoryName + "\\Extract");

                        try
                        {
                            using (ZipArchive archive = ZipFile.OpenRead(item.FullName.ToString()))
                            {
                                foreach (ZipArchiveEntry entry in archive.Entries)
                                {
                                    FileInfo F = new FileInfo(item.DirectoryName + "\\Extract\\" + entry);

                                    if (!F.Exists)
                                    {
                                        entry.ExtractToFile(Path.Combine(item.DirectoryName + "\\Extract\\", entry.FullName));
                                    }

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Cannot open file " + item.FullName + " if you downloaded this file, try downloading the file again");
                        }
                    }
                    Progress.Merge_Unzip_Print(Count + 1);
                    Count++;

                }

            }


            listBox1.Items.Clear();

            Files = null;

            Progress.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog Dialog = new OpenFileDialog();

            if (checkBox1.Checked)
            {
                Dialog.Filter = "DB Files (*.db)| *.db";
                Dialog.InitialDirectory = "C:\\Automation\\DB\\yield";
                Dialog.Multiselect = true;
                Dialog.ShowDialog();
            }
            Csv_Interface = CSV.Open(Key);
            Data_Interface = Data_Edit.Open(Key);
            DB_Interface = DB.Open(Key);

            Progress_Form Progress = new Progress_Form("MERGE", 0);
            Progress.Merge_Init(Files.Length);
            Progress.Show();
            //Sample_Inf = new Dictionary<string, List<string>>[3];

            //for(int sample_inf_index = 0; sample_inf_index < Sample_Inf.Length; sample_inf_index++)
            //{
            //    Sample_Inf[sample_inf_index] = new Dictionary<string, List<string>();
            //}


            long Data_Count = 0;
            long Data_Count_for_Limit = 0;


            Data_Interface.Data_Table = "data0";

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();



            DB_Interface._SUBLOT_Flag = false;

            for (int i = 0; i < Files.Length; i++)
            {
                Csv_Interface.Read_Open(this.Files_Full_Name[i]);
                bool Flag = false;

                #region

                if (i == 0)
                {
                    while (!Csv_Interface.StreamReader.EndOfStream)
                    {
                        Csv_Interface.Read();

                        if (Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1].Contains("SUBLOT"))
                        {
                            List<string> Array = Csv_Interface.Get_String.ToList();

                            Array.RemoveAt(Csv_Interface.Get_String.Length - 1);

                            Csv_Interface.Get_String = Array.ToArray();

                            DB_Interface._SUBLOT_Flag = true;
                        }

                        Flag = Data_Interface.Find_First_Row(Csv_Interface.Get_String);


                        if (Flag) break;
                        if (Csv_Interface.Get_String[0].ToUpper() == "LOT")
                        {
                            DB_Interface.Lot_ID = Csv_Interface.Get_String[1];

                        }
                        else if (Csv_Interface.Get_String[0].ToUpper() == "SUBLOT")
                        {
                            DB_Interface.SubLot_ID = Csv_Interface.Get_String[1];
                        }
                        else if (Csv_Interface.Get_String[0].ToUpper() == "HOSTIPADDRESS")
                        {
                            DB_Interface.Tester_ID = Csv_Interface.Get_String[1];
                        }

                    }

                    Data_Interface.Define_DB_Count(Csv_Interface.Get_String);

                    for (int l = 0; l < Csv_Interface.Get_String.Length; l++)
                    {
                        if (Csv_Interface.Get_String[l].ToUpper() == "SBIN")
                        {
                            DB_Interface.Bin_place = 1;

                        }

                    }
                    if (checkBox1.Checked)
                    {

                        DB_Interface.Open_DB(Dialog.FileNames, Data_Interface);
                        DB_Interface.trans(Data_Interface);

                        for (int k = 0; k < 10; k++)
                        {
                            string Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

                            DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(0, Query);
                        }

                        string[] data = new string[0];

                        for (int k = 0; k < DB_Interface.Table_Count; k++)
                        {
                            string Query = "Select DISTINCT id from data" + k;
                            data = DB_Interface.Get_Data_By_Query(Query);

                            //   Lot = Lot.Concat(data).ToArray();

                        }

                        Data_Count = Convert.ToInt64(data[data.Length - 1]);
                        Data_Count++;

                        for (int k = 0; k < DB_Interface.Table_Count; k++)
                        {
                            string Query = "select count(HBIN) from data" + k;

                            Data_Count += DB_Interface.Get_Sample_Count(0, Query);
                        }
                    }
                    else
                    {
                        Dir.Dir_Directory Dir = new Dir.Dir_Directory("C:\\Automation\\DB\\YIELD\\" + Files[0]);

                        Data_Interface.Make_New_header();

                        DB_Interface.Open_DB(Files[0], Data_Interface);
                        DB_Interface.DropTable(Data_Interface, "");

                        DB_Interface.trans(Data_Interface);
                    }


                    Flag = false;

                    while (!Csv_Interface.StreamReader.EndOfStream)
                    {
                        Csv_Interface.Read();
                        if (Csv_Interface.Get_String[0].Contains("PID"))
                        {
                            Flag = true;
                        }
                        else if (Csv_Interface.Get_String[0].Contains("HighL"))
                        {
                            Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                        }
                        else if (Csv_Interface.Get_String[0].Contains("LowL"))
                        {
                            Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                        }
                        if (Flag) break;
                    }
                    if (checkBox1.Checked)
                    {

                    }
                    else
                    {
                        DB_Interface.Clotho_Spec_Flag = true;
                        DB_Interface.Insert_Header(Data_Interface);
                        DB_Interface.Insert_Spec_Data("Clotho_Spec");
                        DB_Interface.Make_table("INF");
                        DB_Interface.Make_table2(Data_Interface, "REFHEADER");

                        DB_Interface.Insert_Ref_Header_Data(Data_Interface);


                        Data_Interface.Data_Table = "data0";
                        DB_Interface.Clotho_Spec_Flag = false;
                    }

                }
                else
                {
                    Flag = false;
                    while (!Csv_Interface.StreamReader.EndOfStream)
                    {
                        Csv_Interface.Read();


                        if (Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1].Contains("SUBLOT"))
                        {
                            List<string> Array = Csv_Interface.Get_String.ToList();

                            Array.RemoveAt(Csv_Interface.Get_String.Length - 1);

                            Csv_Interface.Get_String = Array.ToArray();
                            //  Flag = true;
                            DB_Interface._SUBLOT_Flag = true;
                        }

                        // /i/*f (Flag) break;*/

                        if (Csv_Interface.Get_String[0].ToUpper() == "LOT")
                        {
                            DB_Interface.Lot_ID = Csv_Interface.Get_String[1];
                        }
                        else if (Csv_Interface.Get_String[0].ToUpper() == "SUBLOT")
                        {
                            DB_Interface.SubLot_ID = Csv_Interface.Get_String[1];
                        }
                        else if (Csv_Interface.Get_String[0].ToUpper() == "HOSTIPADDRESS")
                        {
                            DB_Interface.Tester_ID = Csv_Interface.Get_String[1];
                        }
                        if (Csv_Interface.Get_String[0].Contains("PID"))
                        {
                            Flag = true;
                        }
                        else if (Csv_Interface.Get_String[0].Contains("HighL"))
                        {
                            Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                        }
                        else if (Csv_Interface.Get_String[0].Contains("LowL"))
                        {
                            Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                            DB_Interface.Delete_Spec_Data("Clotho_Spec");
                            DB_Interface.Insert_Spec_Data("Clotho_Spec");
                        }
                        if (Flag) break;
                    }
                }

                #endregion

                DB_Interface.Insert_ThreadFlags = new ManualResetEvent[2];
                DB_Interface.Insert_Thread_Wait = new bool[2];

                DB_Interface.TheFirst_Trashes_Header_Count = Data_Interface.TheFirst_Trashes_Header_Count;
                DB_Interface.TheEnd_Trashes_Header_Count = Data_Interface.TheEnd_Trashes_Header_Count;

                DB_Interface.TheFirst_Trashes_Header_Count = 0;
                DB_Interface.TheEnd_Trashes_Header_Count = 0;

                Data_Interface.Getstring = Csv_Interface.Get_String;

                DB_Interface.Bin = Csv_Interface.Get_String[DB_Interface.Bin_place];

                if (DB_Interface._SUBLOT_Flag)
                {

                    DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                    DB_Interface.SubLot_ID = Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1];
                    DB_Interface.Tester_ID = Csv_Interface.Get_String[5];
                    DB_Interface.Site = Csv_Interface.Get_String[5];

                }
                else
                {
                    DB_Interface.Site = DB_Interface.Tester_ID;
                }

                Data_Count++;
                Data_Count_for_Limit++;

                for (int thread_i = 0; thread_i < 2; thread_i++)
                {
                    DB_Interface.Insert_ThreadFlags[thread_i] = new ManualResetEvent(false);
                }
                ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));


                DB_Interface.Insert_ThreadFlags[1].Set();

                DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

                string[] GetData = Csv_Interface.Read();

                Data_Interface.Getstring = GetData;

                DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

                Data_Count++;
                Data_Count_for_Limit++;


                while (!Csv_Interface.StreamReader.EndOfStream)
                {
                    if (DB_Interface.Limit < Data_Count_for_Limit)
                    {

                        DB_Interface.Commit(Data_Interface);

                        table_Count++;
                        Data_Interface.Data_Table = "data" + table_Count;
                        Data_Count_for_Limit = 0;
                        DB_Interface.Insert_Header(Data_Interface);

                        Data_Interface.Data_Table = "data" + table_Count;
                        DB_Interface.trans(Data_Interface);

                    }

                    for (int thread_i = 0; thread_i < 2; thread_i++)
                    {
                        DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                    }
                    ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

                    GetData = Csv_Interface.Read_Test();

                    DB_Interface.Insert_ThreadFlags[1].Set();

                    DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                    DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

                    Data_Interface.Getstring = GetData;

                    if (DB_Interface._SUBLOT_Flag)
                    {

                        DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                        DB_Interface.SubLot_ID = Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1];
                        DB_Interface.Tester_ID = Csv_Interface.Get_String[5];
                        DB_Interface.Site = Csv_Interface.Get_String[5];

                    }

                    DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

                    Progress.Merge_Print(Data_Count, i);

                    Data_Count++;
                    Data_Count_for_Limit++;
                }

                for (int thread_i = 0; thread_i < 2; thread_i++)
                {
                    DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                }
                ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

                DB_Interface.Insert_ThreadFlags[1].Set();

                DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();


                Csv_Interface.Read_Close();

            }

            if (table_Count == 0) table_Count = 1;
            Progress.Close();
            DB_Interface.Commit(Data_Interface);

            Review();
            Gridview(true);
            Gridview2(true);
            Listbox2_Define();
            listBox2.Show();

            double Testime = TestTime1.Elapsed.TotalMilliseconds;
        }

        private void button6_Click(object sender, EventArgs e)
        {

            Dialog = new FolderBrowserDialog();
            Dialog1 = new OpenFileDialog();

            if (checkBox1.Checked)
            {
                Dialog1.Filter = "DB Files (*.db)| *.db";
                Dialog1.InitialDirectory = "C:\\Automation\\DB\\yield";
                Dialog1.Multiselect = true;
                Dialog1.ShowDialog();
            }

            Dialog.ShowDialog();

            if (Dialog.SelectedPath != "")
            {

                string selected = Dialog.SelectedPath;

                Csv_Interface = CSV.Open(Key);
                Data_Interface = Data_Edit.Open(Key);
                DB_Interface = DB.Open(Key);

                DirectoryInfo di = new DirectoryInfo(selected);

                List<FileInfo> Filedata = DirSeach(selected);

                int data_Count = 0;
                for (int k = 0; k < Filedata.Count; k++)
                {
                    int index = Filedata[k].Name.Length;
                    string ss = Filedata[k].Name;
                    string Dumy = Filedata[k].Name.ToString().ToUpper().Substring(index - 4, 4);

                    int a = 0;
                    if (Dumy == ".CSV")
                    {
                        data_Count++;
                    }

                }

                Gridview3(true, Filedata);


            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            DB_Interface.Close(Data_Interface);
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // DB_Interface.Lot_ID = Lot[i];

            Export();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog Dialog = new OpenFileDialog();


            Dialog.Filter = "DB Files (*.db)| *.db";
            Dialog.InitialDirectory = "C:\\Automation\\DB\\yield";
            Dialog.Multiselect = true;
            Dialog.ShowDialog();


            Csv_Interface = CSV.Open(Key);
            Data_Interface = Data_Edit.Open(Key);
            DB_Interface = DB.Open(Key);

            if (Dialog.FileName != "")
            {
                string Filename = Dialog.FileName.Substring(Dialog.FileName.LastIndexOf("\\") + 1);

                Data_Interface.DB_Count = Dialog.FileNames.Length;

                DB_Interface.Open_DB(Filename.Substring(0, Filename.Length - 5) + ".csv", Data_Interface);

             //   DataCheck_S4pd();


                Review();
                Gridview(true);
                // DB_Interface.Commit(Data_Interface);

                int Count = DB_Interface.Get_Column_Count(Data_Interface, "select COLUMNCOUNT from INF");

                Data_Interface.Per_DB_Column_Count = new int[Data_Interface.DB_Count];

                for (int i = 0; i < Data_Interface.DB_Count - 1; i++)
                {
                    Data_Interface.Per_DB_Column_Count[i] = 1993;
                }

                Data_Interface.Per_DB_Column_Count[Data_Interface.DB_Count - 1] = Count;


                DB_Interface.Get_From_Db_Ref_Header(Data_Interface);

                Files = new string[1];
                Files[0] = Filename.Substring(0, Filename.Length - 5) + ".csv";

                Listbox2_Define();
                listBox2.Show();

                Gridview2(true);
            }




        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (Selected_data_type.Count != 0)
            {
                checkBox2.Checked = true;

                if(checkBox2.Checked)
                {
                    Merge_S4pd();
                }
                string selected = Dialog.SelectedPath;

                Csv_Interface = CSV.Open(Key);
                Data_Interface = Data_Edit.Open(Key);
                DB_Interface = DB.Open(Key);

                DirectoryInfo di = new DirectoryInfo(selected);

                Filedata = DirSeach(selected);

                int data_Count = 0;
                for (int k = 0; k < Filedata.Count; k++)
                {
                    int index = Filedata[k].Name.Length;
                    string ss = Filedata[k].Name;
                    string Dumy = Filedata[k].Name.ToString().ToUpper().Substring(index - 4, 4);

                    int a = 0;
                    if (Dumy == ".CSV")
                    {
                        data_Count++;
                    }

                }


                Make_up_File_List();

                long Data_Count = 0;
                Data_Count = 0;

                foreach (KeyValuePair<string, List<FileInfo>> D in Dic_File)
                {

                    Progress_Form Progress = new Progress_Form("MERGE", 0);
                    Progress.Merge_Init(D.Value.Count);
                    Progress.Show();

           
                    long Data_Count_for_Limit = 0;
                    int Count = 0;

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    DB_Interface._SUBLOT_Flag = false;

                    foreach (FileInfo L in D.Value)
                    {

                   

                        Csv_Interface.Read_Open(L.FullName);
                        bool Flag = false;

                        #region

                        if (Count == 0)
                        {
                            while (!Csv_Interface.StreamReader.EndOfStream)
                            {
                                Csv_Interface.Read();

                                if (Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1].ToUpper().Contains("SUBLOT"))
                                {
                                    List<string> Array = Csv_Interface.Get_String.ToList();

                                    Array.RemoveAt(Csv_Interface.Get_String.Length - 1);

                                    Csv_Interface.Get_String = Array.ToArray();

                                    DB_Interface._SUBLOT_Flag = true;
                                }

                                Flag = Data_Interface.Find_First_Row(Csv_Interface.Get_String);


                                if (Flag) break;

                                if (Csv_Interface.Get_String[0].ToUpper() == "LOT")
                                {
                                    DB_Interface.Lot_ID = Csv_Interface.Get_String[1];
                                    DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "SUBLOT")
                                {
                                    DB_Interface.SubLot_ID = Csv_Interface.Get_String[1];
                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "HOSTIPADDRESS")
                                {
                                    DB_Interface.Tester_ID = Csv_Interface.Get_String[1];
                                }

                            }

                            Data_Interface.Define_DB_Count(Csv_Interface.Get_String);

                            for (int l = 0; l < Csv_Interface.Get_String.Length; l++)
                            {
                                if (Csv_Interface.Get_String[l].ToUpper() == "SBIN")
                                {
                                    DB_Interface.Bin_place = 1;

                                }

                            }


                            if (checkBox1.Checked)
                            {


                                DB_Interface.Open_DB(Dialog1.FileNames, Data_Interface);
                                DB_Interface.trans(Data_Interface);


                                string[] Filedatas = new string[0];
                                string Query = "";


                                for (int k = 0; k < 1; k++)
                                {

                                    Query = "Select File from Files where File = '" + L.Name + "'";
                                    Filedatas = DB_Interface.Get_Data_By_Query(Query);

                                }

                                if (Filedatas.Length != 0)
                                {
                                    Csv_Interface.Read_Close();
                                    goto Next;
                                }




                                Data_Interface.Make_New_header();

                            }
                            else
                            {
                                Dir = new Dir.Dir_Directory("C:\\Automation\\DB\\YIELD\\" + L.Name);

                                Data_Interface.Make_New_header();

                                DB_Interface.Open_DB(L.Name, Data_Interface);
                                DB_Interface.DropTable(Data_Interface, "");

                                DB_Interface.trans(Data_Interface);
                            }


                            Flag = false;

                            while (!Csv_Interface.StreamReader.EndOfStream)
                            {
                                Csv_Interface.Read();
                                if (Csv_Interface.Get_String[0].Contains("PID"))
                                {
                                    Flag = true;
                                }
                                else if (Csv_Interface.Get_String[0].Contains("HighL"))
                                {
                                    Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                                }
                                else if (Csv_Interface.Get_String[0].Contains("LowL"))
                                {
                                    Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                                }
                                if (Flag) break;
                            }
                            if (checkBox1.Checked)
                            {
                                string[] Table = new string[0];
                          
                                DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');
                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Lot_ID + "'";

                                    Table = DB_Interface.Get_Data_By_Query(Query);
                                }

                                if (Table.Length == 0)
                                {
                                    DB_Interface.Clotho_Spec_Flag = false;
                                    DB_Interface.Insert_Header(Data_Interface);
                                }

                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "Select DISTINCT id from " + DB_Interface.Lot_ID;
                                    string[] data = DB_Interface.Get_Data_By_Query(Query);
                                    Data_Count = data.Length;
                                    Data_Count++;
                                }



                            }
                            else
                            {
                                DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                                DB_Interface.Clotho_Spec_Flag = true;
                                DB_Interface.Insert_Header(Data_Interface);
                                DB_Interface.Insert_Spec_Data("Clotho_Spec");
                                DB_Interface.Make_table("INF");
                                DB_Interface.Make_table2(Data_Interface, "REFHEADER");

                                DB_Interface.Make_table_For_Filename(Data_Interface, "Files");


                                DB_Interface.Insert_Ref_Header_Data(Data_Interface);


                                Data_Interface.Data_Table = "data0";
                                DB_Interface.Clotho_Spec_Flag = false;
                            }

                        }
                        else
                        {
                            string[] Filedatas = new string[0];

                            for (int k = 0; k < 1; k++)
                            {

                                string Query = "Select File from Files where File = '" + L.Name + "'";
                                Filedatas = DB_Interface.Get_Data_By_Query(Query);

                            }

                            if (Filedatas.Length != 0)
                            {
                                Csv_Interface.Read_Close();
                                goto Next;
                            }
                            else
                            {


                            }

                            DB_Interface.trans(Data_Interface);

                            Flag = false;
                            while (!Csv_Interface.StreamReader.EndOfStream)
                            {
                                Csv_Interface.Read();


                                if (Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1].Contains("SUBLOT"))
                                {
                                    List<string> Array = Csv_Interface.Get_String.ToList();

                                    Array.RemoveAt(Csv_Interface.Get_String.Length - 1);

                                    Csv_Interface.Get_String = Array.ToArray();
                                    //  Flag = true;
                                    DB_Interface._SUBLOT_Flag = true;
                                }

                                // /i/*f (Flag) break;*/

                                if (Csv_Interface.Get_String[0].ToUpper() == "LOT")
                                {
                                    DB_Interface.Lot_ID = Csv_Interface.Get_String[1];
                                    DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');
                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "SUBLOT")
                                {
                                    DB_Interface.SubLot_ID = Csv_Interface.Get_String[1];
                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "HOSTIPADDRESS")
                                {
                                    DB_Interface.Tester_ID = Csv_Interface.Get_String[1];
                                }
                                if (Csv_Interface.Get_String[0].Contains("PID"))
                                {
                                    Flag = true;
                                }
                                else if (Csv_Interface.Get_String[0].Contains("HighL"))
                                {
                                    Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                                }
                                else if (Csv_Interface.Get_String[0].Contains("LowL"))
                                {
                                    Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                                    DB_Interface.Delete_Spec_Data("Clotho_Spec");
                                    DB_Interface.Insert_Spec_Data("Clotho_Spec");
                                }
                                if (Flag) break;
                            }
                            DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');


                            string[] Table = new string[0];

                            for (int k = 0; k < 1; k++)
                            {
                                string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Lot_ID + "'";

                                Table = DB_Interface.Get_Data_By_Query(Query);
                            }

                            if (Table.Length == 0)
                            {
                                DB_Interface.Clotho_Spec_Flag = false;
                                DB_Interface.Insert_Header(Data_Interface);
                            }


                            DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                            for (int k = 0; k < 1; k++)
                            {
                                string Query = "Select Parameter from " + DB_Interface.Lot_ID;
                                string[] data = DB_Interface.Get_Data_By_Query(Query);
                              //  Data_Count = data.Length;
                                Data_Count++;
                            }



                        }

                        #endregion

                        DB_Interface.Insert_ThreadFlags = new ManualResetEvent[2];
                        DB_Interface.Insert_Thread_Wait = new bool[2];

                        DB_Interface.TheFirst_Trashes_Header_Count = Data_Interface.TheFirst_Trashes_Header_Count;
                        DB_Interface.TheEnd_Trashes_Header_Count = Data_Interface.TheEnd_Trashes_Header_Count;

                        DB_Interface.TheFirst_Trashes_Header_Count = 0;
                        DB_Interface.TheEnd_Trashes_Header_Count = 0;

                        Data_Interface.Getstring = Csv_Interface.Get_String;

                        DB_Interface.Bin = Csv_Interface.Get_String[DB_Interface.Bin_place];

                        Data_Interface.Getstring[8] = Data_Interface.Getstring[8].Replace('-', '_');

                        if (DB_Interface._SUBLOT_Flag)
                        {
                            DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                            DB_Interface.SubLot_ID = Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1];
                            DB_Interface.Tester_ID = Csv_Interface.Get_String[5];
                            DB_Interface.Site = Csv_Interface.Get_String[5];
                            DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                        }
                        else
                        {
                            DB_Interface.Site = DB_Interface.Tester_ID;
                        }

                        Data_Count++;
                        Data_Count_for_Limit++;

                        for (int thread_i = 0; thread_i < 2; thread_i++)
                        {
                            DB_Interface.Insert_ThreadFlags[thread_i] = new ManualResetEvent(false);
                        }
                        ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));


                        DB_Interface.Insert_ThreadFlags[1].Set();

                        DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                        DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

                        string[] GetData = Csv_Interface.Read();

                        Data_Interface.Getstring = GetData;
                        Csv_Interface.Get_String[8] = Csv_Interface.Get_String[8].Replace('-', '_');
                        DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

                        Data_Count++;
                        Data_Count_for_Limit++;


                        while (!Csv_Interface.StreamReader.EndOfStream)
                        {
                            if (DB_Interface.Lot_ID != Csv_Interface.Get_String[8])
                            {
                                DB_Interface.Lot_ID = Csv_Interface.Get_String[8].Replace('-', '_');

                                string[] Table = new string[0];

                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Lot_ID + "'";

                                    Table = DB_Interface.Get_Data_By_Query(Query);
                                }

                                if (Table.Length == 0)
                                {
                                    DB_Interface.Clotho_Spec_Flag = false;
                                    DB_Interface.Insert_Header(Data_Interface);
                                }



                            }

                            for (int thread_i = 0; thread_i < 2; thread_i++)
                            {
                                DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                            }
                            ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

                            GetData = Csv_Interface.Read_Test();
                            Csv_Interface.Get_String[8] = Csv_Interface.Get_String[8].Replace('-', '_');

                            if (DB_Interface.Lot_ID != Csv_Interface.Get_String[8])
                            {
                                  DB_Interface.Lot_ID = Csv_Interface.Get_String[8].Replace('-', '_');

                                string[] Table = new string[0];

                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Lot_ID + "'";

                                    Table = DB_Interface.Get_Data_By_Query(Query);
                                }

                                if (Table.Length == 0)
                                {
                                    DB_Interface.Clotho_Spec_Flag = false;
                                    DB_Interface.Insert_Header(Data_Interface);
                                }


                                //   DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                            }

                            DB_Interface.Insert_ThreadFlags[1].Set();

                            DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                            DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

                            Data_Interface.Getstring = GetData;

                            Csv_Interface.Get_String[8] = Csv_Interface.Get_String[8].Replace('-', '_');
                            if (DB_Interface._SUBLOT_Flag)
                            {

                                DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                                DB_Interface.SubLot_ID = Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1];
                                DB_Interface.Tester_ID = Csv_Interface.Get_String[5];
                                DB_Interface.Site = Csv_Interface.Get_String[5];

                      

                            }

                            DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

                            Progress.Merge_Print(Data_Count, Count);

                            Data_Count++;
                            Data_Count_for_Limit++;
                        }

                        for (int thread_i = 0; thread_i < 2; thread_i++)
                        {
                            DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                        }
                        ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

                        DB_Interface.Insert_ThreadFlags[1].Set();

                        DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                        DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();


                        Csv_Interface.Read_Close();

                        DB_Interface.Insert_Files_Name(L.Name);

                        DB_Interface.Commit(Data_Interface);

                        Next:

                        Count++;
                    }
                    Progress.Close();
                }


                Review();
                Gridview(true);
                Gridview2(true);
                Listbox2_Define();
                listBox2.Show();

            }
            else
            {
                MessageBox.Show("Please Select Data Type.");
            }
            Selected_data_type = null;
        }

        public void Export()
        {
            JMP_Class.Script Distribution_Script;


            Csv_Interface = CSV.Open(Key);
            JMP.Open("MERGE");
            JMP_Interface = JMP.Open("MERGE");

            string Query = "";
            //DB_Interface.Table_Count = 0;

            //for (int k = 0; k < 1; k++)
            //{
            //    // string Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

            //    Query = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY Name";

            //    DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(Data_Interface, Query);
            //    DB_Interface.Table_Count = DB_Interface.Table_Count - 4;
            //}

            ///   if (DB_Interface.Table_Count == 0) DB_Interface.Table_Count = 1;
       

            //if (!DB_Interface._Flag)
            //{
                Matching_Lot_data();
                DB_Interface.Matching_Lots = Matching_Lots;
         //   }
      



            DB_Interface.Get_Saved_Spec(Data_Interface);
            DB_Interface.Get_From_Db_Data_for_Anly(Data_Interface);

            JMP_Interface.Open_Session(true);

            string File = "";

            if (DB_Interface._Flag)
            {
                File = DB_Interface.Filename;
            }
            else
            {
                File = Files[0].Substring(0, Files[0].Length - 4);
            }
         
            for (int i = 0; i < Data_Interface.DB_Count; i++)
            {
                if (i == 0)
                {
                    JMP_Interface.Open_Document("C:\\Automation\\DB\\YIELD\\" + Files[0] + "\\" + File + "_" + i + ".csv");
                    JMP_Interface.GetDataTable();
                }

                else
                {
                    JMP_Interface.Open_Document2("C:\\Automation\\DB\\YIELD\\" + Files[0] + "\\" + File + "_" + i + ".csv");
                    JMP_Interface.GetDataTable2();
                }


                if (i == 1)
                {
                    JMP_Interface.Join("Join" + i);
                    JMP_Interface.Close_Dt(File + "_" + (i - 1));
                    JMP_Interface.Close_Dt2(File + "_" + (i));
                }
                else if (i > 1)
                {
                    JMP_Interface.Join2("Join", i);
                    JMP_Interface.Close_Dt2(File + "_" + (i));

                    JMP_Interface.Close_JoinDT("Join" + (i - 1));
                }

            }

            //  JMP_Interface.Close_JoinDT("Join" + Convert.ToString(Data_Interface.DB_Count - 1));

            Distribution_Script = null;


            // Distribution_Script = JMP_Interface.Make_Script("SAVE", SaveData, FilePaht);
            string[] Array = Files[0].Split('_');
            string New_Filename = Array[0] + "_" + Array[1] + "_" + DB_Interface.Lot_ID;

            Distribution_Script = JMP_Interface.Make_Script_Save("SAVE", "C:\\Automation\\DB\\YIELD\\" + Files[0] + "\\" + File + "_Merge");
            Csv_Interface.Write_Open("C:\\temp\\dummy\\SAVE.jsl");

            Csv_Interface.WriteScript(Distribution_Script.Scrip_Data);
            Csv_Interface.Write_Close();

            JMP_Interface.GetDataTable();
            JMP_Interface.Run_Script("C:\\temp\\dummy\\SAVE.jsl");

            for (int i = 0; i < Data_Interface.DB_Count; i++)
            {
                FileInfo fileDel = new FileInfo(@"C:\\Automation\\DB\\YIELD\\" + Files[0] + "\\" + File + "_" + i + ".csv");
                if (fileDel.Exists) // 삭제할 파일이 있는지
                {
                    fileDel.Delete(); // 없어도 에러안남
                }

            }


        }

        public void m_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string Query = "";
            switch (e.ClickedItem.Text)
            {
                case "Delete":
                    AddDic_Tree_Info();



                    DB_Interface.trans(Data_Interface);

                    for (int k = 0; k < DB_Interface.Table_Count; k++)
                    {
                        Query = "Delete from data" + k + " where LOTID = '" + Node1 + "' and  SUBLOT = '" + Node2 + "' and BIN = '" + Node3 + "'";
                        DB_Interface.Delete_Lot_Data(Query);
                    }

                    DB_Interface.Commit(Data_Interface);

                    this.FindNode = new Dictionary<string, List<string>>();
                    Gridview(true);
                    Review();

                    foreach (KeyValuePair<string, List<string>> key in this.FindNode)
                    {
                        treeView1.Nodes[Convert.ToInt16(key.Value[0])].Nodes[Convert.ToInt16(key.Value[1])].Nodes[0].Nodes[Convert.ToInt16(key.Value[2])].ForeColor = Color.Red;
                    }

                    break;
                case "Hide":
                    AddDic_Tree_Info();

                    DB_Interface.trans(Data_Interface);

                    for (int k = 0; k < DB_Interface.Table_Count; k++)
                    {
                        Query = "Update data" + k + " SET FAIL = '1' where LOTID = '" + Node1 + "' and  SUBLOT = '" + Node2 + "' and BIN = '" + Node3 + "'";
                        DB_Interface.Delete_Lot_Data(Query);

                    }

                    DB_Interface.Commit(Data_Interface);

                    this.FindNode = new Dictionary<string, List<string>>();
                    Gridview(true);
                    Review();

                    foreach (KeyValuePair<string, List<string>> key in this.FindNode)
                    {
                        treeView1.Nodes[Convert.ToInt16(key.Value[0])].Nodes[Convert.ToInt16(key.Value[1])].Nodes[0].Nodes[Convert.ToInt16(key.Value[2])].ForeColor = Color.Blue;
                    }


                    break;
                case "UnHide":
                    AddDic_Tree_Info();

                    DB_Interface.trans(Data_Interface);

                    for (int k = 0; k < DB_Interface.Table_Count; k++)
                    {
                        Query = "Update data" + k + " SET FAIL = '0' where LOTID = '" + Node1 + "' and  SUBLOT = '" + Node2 + "' and BIN = '" + Node3 + "'";
                        DB_Interface.Delete_Lot_Data(Query);
                    }

                    DB_Interface.Commit(Data_Interface);

                    this.FindNode = new Dictionary<string, List<string>>();
                    Gridview(true);
                    Review();

                    foreach (KeyValuePair<string, List<string>> key in this.FindNode)
                    {
                        treeView1.Nodes[Convert.ToInt16(key.Value[0])].Nodes[Convert.ToInt16(key.Value[1])].Nodes[0].Nodes[Convert.ToInt16(key.Value[2])].ForeColor = Color.Black;
                    }


                    break;

                case "Export CSV File":

                    int length = treeView1.SelectedNode.Text.Length;

                    int index = treeView1.SelectedNode.Text.IndexOf("(");


                    Matching_Lot_data();


                    //DB_Interface.Lot_ID = Matching_Lots[treeView1.SelectedNode.Text.Substring(0, index - 1)];

                    //Dictionary<string, string> t = new Dictionary<string, string>();

                    //foreach (KeyValuePair<string, string> d in Matching_Lots)
                    //{
                    //    if (d.Key.ToString() == treeView1.SelectedNode.Text.Substring(0, index - 1))
                    //    {
                    //        t.Add(d.Key, d.Value);
                    //    }

                    //}

                   // DB_Interface.Matching_Lot = t;

                    DB_Interface.trans(Data_Interface);

                    DB_Interface._Flag = true;

                    Export();

                    DB_Interface.Commit(Data_Interface);
                    foreach (KeyValuePair<string, List<string>> key in this.FindNode)
                    {
                        treeView1.Nodes[Convert.ToInt16(key.Value[0])].Nodes[Convert.ToInt16(key.Value[1])].Nodes[0].Nodes[Convert.ToInt16(key.Value[2])].ForeColor = Color.Black;
                    }

                    this.FindNode = new Dictionary<string, List<string>>();
                    //   Gridview(false)

                    DB_Interface._Flag = false;
                    break;

                case "Change Lot ID":

                    string Value = "";


                    DialogResult test = InputBox("", Selected_Lot, ref Value);

                    if (test != DialogResult.Cancel)
                    {
                        DB_Interface.trans(Data_Interface);
                        DB_Interface.LOTID_Update(Selected_Lot, Value, "");
                        DB_Interface.Commit(Data_Interface);
                        Gridview(true);
                        Review();
                        Listbox2_Define();
                        listBox2.Show();
                    }

                    break;
            }

        }

        public void m2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

            switch (e.ClickedItem.Text)
            {
                case "Change Lot ID":




                    break;

                case "Export CSV File":

                    //  DB_Interface.Lot_ID = treeView1.SelectedNode.Text.Substring(0, index - 1);



                    Matching_Lot_data();


                  //  DB_Interface.Lot_ID = Matching_Lots[treeView1.SelectedNode.Text.Substring(0, index - 1)];

                    Dictionary<string, string> t = new Dictionary<string, string>();



                    DB_Interface.Matching_Lots = Matching_Lots;

                 

                    DB_Interface._Flag = true;


                    int i = 0;
                    foreach (Object selecteditem in listBox2.SelectedItems)
                    {
                        DB_Interface.trans(Data_Interface);
                        Dictionary<string, string> test = new Dictionary<string, string>();


                        string d = "";
                        foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                        {
                            Dictionary<string, List<string>> tests = key.Value;


                            foreach (KeyValuePair<string, List<string>> ts in tests)
                            {
                                if (ts.Key == selecteditem.ToString())
                                {
                                    d = key.Key;
                                    DB_Interface.Lot_ID = selecteditem.ToString();
                                    DB_Interface.Matching_Lot = tests;
                                    break;
                                }
                            }

                        }

                       


                        //  test = Matching_Lots[selecteditem.ToString()]
                        //    test.Add(selecteditem.ToString(), Matching_Lots[selecteditem.ToString()]);



               

                        Export();

                        i++;
                        DB_Interface.Commit(Data_Interface);
                    }

                    DB_Interface._Flag = false;

                
                    foreach (KeyValuePair<string, List<string>> key in this.FindNode)
                    {
                        treeView1.Nodes[Convert.ToInt16(key.Value[0])].Nodes[Convert.ToInt16(key.Value[1])].Nodes[0].Nodes[Convert.ToInt16(key.Value[2])].ForeColor = Color.Black;
                    }

                    this.FindNode = new Dictionary<string, List<string>>();
                    //   Gridview(false);
                    break;

            }

        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                m = new ContextMenuStrip();

                if (e.Button == MouseButtons.Right)
                {
                    if (treeView1.SelectedNode.Level == 0)
                    {

                        int index = treeView1.SelectedNode.FullPath.IndexOf('(');

                        Selected_Lot = treeView1.SelectedNode.FullPath.Substring(0, index - 1);

                      //  m.Items.Add("Change Lot ID");
                        m.Items.Add("Export CSV File");


                        m.ItemClicked += new ToolStripItemClickedEventHandler(m_ItemClicked);


                        m.Show(treeView1, new Point(e.X, e.Y));

                        x_Text = e.X;
                        y_Text = e.Y;


                    }

                    else if (treeView1.SelectedNode.Level == 3)
                    {
                        //  ContextMenuStrip m = new ContextMenuStrip();

                        //  m.Items.Add("Delete");
                        //  m.Items.Add("Hide");
                        //  m.Items.Add("UnHide");
                        //  m.ItemClicked += new ToolStripItemClickedEventHandler(m_ItemClicked);

                        //  m.Show(treeView1, new Point(e.X, e.Y));
                    }
                }
            }
            catch
            {

            }
        }

        public void AddDic_Tree_Info()
        {
            Node3 = treeView1.SelectedNode.Text;
            string[] dummy = Node3.Split('(');
            Node3 = dummy[0].Trim();

            dummy = treeView1.SelectedNode.FullPath.Split('\\');

            dummy = dummy[1].Split('(');
            Node2 = dummy[0].Trim();

            dummy = treeView1.SelectedNode.FullPath.Split('\\');

            dummy = dummy[0].Split('(');
            Node1 = dummy[0].Trim();

            int _Node_Find = 0;
            this.FindNode_List = new List<string>();

            int Selected_Index = treeView1.SelectedNode.Index;

            foreach (KeyValuePair<string, List<string>> key in this.Lot_Information)
            {
                if (Node1 == key.Key.ToString())
                {
                    this.FindNode_List.Add(Convert.ToString(_Node_Find));
                    Find_Scroll = _Node_Find;

                    for (int i = 0; i < key.Value.Count; i++)
                    {
                        if (Node2 == key.Value[i].ToString())
                        {
                            this.FindNode_List.Add(Convert.ToString(i));
                        }
                    }
                    ////for (int i = 0; i < Bin.Length; i++)
                    ////{
                    ////    if (Node3 == Bin[i].ToString())
                    ////    {
                    this.FindNode_List.Add(Convert.ToString(Selected_Index));
                    ////    }
                    ////}
                }
                _Node_Find++;
            }
            if (!this.FindNode.ContainsKey(Node1 + "_" + Node2 + "_" + Node3))
            {
                this.FindNode.Add(Node1 + "_" + Node2 + "_" + Node3, this.FindNode_List);
            }
        }

        public void Review()
        {
            string Query = "";
            Lot = new string[0];

            treeView1.Nodes.Clear();

            Matching_Lot_data();



            int j = 0;
            int loop = 0;
            Bin = new Dictionary<string, double[]>();

            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                loop = 0;

                Query = "";
                double[] Lot_Total = new double[0];
              //  string[] Sub_Count = new string[key.Value.Count];
                double[] Sub_Count = new double[0];
                double[] Bin_Total = new double[0];
                double[] Lot_Count = new double[0];
                double[] data = new double[0];
        
     
                Dictionary<string, List<string>> test = key.Value;

                int Lot_N = 0;

                _Node = new TreeNode[key.Value.Count];



                foreach ( KeyValuePair<string, List<string>> t in test)
                {

                    Query = "Select count(*) from " + key.Key;
                    Lot_Total = DB_Interface.Get_Find_Bin(Query);

                    _Node[Lot_N] = new TreeNode(t.Key.ToString() + " (" + Lot_Total[0] + ")");

                    int Bin_N = 0;
                    foreach (string L in t.Value)
                    {
                        Query = "Select count(*) from " + key.Key + " where LOTID = '" + t.Key.ToString() + "' and SUBLOT = '" + L + "'";
                        Sub_Count = DB_Interface.Get_Find_Bin(Query);

                        _Node[Lot_N].Nodes.Add(L + " (" + Sub_Count[0] + ")");


                        Query = "Select DISTINCT BIN from " + key.Key + " where LOTID = '" + t.Key.ToString() + "' and SUBLOT = '" + L + "'";
                        Bin_Total = DB_Interface.Get_Find_Bin(Query);

                        Array.Sort(Bin_Total);

                      
                        for (int k = 0; k < Bin_Total.Length; k ++)
                        {
                            Query = "Select count(BIN) from " + key.Key + " where LOTID = '" + t.Key.ToString() + "' and SUBLOT = '" + L + "' and Bin = '" + Bin_Total[k] + "'";
                            double[] Bin_Totals = DB_Interface.Get_Find_Bin(Query);

                            _Node[Lot_N].Nodes[Bin_N].Nodes.Add(Bin_Total[k] + " (" + Bin_Totals[0] + ")");
                        }
                        Bin_N++;
                    }

                    treeView1.Nodes.Add(_Node[Lot_N]);
                    Lot_N++;

                }

                j++;
                loop++;

            }
            treeView1.ExpandAll();
        }

        public void Gridview(bool Flag)
        {
            Dt.Columns.Clear();
            Dt.Rows.Clear();

            int i = 0;
            long total = 0;

            double[] Lots = new double[this.Lot_Information.Count];
            double[] Bin = new double[0];
            string Query = "";

            object[] Values = new object[2];
            double[] Bin_Total = new double[0];

            if (Flag)
            {
                Dt.Columns.Add("Info");
                dataGridView1.Columns[0].Width = 50;

                Dt.Columns.Add("Count");
                dataGridView1.Columns[1].Width = 60;
            }

            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                Dictionary<string, List<string>> test = key.Value;


                foreach (KeyValuePair<string, List<string>> t in test)
                {
                    Query = "Select DISTINCT BIN from " + key.Key;
                    double[] data = DB_Interface.Get_Find_Bin(Query);

                    Bin = Bin.Concat(data).ToArray();
                }

            }

            Bin = Bin.Distinct().ToArray();
            Array.Sort(Bin);

            for (int k = 0; k < Bin.Length; k++)
            {
                Bin_Total = new double[0];

                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {

                    Dictionary<string, List<string>> test = key.Value;


                    foreach (KeyValuePair<string, List<string>> t in test)
                    {
                        Query = "Select BIN from " + key.Key + " where BIN = '" + Bin[i] + "' AND FAIL like '0'";
                        double[] data = DB_Interface.Get_Find_Bin(Query);

                        Bin_Total = Bin_Total.Concat(data).ToArray();
                    }


                }

                Values[0] = "Bin" + Bin[i];
                Values[1] = Bin_Total.Length;

                total += Bin_Total.Length;
                Dt.Rows.Add(Values);
                i++;
            }

   


            Databind.DataMember = Dt.TableName;
            dataGridView1.Visible = true;


            Values[0] = "Total";
            Values[1] = total;
            Dt.Rows.Add(Values);

            Values[0] = "Hidden";

            double[] Dummy_Test = new double[0];


            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                Dictionary<string, List<string>> test = key.Value;

                foreach (KeyValuePair<string, List<string>> t in test)
                {
                    Query = "Select FAIL from " + key.Key + " where FAIL like '1'";

                    double[] data = DB_Interface.Get_Find_Bin(Query);

                    Dummy_Test = Dummy_Test.Concat(data).ToArray();
                }

            }


            Values[1] = Dummy_Test.Length;
            Dt.Rows.Add(Values);

            Databind.DataSource = Dt;

        }

        public void Gridview2(bool Flag)
        {
            Dt2.Rows.Clear();
            if (Flag)
            {
                Dt2.Columns.Clear();


                Dt2.Columns.Add("Current_LOT_ID");
              //  dataGridView2.Columns[0].Width = 60;
               // dataGridView2.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                Dt2.Columns.Add("Changes_LOT_ID");
             //   dataGridView2.Columns[1].Width = 60;
             //   dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                Dt2.Columns.Add("Current_WAFER_ID");
              //  dataGridView2.Columns[2].Width = 60;
              //  dataGridView2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                Dt2.Columns.Add("Changes_WAFER_ID");
             //   dataGridView2.Columns[3].Width = 60;
             //   dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }


            int i = 0;
            double[] Lots = new double[this.Lot_Information.Count];
            double[] Bin = new double[1];
            string Query = "";


            Databind2.DataMember = Dt.TableName;
            dataGridView2.Visible = true;

            Matching_Lot_data();
            DB_Interface.Matching_Lots = Matching_Lots;

            object[] Values = new object[4];
            int k = 0;
            i = 0;
            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                Dictionary<string, List<string>> tests = key.Value;


                foreach (KeyValuePair<string, List<string>> ts in tests)
                {

                        Dt2.Rows.Add(Values);
                        i++;

                }
            }

            //        foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, List<string>>>> key in this.Matching_Lots)
            //{
            //    Dictionary<string, Dictionary<string, List<string>>> test = key.Value;
          
            //    int k = 0;
            //    i = 0;
            //    foreach (KeyValuePair<string, Dictionary<string, List<string>>> keys in test)
            //    {
            //        Dictionary<string, List<string>> test1 = keys.Value;
               

            //        foreach (KeyValuePair<string, List<string>> t in test1)
            //        {
            //            if(i == 0)
            //            {
            //                /// dataGridView2[1, 2].Value = 1;
            //                //   Values[0] = t.Key;
            //                //   Values[1] = t.Key;
            //                Dt2.Rows.Add(Values);
            //                i++;
            //            }
            //            else
            //            {
                   
            //               //     Values[2] = t.Value[k];
            //               //     Values[3] = t.Value[k];
                            
            //                k++;
                     
            //            }
            //        }
               
            //    }
        
            //}


            int Lot_Col = 0;
            int Rpw = 0;

            int Wafer_Col = 2;

            for (int s = 0; s < Lot.Length; s++)
            {
                Query = "Select DISTINCT LOTID from " + Lot[s];
                string[] data = DB_Interface.Get_Data_By_Query(Query);
                string dummy = "";

                for (int p = 0; p < data.Length; p++)
                {

                    if (data.Length != 0)
                    {
                        if (p == data.Length - 1)
                        {
                            dummy += data[p];
                    
                            dataGridView2[0, Rpw].Value = dummy;
                            Rpw++;
                        }
                        else
                        {
                            dummy += data[p] + "&";
                        }


                    }
                    else
                    {
                        dataGridView2[0, Rpw].Value = data[p];
                    }

                }
            }
            Rpw = 0;
            for (int s = 0; s < Lot.Length; s++)
            {
                Query = "Select DISTINCT WAFER_ID from " + Lot[s];
                string[] data = DB_Interface.Get_Data_By_Query(Query);
                string dummy = "";

                for (int p = 0; p < data.Length; p++)
                {

                    if (data.Length != 0)
                    {
                        if (p == data.Length - 1)
                        {
                            dummy += data[p];
                            dataGridView2[2, Rpw].Value = dummy;
                            Rpw++;
                        }
                        else
                        {
                            dummy += data[p] + "&";
                        }


                    }
                    else
                    {
                        dataGridView2[2, Rpw].Value = data[p];
                    }

                }
            }


            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView2.RowHeadersVisible = false;
            dataGridView2.AllowUserToAddRows = false;
          //  dataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            //  dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dataGridView2.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f =>

            {

                f.SortMode = DataGridViewColumnSortMode.NotSortable; // sort 막기


            });


            Databind2.DataSource = Dt2;
         //   dataGridView2.Columns[0].Width = 10;
         //   dataGridView2.Columns[1].Width = 10;
         //   dataGridView2.Columns[2].Width = 10;
          //  dataGridView2.Columns[3].Width = 10;
        }

        public void Gridview3(bool Flag, List<FileInfo> Files)
        {

            data_type = new List<string>();

            for (int i = 0; i < Files.Count; i++)
            {
                string[] split = Files[i].Name.Split('_');

                if (!data_type.Contains(split[0]))
                {
                    data_type.Add(split[0]);
                }
            }

            FileInfo[] list = Files.Distinct().ToArray();


            if (Flag)
            {
                dataGridView3.ColumnCount = 1;
                dataGridView3.Columns[0].Name = "Data Type";
    

                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();

                dataGridView3.Columns.Add(chk);

                chk.HeaderText = "Check";
                // chk.Name = "ok";
            }


            for (int i = 0; i < data_type.Count; i++)
            {
                dataGridView3.Rows.Add("");

            }


            for (int i = 0; i < data_type.Count; i++)
            {

                dataGridView3.Rows[i].Cells[0].Value = data_type[i];
            }

            dataGridView3.Visible = true;

            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
         //   dataGridView3.RowHeadersVisible = false;
            dataGridView3.AllowUserToAddRows = false;
            //  dataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            //  dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dataGridView3.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f =>

            {

                f.SortMode = DataGridViewColumnSortMode.NotSortable; // sort 막기


            });

        }
        public static DialogResult InputBox(string title, string content, ref string value)
        {
            Form form = new Form();
            // PictureBox picture = new PictureBox();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.ClientSize = new Size(300, 100);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            form.Text = "Change Lot ID";

            label.Text = content;
            textBox.Text = content;
            buttonOk.Text = "확인";
            buttonCancel.Text = "취소";

            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;


            label.SetBounds(10, 17, 100, 20);
            textBox.SetBounds(10, 40, 280, 20);
            buttonOk.SetBounds(135, 70, 70, 20);
            buttonCancel.SetBounds(215, 70, 70, 20);

            DialogResult dialogResult = form.ShowDialog();

            value = textBox.Text;
            return dialogResult;
        }

        public void Listbox2_Define()
        {
            listBox2.Items.Clear();

            //for (int k = 0; k < Lot.Length; k++)
            //{
            //    string Query = "Select DISTINCT LOTID from " + Lot[k];
            //    string[] data = DB_Interface.Get_Data_By_Query(Query);

            //    Lot = Lot.Concat(data).ToArray();

            //}

            //Lot = Lot.Distinct().ToArray();
            //Array.Sort(Lot);



            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {

                Dictionary<string, List<string>> test = key.Value;


                foreach (KeyValuePair<string, List<string>> t in test)
                {
                    listBox2.Items.Add(t.Key);
                }
            }

            //foreach(KeyValuePair<string,string> d in Matching_Lots)
            //{
            //    listBox2.Items.Add(d.Key);
            //}







            //for (int k = 0; k < Lot.Length; k++)
            //{
            //    listBox2.Items.Add(Matching_Lots[k]);
            //}

            //listBox2.Height = 29 * Lot.Length;

            //   listBox2.Location.
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            //  this.SubLot = new List<string>();
            //  this.Lot_Information = new Dictionary<string, List<string>>();
            //  _Lot_Information_Dummy = new List<string>();

            this.listBox1.Items.Clear();

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                this.Files_Full_Name = (string[])e.Data.GetData(DataFormats.FileDrop);
                this.Files = new string[Files_Full_Name.Length];
                int i = 0;

                if (List_Files == null)
                {
                    List_Files = new List<string>();

                }
                foreach (string str in Files_Full_Name)
                {
                    if (!List_Files.Contains(str))
                    {
                        List_Files.Add(str);
                    }

                }

                List_Files.Sort();
                this.Files = new string[List_Files.Count];
                this.Files_Full_Name = new string[List_Files.Count];
                foreach (string str in List_Files)
                {
                    this.Files_path = "";
                    int lenth = str.LastIndexOf("\\");
                    this.Files[i] = str.Substring(str.LastIndexOf("\\") + 1);
                    this.Files_path = str.Substring(0, lenth);
                    this.Files_Full_Name[i] = str;
                    //string[] _Lot = Files[i].Split('_');

                    //if (!this.Lot_Information.ContainsKey(_Lot[1]))
                    //{
                    //    this.Lot_Information.Add(_Lot[1], this.SubLot);
                    //    _Lot_Information_Dummy.Add(_Lot[1]);
                    //}
                    this.listBox1.Items.Add(Files[i]);
                    i++;
                }

            }
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }

        private void listBox2_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            Brush brush;
            if (this.listBox2.SelectedIndex == e.Index)
                brush = Brushes.White;
            else
                brush = Brushes.Black;



            string output = listBox2.Items[e.Index].ToString();
            float olength = e.Graphics.MeasureString(output, e.Font).Width;
            float Height = e.Graphics.MeasureString(output, e.Font).Height;
            float pos = listBox2.Width - olength;
            brush = new SolidBrush(e.ForeColor);

            listBox2.ItemHeight = Convert.ToInt16(Height + 1);

            //  e.Graphics.DrawString(output, e.Font, brush, pos, e.Bounds.Top);
            e.Graphics.DrawString(output, e.Font, brush, 0, e.Bounds.Top);

            listBox2.Height = Convert.ToInt16(e.Graphics.MeasureString(output, e.Font).Height * Lot.Length) + 10;
            //  listBox2.Width = Convert.ToInt16(olength) + 10;
            listBox2.Width = Convert.ToInt16(140) + 10;

            ////Font의 Height를 더한 만큼 좌표를 변경합니다.
            //int x = e.Bounds.X + e.Font.Height;
            //int y = e.Bounds.Y + e.Font.Height;

            //e.Graphics.DrawString(this.listBox2.Items[e.Index].ToString(),
            //    e.Font, brush, x, y, StringFormat.GenericDefault);
            //e.DrawFocusRectangle();
            int X = listBox2.Location.X;
            int Y = listBox2.Location.Y;


          //  dataGridView1.Location = new Point(X, Y + Convert.ToInt16(e.Graphics.MeasureString(output, e.Font).Height * Lot.Length) + 10);

        }

        private void listBox2_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void listBox2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {

                ContextMenuStrip m2 = new ContextMenuStrip();

                m2.Items.Add("Export CSV File");

                m2.ItemClicked += new ToolStripItemClickedEventHandler(m2_ItemClicked);

                m2.Show(listBox2, new Point(e.X, e.Y));


            }


        }

        static List<FileInfo> DirSeach(string Dir, List<FileInfo> temp = null)
        {
            if (temp == null) temp = new List<FileInfo>();
            DirectoryInfo di = new DirectoryInfo(Dir);

            foreach (var s in Directory.GetDirectories(Dir)) DirSeach(s, temp);
            foreach (var item in di.GetFiles()) temp.Add(item);
            return temp;
        }

        private void button7_Click(object sender, EventArgs e)
        {


            Matching_Lot_data();

            for (int j = 0; j < Node_Find.Count; j++)
            {
            

                string[] split_name = Node_Find[j].Split(',');


                string D = "";

                foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                {
                    Dictionary<string, List<string>> test = key.Value;


                    foreach (KeyValuePair<string, List<string>> t in test)
                    {
                        if (t.Key == split_name[0])
                        {
                            D = key.Key;
                        }
                    }

                }

                if (split_name.Length == 1)
                {

                    DB_Interface.trans(Data_Interface);
                        string Query = "Delete from " + D + " where LOTID = '" + split_name[0] + "'";
                        DB_Interface.Delete_Lot_Data(Query);
                        DB_Interface.Commit(Data_Interface);
                  //  }
                }
                else if (split_name.Length == 2)
                {
                   // for (int k = 0; k < DB_Interface.Table_Count; k++)
                  //  {
                        DB_Interface.trans(Data_Interface);
                        string Query = "Delete from " + D + " where LOTID = '" + split_name[0] + "' and  SUBLOT = '" + split_name[1] + "'";
                        DB_Interface.Delete_Lot_Data(Query);
                        DB_Interface.Commit(Data_Interface);
                  //  }
                }
                else if (split_name.Length == 3)
                {
                  //  for (int k = 0; k < DB_Interface.Table_Count; k++)
                  //  {
                        DB_Interface.trans(Data_Interface);
                        string Query = "Delete from " + D + " where LOTID = '" + split_name[0] + "' and  SUBLOT = '" + split_name[1] + "' and BIN = '" + split_name[2] + "'";
                        DB_Interface.Delete_Lot_Data(Query);
                        DB_Interface.Commit(Data_Interface);
                 //   }
                }
         
            }
            // DB_Interface.Commit(Data_Interface);
          //  Matching_Lot_data();
            this.FindNode = new Dictionary<string, List<string>>();
            Review();
            Gridview(true);
            Gridview2(true);
            Listbox2_Define();

            Node_Find = new List<string>();
            foreach (KeyValuePair<string, List<string>> key in this.FindNode)
            {
                treeView1.Nodes[Convert.ToInt16(key.Value[0])].Nodes[Convert.ToInt16(key.Value[1])].Nodes[0].Nodes[Convert.ToInt16(key.Value[2])].ForeColor = Color.Red;
            }
        
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            treeView1.AfterCheck -= treeView1_AfterCheck;
            ChildNodeChecking(e.Node);
            //   ParentNodeChecking(e.Node);
            treeView1.AfterCheck += treeView1_AfterCheck;


        }

        private void ChildNodeChecking(TreeNode selectNode)
        {
            string[] Node_path = selectNode.FullPath.Split('\\');

            if (selectNode.Checked)
            {
                if (Node_path.Length == 1)
                {
                    int index = Node_path[0].IndexOf('(');
                    Node_path[0] = Node_path[0].Substring(0, index - 1);


                    if (Node_Find.Contains(Node_path[0]))
                    {

                    }
                    else
                    {
                        Node_Find.Add(Node_path[0]);
                    }
                }
                else if (Node_path.Length == 2)
                {
                    int index = Node_path[0].IndexOf('(');
                    Node_path[0] = Node_path[0].Substring(0, index - 1);

                    index = Node_path[1].IndexOf('(');
                    Node_path[1] = Node_path[1].Substring(0, index - 1);

                    if (Node_Find.Contains(Node_path[0] + "," + Node_path[1]))
                    {

                    }
                    else
                    {
                        Node_Find.Add(Node_path[0] + "," + Node_path[1]);
                    }
                }
                else if (Node_path.Length == 3)
                {
                    int index = Node_path[0].IndexOf('(');
                    Node_path[0] = Node_path[0].Substring(0, index - 1);

                    index = Node_path[1].IndexOf('(');
                    Node_path[1] = Node_path[1].Substring(0, index - 1);

                    index = Node_path[2].IndexOf('(');
                    Node_path[2] = Node_path[2].Substring(0, index - 1);

                    if (Node_Find.Contains(Node_path[0] + "," + Node_path[1] + "," + Node_path[2]))
                    {

                    }
                    else
                    {
                        Node_Find.Add(Node_path[0] + "," + Node_path[1] + "," + Node_path[2]);
                    }
                }
            }
            else
            {
                if (Node_path.Length == 1)
                {
                    int index = Node_path[0].IndexOf('(');
                    Node_path[0] = Node_path[0].Substring(0, index - 1);


                    if (Node_Find.Contains(Node_path[0]))
                    {
                        Node_Find.Remove(Node_path[0]);
                    }

                }

                else if (Node_path.Length == 2)
                {
                    int index = Node_path[0].IndexOf('(');
                    Node_path[0] = Node_path[0].Substring(0, index - 1);

                    index = Node_path[1].IndexOf('(');
                    Node_path[1] = Node_path[1].Substring(0, index - 1);

                    if (Node_Find.Contains(Node_path[0] + "," + Node_path[1]))
                    {
                        Node_Find.Remove(Node_path[0] + "," + Node_path[1]);
                    }

                }
                else if (Node_path.Length == 3)
                {
                    int index = Node_path[0].IndexOf('(');
                    Node_path[0] = Node_path[0].Substring(0, index - 1);

                    index = Node_path[1].IndexOf('(');
                    Node_path[1] = Node_path[1].Substring(0, index - 1);

                    index = Node_path[2].IndexOf('(');
                    Node_path[2] = Node_path[2].Substring(0, index - 1);

                    if (Node_Find.Contains(Node_path[0] + "," + Node_path[1] + "," + Node_path[2]))
                    {
                        Node_Find.Remove(Node_path[0] + "," + Node_path[1] + "," + Node_path[2]);
                    }
                }
            }
            return;
        }

        public void ParentNodeChecking(TreeNode selectNode)
        {
            TreeNode t = selectNode.Parent;
            if (t != null)
            {
                t.Checked = true;
                foreach (TreeNode tn in t.Nodes)
                {
                    if (!tn.Checked)
                    {
                        t.Checked = false; break;
                    }
                }
                ParentNodeChecking(t);
            }
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {


            if (e.Control && e.KeyCode == Keys.C)
            {
                DataObject Do = dataGridView2.GetClipboardContent();
                Clipboard.SetDataObject(Do);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int row = dataGridView2.CurrentCell.RowIndex;
                int col = dataGridView2.CurrentCell.ColumnIndex;
                foreach (string line in lines)
                {
                    if (row < dataGridView2.RowCount && line.Length > 0)
                    {
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.GetLength(0); ++i)
                        {
                            if (col + i < dataGridView2.ColumnCount)
                            {
                                dataGridView2[col + i, row].Value =
                                Convert.ChangeType(cells[i], dataGridView2[col + i, row].ValueType);
                            }
                            else
                            {
                                break;
                            }
                        }
                        row++;
                    }
                    else
                    {
                        break;
                    }
                }
                //  bindingSource[index].DataSource = _dataTable[index];
                dataGridView2.Update();
            }
            else if (e.Control && e.KeyCode == Keys.Z)
            {


            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {


            Matching_Lot_data();
            int row = 0;
            Edit_data = new List<string>();

 
 
            for (int k = 0; k < Lot.Length; k++)
            {
                string imformation = "";


                for (int i = 0; i < 4; i++)
                {
                    if (i == 3)
                    {
                        imformation += dataGridView2[i, row].Value.ToString();

                    }
                    else
                    {
                        imformation += dataGridView2[i, row].Value.ToString() + ",";
                    }

                }
                Edit_data.Add(imformation);
                row++;
            }

            // DB_Interface.DropTable(Data_Interface, "B909LL333");

            Progress_Form Progress = new Progress_Form("MERGE", 0);
            Progress.Merge_Init(Edit_data.Count);
            Progress.Show();


            //  DB_Interface.trans(Data_Interface);
            for (int j = 0; j < Edit_data.Count; j++)
            {

                string[] Edits = Edit_data[j].Split(',');

                Stopwatch TestTime1 = new Stopwatch();
                TestTime1.Restart();
                TestTime1.Start();

                for (int index = 1; index < 2; index++)
                {
                    string Query = "";
                    string Query2 = "";
                    for (int table = 0; table < 1; table++)
                    {

                        if (Edits[2] != Edits[3] && Edits[3] != "")
                        {
                            DB_Interface.trans(Data_Interface);
                            Query = "Update " + Lot[j] + " set WAFER_ID = '" + Edits[3] + "'";
                            Query2 = "";

                            DB_Interface.LOTID_Update(Query, Query2, "WAFER_ID");
                            DB_Interface.Commit(Data_Interface);
                        }

                        if (Edits[1] != Edits[0] && Edits[1] != "")
                        {
                            DB_Interface.trans(Data_Interface);
                            Query = "Update " + Lot[j] + " set LOTID = '" + Edits[1] + "'";
                            Query2 = "Update " + Lot[j] + " set LOT_ID = '" + Edits[1] + "'";


                            DB_Interface.LOTID_Update(Query, Query2, "LOTID");
                            DB_Interface.Commit(Data_Interface);
         


                        }

                    }

                    Progress.Merge_Print(Edit_data.Count, j);

                }
                double ss = TestTime1.ElapsedMilliseconds; 
            }

            //Matching_Lots = new Dictionary<string, string>();

            //for (int Lot_index = 0; Lot_index < Lot.Length; Lot_index++)
            //{
            //    string Lot_string = Lot[Lot_index];
            //    string[] Sub_Lot = new string[0];

            //    string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
            //    string[] Lot_data = DB_Interface.Get_Data_By_Query(Query);

            //    Query = "Select DISTINCT SUBLOT from " + Lot[Lot_index];
            //    //  Query = "Select DISTINCT SUBLOT from " + Lot[Lot_index] + " where LotID = '" + Lot_data[0] + "'";
            //    string[] data = DB_Interface.Get_Data_By_Query(Query);


            //    Matching_Lots.Add(Lot_data[0], Lot[Lot_index]);

            //}

            //  Gridview(true);

            Matching_Lot_data();

            Gridview2(false);
            Review();
            Listbox2_Define();
            listBox2.Show();
            Progress.Close();


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
                else if (Lot[i] == "Trace_Info")
                {
                    Lot[i] = ""; k++;
                }
                else if (Lot[i].Contains("CHAN"))
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

            Matching_Lots = new Dictionary<string, Dictionary<string, List<string>>>();

            for (int Lot_index = 0; Lot_index < Lot.Length; Lot_index++)
            {
                string Lot_string = Lot[Lot_index];
                string[] Sub_Lot = new string[0];

                // string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
                string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
                string[] Lot_data = DB_Interface.Get_Data_By_Query(Query);
                Lot_Information = new Dictionary<string, List<string>>();

                Query = "Select count(*) from " + Lot[Lot_index];
                string[] dummy = DB_Interface.Get_Data_By_Query(Query);

                if (Convert.ToInt16(dummy[0]) == 0)
                {
                    DB_Interface.DropTable(Data_Interface, Lot[Lot_index]);
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
                    Matching_Lots.Add(Lot[Lot_index], Lot_Information);
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
                    Matching_Lots.Add(Lot[Lot_index], Lot_Information);
                }
            }
        }

        //public void Remove_table()
        //{

        //    int k = 0;
        //    for(int i = 0; i < Lot.Length; i++)
        //    {
        //        if(Lot[i] == "Clotho_Spec")
        //        {
              
        //            Lot[i] = ""; k++;
        //        }
        //        else if (Lot[i] == "Files")
        //        {
        //            Lot[i] = ""; k++;
        //        }
        //        else if (Lot[i] == "INF")
        //        {
        //            Lot[i] = ""; k++;

        //        }
        //        else if (Lot[i] == "REFHEADER")
        //        {
        //            Lot[i] = "";k++;
        //        }
        //        else if (Lot[i] == "Customer_Spec")
        //        {
        //            Lot[i] = ""; k++;
        //        }

        //    }

        //    Lot = Lot.Where(x => !string.IsNullOrEmpty(x)).ToArray();
        //}

        //public void Matching_Lot_data()
        //{
        //    Lot = new string[0];

        //    for (int k = 0; k < 1; k++)
        //    {
        //        // string Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

        //        string Query = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY Name";

        //        Lot = DB_Interface.Get_Data_By_Query(Query);

        //    }

        //    Remove_table();

        //    Matching_Lots = new Dictionary<string, Dictionary<string, List<string>>> ();

        //    Matching_Lots_Test = new Dictionary<string, Dictionary<string, Dictionary<string, List<string>>>>();

        //    for (int Lot_index = 0; Lot_index < Lot.Length; Lot_index++)
        //    {
        //        string Lot_string = Lot[Lot_index];
        //        string[] Sub_Lot = new string[0];

        //       // string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
        //        string Query = "Select DISTINCT LotID from " + Lot[Lot_index];
        //        string[] Lot_data = DB_Interface.Get_Data_By_Query(Query);
        //        Lot_Information = new Dictionary<string, List<string>>();
        //        information = new Dictionary<string, Dictionary<string, List<string>>>();


        //        Query = "Select count(*) from " + Lot[Lot_index];
        //        string[] dummy = DB_Interface.Get_Data_By_Query(Query);

        //        if (Convert.ToInt16(dummy[0]) == 0)
        //        {
      
        //            DB_Interface.trans(Data_Interface);
        //            DB_Interface.DropTable(Data_Interface, Lot[Lot_index]);
        //            DB_Interface.Commit(Data_Interface);
        //            break;
        //        }

        //        if(Lot_index == 18)
        //        {

        //        }
        //        if (Lot_data.Length != 0 && Lot_data.Length == 1)
        //        {
        //            Query = "Select DISTINCT SUBLOT from " + Lot[Lot_index] + " where LotID = '" + Lot_data[0] + "'";
        //            string[] data = DB_Interface.Get_Data_By_Query(Query);

        //            Sub_Lot = Sub_Lot.Concat(data).ToArray();

        //            Sub_Lot = Sub_Lot.Distinct().ToArray();
        //            Array.Sort(Sub_Lot);

        //            _Lot_Information_Dummy = new List<string>();

        //            for (int k = 0; k < Sub_Lot.Length; k++)
        //            {
        //                _Lot_Information_Dummy.Add(Sub_Lot[k]);
        //            }

        //            Lot_Information.Add(Lot_data[0], _Lot_Information_Dummy);
        //            information.Add("LOTID", Lot_Information);


        //            Matching_Lots.Add(Lot[Lot_index], Lot_Information);
             
        //            string[] Wafer = new string[0];

        //            Query = "Select DISTINCT WAFER_ID from " + Lot[Lot_index];
        //            data = DB_Interface.Get_Data_By_Query(Query);

        //            Wafer = Wafer.Concat(data).ToArray();

        //            Wafer = Wafer.Distinct().ToArray();
        //            Array.Sort(Wafer);

        //            _Lot_Information_Dummy = new List<string>();
        //            Lot_Information = new Dictionary<string, List<string>>();

        //            for (int k = 0; k < Wafer.Length; k++)
        //            {
        //                _Lot_Information_Dummy.Add(Wafer[k]);
        //            }

        //            Lot_Information.Add(Lot_data[0], _Lot_Information_Dummy);
        //            information.Add("WAFERID", Lot_Information);

        //            Matching_Lots_Test.Add(Lot[Lot_index], information);
        //        }
        //        else if(Lot_data.Length > 1)
        //        {
        //            for (int i = 0; i < Lot_data.Length; i++)
        //            {
        //                Query = "Select DISTINCT SUBLOT from " + Lot[Lot_index] + " where LotID = '" + Lot_data[i] + "'";
        //                string[] data = DB_Interface.Get_Data_By_Query(Query);

        //                Sub_Lot = Sub_Lot.Concat(data).ToArray();

        //                Sub_Lot = Sub_Lot.Distinct().ToArray();
        //                Array.Sort(Sub_Lot);

        //                _Lot_Information_Dummy = new List<string>();

        //                for (int k = 0; k < Sub_Lot.Length; k++)
        //                {
        //                    _Lot_Information_Dummy.Add(Sub_Lot[k]);
        //                }

        //                   Lot_Information.Add(Lot_data[i], _Lot_Information_Dummy);

        //            }
        //            information.Add("LOTID", Lot_Information);
        //            Matching_Lots.Add(Lot[Lot_index], Lot_Information);


        //            Lot_Information = new Dictionary<string, List<string>>();

        //            for (int i = 0; i < Lot_data.Length; i++)
        //            {
        //                string[] Wafer = new string[0];

        //                Query = "Select DISTINCT WAFER_ID from " + Lot[Lot_index];
        //                string[] data = DB_Interface.Get_Data_By_Query(Query);

        //                Wafer = Wafer.Concat(data).ToArray();

        //                Wafer = Wafer.Distinct().ToArray();
        //                Array.Sort(Wafer);

        //                _Lot_Information_Dummy = new List<string>();
                 

        //                for (int k = 0; k < Wafer.Length; k++)
        //                {
        //                    _Lot_Information_Dummy.Add(Wafer[k]);
        //                }

        //                Lot_Information.Add(Lot_data[i], _Lot_Information_Dummy);
               

        //            }

        //            information.Add("WAFERID", Lot_Information);
    

        //            Matching_Lots_Test.Add(Lot[Lot_index], information);
        //        }
        //    }
        //}

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            bool isChecked = Convert.ToBoolean(dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue);
            if (isChecked)
            {
                if (!Selected_data_type.Contains(dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString()))
                {
                    Selected_data_type.Add(dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString());
                }

            }
            else
            {
                if (Selected_data_type.Contains(dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString()))
                {
                    Selected_data_type.Remove(dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString());
                }
            }
        }

        public void Make_up_File_List()
        {
            Dic_File = new Dictionary<string, List<FileInfo>>();
            if (checkBox2.Checked)
            {
                for (int i = 0; i < Selected_data_type.Count; i++)
                {
                    List<FileInfo> Dummy = new List<FileInfo>();

                    for (int k = 0; k < Filedata.Count; k++)
                    {

                        if (Filedata[k].Name.Substring(Filedata[k].Name.Length - 5, 5).ToUpper() == ".S4PD")
                        {
                            string[] split = Filedata[k].Name.Split('_');

                            if (split[0] == Selected_data_type[i])
                            {
                                Dummy.Add(Filedata[k]);
                            }


                        }
                    }
               
                    Dic_File.Add(Selected_data_type[i], Dummy);
                }
            }
            else
            {
                for (int i = 0; i < Selected_data_type.Count; i++)
                {
                    List<FileInfo> Dummy = new List<FileInfo>();

                    for (int k = 0; k < Filedata.Count; k++)
                    {

                        if (Filedata[k].Name.Substring(Filedata[k].Name.Length - 4, 4).ToUpper() == ".CSV")
                        {
                            string[] split = Filedata[k].Name.Split('_');

                            if (split[0] == Selected_data_type[i])
                            {
                                Dummy.Add(Filedata[k]);
                            }


                        }
                    }

                    Dic_File.Add(Selected_data_type[i], Dummy);
                }
            }
        
        }

        public void Merge_S4pd()
        {
            if (Selected_data_type.Count != 0)
            {
                Key = "MERGE_S4PD";
                string selected = Dialog.SelectedPath;

                Csv_Interface = CSV.Open(Key);
                Data_Interface = Data_Edit.Open(Key);
                DB_Interface = DB.Open(Key);

                DirectoryInfo di = new DirectoryInfo(selected);

                Filedata = DirSeach(selected);

                int data_Count = 0;
                for (int k = 0; k < Filedata.Count; k++)
                {
                    int index = Filedata[k].Name.Length;
                    string ss = Filedata[k].Name;
                    string Dumy = Filedata[k].Name.ToString().ToUpper().Substring(index - 4, 4);

                    int a = 0;
                    if (Dumy == ".CSV")
                    {
                        data_Count++;
                    }

                }


                Make_up_File_List();

                long Data_Count = 0;
                Data_Count = 0;
                bool Flag = true;
                foreach (KeyValuePair<string, List<FileInfo>> D in Dic_File)
                {

                    Progress_Form Progress = new Progress_Form("MERGE", 0);
                    Progress.Merge_Init(D.Value.Count);
                    Progress.Show();


                    long Data_Count_for_Limit = 0;
                    int Count = 0;

                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();


                    DB_Interface._SUBLOT_Flag = false;
                    bool Flag_Tran = true;
                    foreach (FileInfo L in D.Value)
                    {

                        string[] split = L.Name.Split('_');

                        string table_name = split[4];
                        string s = table_name.Substring(0, table_name.Length - 5);

                        DB_Interface.Table = table_name;
                        DB_Interface.Lot_ID = split[0].Replace('-', '_');
                        DB_Interface.SubLot_ID = "0";
                        DB_Interface.Tester_ID = "0";
                        DB_Interface.Site = "0";
                        DB_Interface.Bin = "0";
                        DB_Interface.ID_Unit = split[6].Split('.')[0];
       

                        Csv_Interface.Read_Open(L.FullName);
                   

                        #region

                        if (Count == 0)
                        {
                            while (!Csv_Interface.StreamReader.EndOfStream)
                            {
                                Csv_Interface.Read();

                                if (Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1].ToUpper().Contains("SUBLOT"))
                                {
                                    List<string> Array = Csv_Interface.Get_String.ToList();

                                    Array.RemoveAt(Csv_Interface.Get_String.Length - 1);

                                    Csv_Interface.Get_String = Array.ToArray();

                                    DB_Interface._SUBLOT_Flag = true;
                                }

                                Flag = Data_Interface.Find_First_Row(Csv_Interface.Get_String);


                                if (Flag) break;

                                if (Csv_Interface.Get_String[0].ToUpper() == "LOT")
                                {
                                    DB_Interface.Lot_ID = Csv_Interface.Get_String[1];
                                    DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "SUBLOT")
                                {
                                    DB_Interface.SubLot_ID = Csv_Interface.Get_String[1];
                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "HOSTIPADDRESS")
                                {
                                    DB_Interface.Tester_ID = Csv_Interface.Get_String[1];
                                }

                            }

                            Data_Interface.Define_DB_Count(Csv_Interface.Get_String);

                            for (int l = 0; l < Csv_Interface.Get_String.Length; l++)
                            {
                                if (Csv_Interface.Get_String[l].ToUpper() == "SBIN")
                                {
                                    DB_Interface.Bin_place = 1;

                                }

                            }


                            if (checkBox1.Checked)
                            {


                                DB_Interface.Open_DB(Dialog1.FileNames, Data_Interface);


                                if (Flag_Tran)
                                {
                                    DB_Interface.trans(Data_Interface);
                                    Flag_Tran = false;
                                }


                                string[] Filedatas = new string[0];
                                string Query = "";

                     

                                for (int k = 0; k < 1; k++)
                                {
                                 
                                    Query = "Select File from Files where File = '" + L.Name + "'";
                                    Filedatas = DB_Interface.Get_Data_By_Query(Query);

                                }

                                if (Filedatas.Length != 0)
                                {
                                    Csv_Interface.Read_Close();
                                    goto Next;
                                }




                                Data_Interface.Make_New_header();

                            }
                            else
                            {
                                Dir = new Dir.Dir_Directory("C:\\Automation\\DB\\YIELD\\" + L.Name);

                                Data_Interface.Make_New_header();

                                DB_Interface.Open_DB(L.Name, Data_Interface);
                                DB_Interface.DropTable(Data_Interface, "");
                                if (Flag_Tran)
                                {
                                    DB_Interface.trans(Data_Interface);
                                    Flag_Tran = false;
                                }
                            }


                            Flag = false;

                            while (!Csv_Interface.StreamReader.EndOfStream)
                            {
                                Csv_Interface.Read();
                              //  if (Csv_Interface.Get_String[0].Contains("PID"))
                             //   {
                                    Flag = true;
                             //   }
                             //   else if (Csv_Interface.Get_String[0].Contains("HighL"))
                           //     {
                            //        Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                            //    }
                             //   else if (Csv_Interface.Get_String[0].Contains("LowL"))
                              //  {
                              //      Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                              //  }
                                if (Flag) break;
                            }
                            if (checkBox1.Checked)
                            {
                                string[] Table = new string[0];

                             //   DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');
                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Table + "'";

                                    Table = DB_Interface.Get_Data_By_Query(Query);
                                }

                                if (Table.Length == 0)
                                {
                                    DB_Interface.Clotho_Spec_Flag = false;
                                    DB_Interface.Insert_Header(Data_Interface);
                                }

                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "Select DISTINCT id from " + DB_Interface.Table;
                                    string[] data = DB_Interface.Get_Data_By_Query(Query);
                                    Data_Count = data.Length;
                                    Data_Count++;
                                }



                            }
                            else
                            {
                                DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                                DB_Interface.Clotho_Spec_Flag = true;
                                DB_Interface.Insert_Header(Data_Interface);
                                DB_Interface.Insert_Spec_Data("Clotho_Spec");
                                DB_Interface.Make_table("INF");
                                DB_Interface.Make_table2(Data_Interface, "REFHEADER");

                                DB_Interface.Make_table_For_Filename(Data_Interface, "Files");


                                DB_Interface.Insert_Ref_Header_Data(Data_Interface);


                                Data_Interface.Data_Table = "data0";
                                DB_Interface.Clotho_Spec_Flag = false;
                            }

                        }
                        else
                        {
                            string[] Filedatas = new string[0];

                            for (int k = 0; k < 1; k++)
                            {

                                string Query = "Select File from Files where File = '" + L.Name + "'";
                                Filedatas = DB_Interface.Get_Data_By_Query(Query);

                            }

                            if (Filedatas.Length != 0)
                            {
                                Csv_Interface.Read_Close();
                                goto Next;
                            }
                            else
                            {


                            }
                            if (Flag_Tran)
                            {
                                DB_Interface.trans(Data_Interface);
                                Flag_Tran = false;
                            }
                             
                         
                            Flag = false;
                            while (!Csv_Interface.StreamReader.EndOfStream)
                            {
                                Csv_Interface.Read();


                                if (Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1].Contains("SUBLOT"))
                                {
                                    List<string> Array = Csv_Interface.Get_String.ToList();

                                    Array.RemoveAt(Csv_Interface.Get_String.Length - 1);

                                    Csv_Interface.Get_String = Array.ToArray();
                                    //  Flag = true;
                                    DB_Interface._SUBLOT_Flag = true;
                                }

                                // /i/*f (Flag) break;*/

                                if (Csv_Interface.Get_String[0].ToUpper() == "LOT")
                                {
                                    DB_Interface.Lot_ID = Csv_Interface.Get_String[1];
                                    DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');
                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "SUBLOT")
                                {
                                    DB_Interface.SubLot_ID = Csv_Interface.Get_String[1];
                                }
                                else if (Csv_Interface.Get_String[0].ToUpper() == "HOSTIPADDRESS")
                                {
                                    DB_Interface.Tester_ID = Csv_Interface.Get_String[1];
                                }
                                if (Csv_Interface.Get_String[0].ToUpper().Contains("FREQ"))
                                {
                                    Flag = true;
                                    Csv_Interface.Read();
                                }
                                else if (Csv_Interface.Get_String[0].Contains("HighL"))
                                {
                                    Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                                }
                                else if (Csv_Interface.Get_String[0].Contains("LowL"))
                                {
                                    Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, false);

                                    DB_Interface.Delete_Spec_Data("Clotho_Spec");
                                    DB_Interface.Insert_Spec_Data("Clotho_Spec");
                                }
                                if (Flag) break;
                            }
                            DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');


                            string[] Table = new string[0];

                            for (int k = 0; k < 1; k++)
                            {
                                string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Table + "'";

                                Table = DB_Interface.Get_Data_By_Query(Query);
                            }

                            if (Table.Length == 0)
                            {
                                DB_Interface.Clotho_Spec_Flag = false;
                                DB_Interface.Insert_Header(Data_Interface);
                            }


                            DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                            for (int k = 0; k < 1; k++)
                            {
                                string Query = "Select id from " + DB_Interface.Table;
                                string[] data = DB_Interface.Get_Data_By_Query(Query);
                                //  Data_Count = data.Length;
                                Data_Count++;
                            }



                        }

                        #endregion

                        DB_Interface.Insert_ThreadFlags = new ManualResetEvent[2];
                        DB_Interface.Insert_Thread_Wait = new bool[2];

                        DB_Interface.TheFirst_Trashes_Header_Count = Data_Interface.TheFirst_Trashes_Header_Count;
                        DB_Interface.TheEnd_Trashes_Header_Count = Data_Interface.TheEnd_Trashes_Header_Count;

                        DB_Interface.TheFirst_Trashes_Header_Count = 0;
                        DB_Interface.TheEnd_Trashes_Header_Count = 0;

                        Data_Interface.Getstring = Csv_Interface.Get_String;

                      //  DB_Interface.Bin = Csv_Interface.Get_String[DB_Interface.Bin_place];

                      //  Data_Interface.Getstring[8] = Data_Interface.Getstring[8].Replace('-', '_');

                        if (DB_Interface._SUBLOT_Flag)
                        {
                            //DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                            //DB_Interface.SubLot_ID = Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1];
                            //DB_Interface.Tester_ID = Csv_Interface.Get_String[5];
                            //DB_Interface.Site = Csv_Interface.Get_String[5];
                            //DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                        }
                        else
                        {
                    //        DB_Interface.Site = DB_Interface.Tester_ID;
                        }

                        Data_Count++;
                        Data_Count_for_Limit++;

                        for (int thread_i = 0; thread_i < 2; thread_i++)
                        {
                            DB_Interface.Insert_ThreadFlags[thread_i] = new ManualResetEvent(false);
                        }
                        ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));


                        DB_Interface.Insert_ThreadFlags[1].Set();

                        DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                        DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

                        string[] GetData = Csv_Interface.Read();

                        Data_Interface.Getstring = GetData;
                      //  Csv_Interface.Get_String[8] = Csv_Interface.Get_String[8].Replace('-', '_');
                      //  DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

                        Data_Count++;
                        Data_Count_for_Limit++;


                        while (!Csv_Interface.StreamReader.EndOfStream)
                        {
                            if (DB_Interface.Lot_ID != split[0].Replace('-', '_'))
                            {
                                DB_Interface.Lot_ID = Csv_Interface.Get_String[8].Replace('-', '_');

                                string[] Table = new string[0];

                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Lot_ID + "'";

                                    Table = DB_Interface.Get_Data_By_Query(Query);
                                }

                                if (Table.Length == 0)
                                {
                                    DB_Interface.Clotho_Spec_Flag = false;
                                    DB_Interface.Insert_Header(Data_Interface);
                                }



                            }

                            for (int thread_i = 0; thread_i < 2; thread_i++)
                            {
                                DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                            }
                            ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

                            GetData = Csv_Interface.Read_Test();
                         //   Csv_Interface.Get_String[8] = Csv_Interface.Get_String[8].Replace('-', '_');

                            if (DB_Interface.Lot_ID != split[0].Replace('-', '_'))
                            {
                             //   DB_Interface.Lot_ID = Csv_Interface.Get_String[8].Replace('-', '_');

                                string[] Table = new string[0];

                                for (int k = 0; k < 1; k++)
                                {
                                    string Query = "SELECT name FROM sqlite_master WHERE type='table' AND name = '" + DB_Interface.Lot_ID + "'";

                                    Table = DB_Interface.Get_Data_By_Query(Query);
                                }

                                if (Table.Length == 0)
                                {
                                    DB_Interface.Clotho_Spec_Flag = false;
                                    DB_Interface.Insert_Header(Data_Interface);
                                }


                                //   DB_Interface.Lot_ID = DB_Interface.Lot_ID.Replace('-', '_');

                            }

                            DB_Interface.Insert_ThreadFlags[1].Set();

                            DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                            DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

                            Data_Interface.Getstring = GetData;

                        //    Csv_Interface.Get_String[8] = Csv_Interface.Get_String[8].Replace('-', '_');
                            if (DB_Interface._SUBLOT_Flag)
                            {

                                //DB_Interface.Lot_ID = Csv_Interface.Get_String[8];
                                //DB_Interface.SubLot_ID = Csv_Interface.Get_String[Csv_Interface.Get_String.Length - 1];
                                //DB_Interface.Tester_ID = Csv_Interface.Get_String[5];
                                //DB_Interface.Site = Csv_Interface.Get_String[5];



                            }

                          //  DB_Interface.Bin = Data_Interface.Getstring[DB_Interface.Bin_place];

                            Progress.Merge_Print(Data_Count, Count);

                            if(Data_Count_for_Limit == 300)
                            {
                                DB_Interface.Commit(Data_Interface);
                                Flag_Tran = true;
                            }
                              
                            Data_Count++;
                            Data_Count_for_Limit++;
                        }

                        for (int thread_i = 0; thread_i < 2; thread_i++)
                        {
                            DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                        }
                        ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { DB_Interface.Insert_Data(Data_Count); }));

                        DB_Interface.Insert_ThreadFlags[1].Set();

                        DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                        DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();


                        Csv_Interface.Read_Close();

                        DB_Interface.Insert_Files_Name(L.Name);

                     //   DB_Interface.Commit(Data_Interface);

                        Next:

                        Count++;
                    }
                    Progress.Close();
                }

                DB_Interface.Commit(Data_Interface);
                Review();
                Gridview(true);
                Gridview2(true);
                Listbox2_Define();
                listBox2.Show();

            }
            else
            {
                MessageBox.Show("Please Select Data Type.");
            }
            Selected_data_type = null;
        }

        public void DataCheck_S4pd()
        {
            string[] Trace = new string[0];

            for (int k = 0; k < 1; k++)
            {
                // string Query = "select count(*) from sqlite_master where name = 'data" + k + "'";

                string Query = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY Name";

                Trace = DB_Interface.Get_Data_By_Query(Query);

            }

            for (int k = 0; k < Trace.Length; k++)
            {
                if (!Trace[k].Contains("CHAN"))
                {
                    Trace[k] = "";
                }
                if (Trace[k].Contains("CHAN"))
                {
                    Trace[k] = Trace[k].Remove(0, 4);
                }
            }


            Trace = Trace.Where(x => !string.IsNullOrEmpty(x)).ToArray();

            int[] _Trace = new int[Trace.Length];

            for (int kk = 0; kk < Trace.Length; kk++)
            {
                _Trace[kk] = Convert.ToInt16(Trace[kk]);
            }

            Array.Sort(_Trace);

 //           DB_Interface.trans(Data_Interface);
            DB_Interface.DropTable(Data_Interface, "Trace_Info");
//
//            DB_Interface.Commit(Data_Interface);

            DB_Interface.trans(Data_Interface);
            DB_Interface.Make_table_For_Trace("Trace_Info", "", false);



            for (int kk = 0; kk < Trace.Length; kk++)
            {
                Dictionary<string, object[]> T = DB_Interface.Get_Data_By_Query_S4PD("", "CHAN" + _Trace[kk]);

                string S = "";

          
                foreach (KeyValuePair<string, object[]> _D in T)
                {
                    double[] Data = Array.ConvertAll<object, double>(_D.Value, Convert.ToDouble);

                    double Min = Data.Min();
                    double Max = Data.Max();

                 

                    if (Min != 0 && Max != 0)
                    {
                        S += _D.Key + ",";
                    }

                }

                DB_Interface.Make_table_For_Trace(S, "CHAN" + _Trace[kk], true);


            }

            DB_Interface.Commit(Data_Interface);
            //    double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

            ATE.SPARA_Form t = new ATE.SPARA_Form(DB_Interface, _Trace);
        }


    }
}
