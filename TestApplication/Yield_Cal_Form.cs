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
using System.Reflection;
using System.IO;

//using ADGV;


namespace TestApplication
{
    public partial class Yield_Cal_Form : Form
    {
        List<List<int>[]> List;
        List<List<int[]>>[] List2;
        Dictionary<string, List<int>> Dic;
        string Key;
        long Sample;
        long AnalysisSample;

        object[] Valuse;
        string JMP_File;
        public static string CSV_File_Path;


        List<string> ForGross_Fail_Unit;

        bool CellChange_Flag;
        int Data_Count;

        int GetString_length;

        Data_Class.Data_Editing Data_Edit;
        Data_Class.Data_Editing.INT Data_Interface;
        DB_Class.DB_Editing DB;
        DB_Class.DB_Editing.INT DB_Interface;

        JMP_Class.JMP_Editing.INT JMP_Interface;
        JMP_Class.JMP_Editing JMP = new JMP_Class.JMP_Editing();

        CSV_Class.CSV.INT Csv_Interface;
        CSV_Class.CSV CSV = new CSV_Class.CSV();

        PPTX_Class.PPTX_Editing.INT PPTX_Interface;
        PPTX_Class.PPTX_Editing PPTX = new PPTX_Class.PPTX_Editing();

        ToolStripStatusLabel filterStatusLabel = new ToolStripStatusLabel();
        ToolStripStatusLabel showAllLabel = new ToolStripStatusLabel("Show &All");



        DataTable dt = new DataTable();


        DataTable[] _dataTable;
        DataSet _dataSet = new DataSet();
        BindingSource[] bindingSource;


        Dictionary<string, List<int>[]> TestResult_Dic_For_New_Spec = new Dictionary<string, List<int>[]>();

        BindingSource bs;
        bool ForNewSpec;
        bool Already_Done_Anly;

        Dictionary<string, List<forctrlz>> ForCtrlz_Dic = new Dictionary<string, List<forctrlz>>();
        List<forctrlz>[] ForCtrlz_List;
        int ForCtrlz_List_count = 0;

        List<double[]> ForCtrlz_Min;
        List<double[]> ForCtrlz_Max;
        forctrlz Cz;

        DataGridViewCell clickedCell;

        ManualResetEvent[] ThreadFlags;
        bool[] Wait;
        string[] databylot;
        string[] databybin;

        int Coulumn_Count = 11;

        int Selected_Bin;

        string[] LOT;
        string[] SITE;
        string[] BIN;

        List<List<List<int>>[]>[] By_Lot;
        List<List<List<int>>[]>[] By_Site;
        List<List<List<int>>[]>[] By_Bin;

        Dictionary<string, int> Lot_Dic;
        Dictionary<string, int> Site_Dic;
        Dictionary<string, int> Bin_Dic;

        int[] Lot_Yield;
        int[] Site_Yield;
        int[] Bin_Yield;

        Dictionary<string, Bin_Struct>[] Bin_Infor;

        MakeSpec_Form MakeSpec;
        string[] No_Index;
        string[] Paraname;

        System.Windows.Forms.Button[] Analysis;
        System.Windows.Forms.Button[] Std;
        System.Windows.Forms.Button[] Sort;
        System.Windows.Forms.Button[] Save;
        System.Windows.Forms.Button MakeSpecB;

        System.Windows.Forms.CheckBox Check_Box;

        Zuby.ADGV.AdvancedDataGridView[] advancedDataGridView1;
        DataGridView[] Datagrid;
        DataGridView[] Datagrid2;

        int[] Calculate_thread_Strat;
        int[] Calculate_thread_End;
        int[] Sample_Verify;
        int[] Sample_Verify_Lot;

        List<int[]>[] List_Sample_Verify;


        ManualResetEvent[] For_Cal;

        Dictionary<string, Gross> Para;
        List<string> Spec_Number;


        bool Enabel = false;

        List<string> Outlier_List;
        static long None_Sample_Count;
        static long Hidden_Sample_Count;
        string[] Fail_Units;

        string[] Lot;

        int _OutCount;
        List<string> _Lot_Information_Dummy;
        Dictionary<string, List<string>> Lot_Information;
        Dictionary<string, Dictionary<string, List<string>>> Matching_Lots;

        OpenFileDialog Dialog = new OpenFileDialog();
        Dictionary<int, Dictionary<int, string>> Box_Enum = new Dictionary<int, Dictionary<int, string>>();
        Dictionary<string, CSV_Class.For_Box>[] Dic_Test;
        Dictionary<int, Dictionary<int, string>> OrderbySequence = new Dictionary<int, Dictionary<int, string>>();

        string[] N_Spec_Min;
        string[] N_Spec_Max;
        string[] C_Spec_Min;
        string[] C_Spec_Max;

        bool Customer_enable = false;
        bool NPI_enable = false;
        bool CPK_enable = false;
        double CPK_Value = 0f;

       

        public Yield_Cal_Form(List<List<int>[]> List, List<List<int[]>>[] List2, Dictionary<string, List<int>> Dic, string Key, Data_Class.Data_Editing Data_Edit, Data_Class.Data_Editing.INT Data, DB_Class.DB_Editing DB, DB_Class.DB_Editing.INT DB_Interface, int Sample_Count, int Hidden_Count)
        {


           Distribution_Form.Send += new Delete(m_ItemClicked);

            Outlier_List = new List<string>();
            JMP_Interface = JMP.Open(Key);


            //  DB_Interface.List_Gross_Values = new List<Dictionary<string, double[]>[]>();
            //  DB_Interface.Gross_Values1 = new Dictionary<string, double[]>[Data.DB_Count];

            //for (int i = 0; i < Data.DB_Count; i++)
            //{
            //    DB_Interface.Gross_Values1[i] = new Dictionary<string, double[]>();
            //}


            Dir.Dir_Directory Dir = new Dir.Dir_Directory("C:\\Automation\\Yield\\Add_option");
            Dir = new Dir.Dir_Directory("C:\\Automation\\Yield\\Gross_Check_option");
            Dir = new Dir.Dir_Directory("C:\\Temp\\Dummy");

            this.List = List;
            this.List2 = List2;
            this.Dic = Dic;
            this.Key = Key;

            this.Data_Edit = Data_Edit;
            this.Data_Interface = Data;
            this.DB = DB;
            this.DB_Interface = DB_Interface;

            this.Sample = Sample_Count;
            Key = "YIELD";
            InitializeComponent();

            checkBox2.Enabled = false;

            CellChange_Flag = false;


            //  dataGridView1.DoubleBuffered(true);
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);


            ForGross_Fail_Unit = new List<string>();

            Csv_Interface = CSV.Open(Key);

            //bool FileFlag = Dir.File_Exits("C:\\Automation\\Yield\\Add_option\\Option_Parameter.csv");

            //if (!FileFlag)
            //{
            //    Csv_Interface.Write_Open("C:\\Automation\\Yield\\Add_option\\Option_Parameter.csv");
            //    Csv_Interface.Write("Parameter", "Min", "Max");
            //    Csv_Interface.Write_Close();
            //}

            Para = new Dictionary<string, Gross>();
            Spec_Number = new List<string>();
            Gross dummy;


            foreach (string key in Dic.Keys)
            {
                string[] split = key.Split('_');

                if (split.Length > 2 && split[0].ToUpper() != "M")
                {
                    if (!Para.ContainsKey(split[1] + "_" + split[2]))
                    {
                        if (split[1].ToUpper().Contains("IBATT") || split[1].ToUpper().Contains("ICC") || split[1].ToUpper().Contains("IDD") || split[1].ToUpper().Contains("IEFF") || split[1].ToUpper().Contains("ITOTAL") || split[1].ToUpper().Contains("PCON"))
                        {
                            dummy = new Gross(1.20, "MAX/MIN");
                        }
                        else
                        {
                            dummy = new Gross(2, "MAX-MIN");
                        }

                        Para.Add(split[1] + "_" + split[2], dummy);
                    }
                }
            }



            foreach (string key in Dic.Keys)
            {
                string[] split = key.Split('_');

                if (split.Length > 2 && split[0].ToUpper() != "M")
                {
                    if (split[split.Length - 1].ToUpper() == "X")
                    {

                    }
                    else
                    {
                        if (!Spec_Number.Contains(split[split.Length - 1]))
                        {
                            Spec_Number.Add(split[split.Length - 1]);
                        }


                    }
                }
            }


            string[] Combo = new string[100];

            //foreach (KeyValuePair<string, double> o in Gross_Check)
            //{
            //    Combo[i] = o.Key.ToString();
            //    i++;
            //}
            //Array.Resize(ref Combo, i);

            //comboBox1.Items.AddRange(Combo);

            Combo = new string[1];

            Combo[0] = "LOT";
          //  Combo[1] = "SITE";
            //    Combo[2] = "BIN";

            comboBox2.Items.AddRange(Combo);

            Combo = new string[10];

            int ii = 0;
            foreach (KeyValuePair<string, Data_Class.Data_Editing.SWBIN> o in Data_Interface.SWBIN_Dic)
            {
                Combo[ii] = o.Key.ToString();
                ii++;
            }
            Array.Resize(ref Combo, ii);

            comboBox3.Items.AddRange(Combo);

            comboBox3.Text = "1";

            comboBox2.Enabled = false;
            comboBox3.Enabled = false;

            Button_Set();
            Std[0].Enabled = false;
            Std[1].Enabled = false;

            for (int k = 0; k < 1; k++)
            {
                DB_Interface.Table_Count += DB_Interface.Get_Sample_Count(0, "select count(*) from sqlite_master where name = 'Current_Setting'");
            }

            if (DB_Interface.Table_Count == 0)
            {
                Cal();
            }
            else
            {
                if (MessageBox.Show("Do you want to load Latest Setting?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {

                    if (DB_Interface.Table_Count != 0)
                    {
                        Std[0].Enabled = true;
                        Std[1].Enabled = true;


                        Lastest_Setting_Cal(true);
                        Matching_Lot_data();

                        DB_Interface.Matching_Lots = Matching_Lots;

                        for (int d = 0; d < Data_Interface.Clotho_Spcc_List[0].Max.Length; d++)
                        {
                            No_Index = new string[advancedDataGridView1[d].RowCount];
                            Paraname = new string[advancedDataGridView1[d].RowCount];

                            for (int k = 0; k < advancedDataGridView1[d].RowCount; k++)
                            {
                                No_Index[k] = advancedDataGridView1[d].Rows[k].Cells[0].Value.ToString();
                                Paraname[k] = advancedDataGridView1[d].Rows[k].Cells[1].Value.ToString();
                            }
                        }
                    }
                    else
                    {
                        Cal();
                    }

                }
                else
                {
                    Cal();
                }
            }
           





            Analysis[1].Enabled = false;

    
            dataGridView1.Visible = false;

            dataGridView2.Visible = false;
            Gridview2();

            radioButton2.Checked = true;
            //radioButton6.Checked = true;
            tabControl5.TabPages[0].Text = "Distribution";
          //  tabControl5.TabPages[1].Text = "Fit Y by X";
        }

        public void Cal()
        {

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            int i = 0;

            advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView[Data_Interface.Clotho_Spcc_List[0].Max.Length];
            _dataTable = new DataTable[Data_Interface.Clotho_Spcc_List[0].Max.Length];
            bindingSource = new BindingSource[Data_Interface.Clotho_Spcc_List[0].Max.Length];

            for (int f = 0; f < Data_Interface.Clotho_Spcc_List[0].Max.Length; f++)
            {
                string title = "Bin" + (f + 1);
                TabPage myTabPage = new TabPage(title);
                tabControl1.TabPages.Add(myTabPage);

                advancedDataGridView1[f] = new Zuby.ADGV.AdvancedDataGridView();
                _dataTable[f] = new DataTable();
                bindingSource[f] = new BindingSource();
                System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
                dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
                advancedDataGridView1[f].VirtualMode = true;
                advancedDataGridView1[f].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                //  advancedDataGridView1[f].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                advancedDataGridView1[f].Anchor = (AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom);
                advancedDataGridView1[f].AllowUserToAddRows = false;
                advancedDataGridView1[f].AllowUserToDeleteRows = false;
                advancedDataGridView1[f].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                advancedDataGridView1[f].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                advancedDataGridView1[f].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

                //  advancedDataGridView1[f].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                advancedDataGridView1[f].FilterAndSortEnabled = true;
                advancedDataGridView1[f].Location = new System.Drawing.Point(10, 10);
                advancedDataGridView1[f].Name = "advancedDataGridView1";
                advancedDataGridView1[f].RowHeadersVisible = false;
                //   advancedDataGridView1[f].RowTemplate.Height = 40;
                //    advancedDataGridView1[f].Size = new System.Drawing.Size(2854, 1650);
                advancedDataGridView1[f].TabIndex = 19;
                advancedDataGridView1[f].SortStringChanged += new System.EventHandler(this.advancedDataGridView1_SortStringChanged);
                advancedDataGridView1[f].FilterStringChanged += new System.EventHandler(this.advancedDataGridView1_FilterStringChanged);
                advancedDataGridView1[f].CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellDoubleClick);
                advancedDataGridView1[f].CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.advancedDataGridView1_CellMouseUp);
                advancedDataGridView1[f].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellValueChanged);
                advancedDataGridView1[f].KeyDown += new System.Windows.Forms.KeyEventHandler(this.advancedDataGridView1_KeyDown);
                advancedDataGridView1[f].Dock = System.Windows.Forms.DockStyle.Fill;

                bindingSource[f].DataSource = _dataTable[f];
                advancedDataGridView1[f].DataSource = bindingSource[f];

                advancedDataGridView1[f].DoubleBuffereds(true);



                //_dataTable = _dataSet.Tables.Add("TableTest");
                DataColumn[] dtkey = new DataColumn[1];

                _dataTable[f].Columns.Add("No", typeof(int));
                dtkey[0] = _dataTable[f].Columns["No"];
                _dataTable[f].PrimaryKey = dtkey;

                _dataTable[f].Columns.Add("Parameter");
                _dataTable[f].Columns.Add("S_Min", typeof(double));
                _dataTable[f].Columns.Add("S_Max", typeof(double));
                _dataTable[f].Columns.Add("D_Min", typeof(double));
                _dataTable[f].Columns.Add("Median", typeof(double));
                _dataTable[f].Columns.Add("D_Max", typeof(double));
                _dataTable[f].Columns.Add("CPK", typeof(double));
                //   _dataTable.Columns.Add("H_CPK", typeof(double));
                _dataTable[f].Columns.Add("Std", typeof(double));
                _dataTable[f].Columns.Add("%", typeof(double));
                _dataTable[f].Columns.Add("Fail", typeof(int));
                // _dataTable[f].Columns.Add("");
                //   _dataTable.Columns.Add("N_CPL", typeof(double));
                //   _dataTable.Columns.Add("N_CPH", typeof(double));
                //   _dataTable.Columns.Add("");
                // _dataTable[f].Columns.Add("N_Min", typeof(double));
                // _dataTable[f].Columns.Add("N_Max", typeof(double));

                bindingSource[f].DataMember = _dataTable[f].TableName;

                // Cal_Yield(Sample, out test);

                double[] High = Data_Interface.Ref_New_HighSpec;
                double[] Low = Data_Interface.Ref_New_LowSpec;

                double Testtime = TestTime1.Elapsed.TotalMilliseconds;
                Valuse = new object[Coulumn_Count];

                foreach (var item in Dic)
                {
                    if (i != 0)
                    {
                        Valuse[0] = Convert.ToString(i - 1);
                        Valuse[1] = item.Key.ToString();
                        Valuse[2] = null;
                        Valuse[3] = null;
                        Valuse[4] = null;
                        Valuse[5] = null;
                        Valuse[6] = null;
                        Valuse[7] = null;
                        Valuse[8] = null;


                        Valuse[9] = null;
                        Valuse[10] = item.Value[0].ToString();

                        Stopwatch TestTime2 = new Stopwatch();
                        TestTime2.Restart();
                        TestTime2.Start();
                        _dataTable[f].Rows.Add(Valuse);

                        double Testtime2 = TestTime2.Elapsed.TotalMilliseconds;
                    }
                    i++;

                }

                i = 0;
                tabControl1.TabPages[f].Controls.Add(advancedDataGridView1[f]);
            }


            this.Show();

            for (int f = 0; f < Data_Interface.Clotho_Spcc_List[0].Max.Length; f++)
            {
                advancedDataGridView1[f].Columns[0].ReadOnly = true;
                advancedDataGridView1[f].Columns[1].ReadOnly = true;
                advancedDataGridView1[f].Columns[4].ReadOnly = true;
                advancedDataGridView1[f].Columns[5].ReadOnly = true;
                advancedDataGridView1[f].Columns[6].ReadOnly = true;
                advancedDataGridView1[f].Columns[7].ReadOnly = true;
                advancedDataGridView1[f].Columns[8].ReadOnly = true;
                advancedDataGridView1[f].Columns[9].ReadOnly = true;
                advancedDataGridView1[f].Columns[10].ReadOnly = true;

            }

            //for (int r = 0; r < Data_Interface.Clotho_Spcc_List[0].Max.Length; r++)
            //{
            //    //advancedDataGridView1[r].Columns[0].Width = 40;
            //    //advancedDataGridView1[r].Columns[1].Width = 500;
            //    //advancedDataGridView1[r].Columns[2].Width = 60;
            //    //advancedDataGridView1[r].Columns[3].Width = 60;
            //    //advancedDataGridView1[r].Columns[4].Width = 60;
            //    //advancedDataGridView1[r].Columns[5].Width = 60;
            //    //advancedDataGridView1[r].Columns[6].Width = 60;
            //    //advancedDataGridView1[r].Columns[7].Width = 60;
            //    //advancedDataGridView1[r].Columns[8].Width = 60;
            //    //advancedDataGridView1[r].Columns[9].Width = 60;
            //    //advancedDataGridView1[r].Columns[10].Width = 40;

            //}
            ForeColor();
            Datagrid = new DataGridView[Data_Interface.Clotho_Spcc_List[0].Max.Length];

            for (int r = 0; r < Data_Interface.Clotho_Spcc_List[0].Max.Length; r++)
            {
                string title = "Bin" + (r + 1);
                TabPage myTabPage = new TabPage(title);
                tabControl3.TabPages.Add(myTabPage);

                Datagrid[r] = new DataGridView();
                System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                Datagrid[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                Datagrid[r].AllowUserToAddRows = false;
                Datagrid[r].AllowUserToDeleteRows = false;
                Datagrid[r].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                Datagrid[r].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                Datagrid[r].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                Datagrid[r].Location = new System.Drawing.Point(10, 10);
                Datagrid[r].Name = "advancedDataGridView1";
                Datagrid[r].RowHeadersVisible = false;
                Datagrid[r].RowTemplate.Height = 40;
                //  Datagrid[r].Size = new System.Drawing.Size(2854, 1650);
                Datagrid[r].TabIndex = 19;
                Datagrid[r].Dock = System.Windows.Forms.DockStyle.Fill;
                Datagrid[r].ReadOnly = true;
                Datagrid[r].RowHeadersVisible = false;
                Datagrid[r].ColumnCount = 2;
                Datagrid[r].BackgroundColor = Color.White;


                for (i = 0; i < 6; i++)
                {
                    Valuse = new object[2];

                    switch (i)
                    {
                        case 0:
                            Valuse[0] = "Total Sample";
                            Valuse[1] = this.Sample;
                            break;
                        case 1:
                            Valuse[0] = "Analysis Sample";
                            Valuse[1] = this.Sample - Hidden_Sample_Count;
                            break;
                        case 2:
                            Valuse[0] = "Pass";
                            Valuse[1] = 0;
                            break;
                        case 3:
                            Valuse[0] = "Fail";
                            Valuse[1] = 0;
                            break;
                        case 4:
                            Valuse[0] = "Percent";
                            Valuse[1] = 0;
                            break;
                        case 5:
                            Valuse[0] = "Hidden";
                            Valuse[1] = Hidden_Sample_Count;
                            break;
                    }



                    Datagrid[r].Rows.Add(Valuse);

                }

                Datagrid[r].Columns[0].Width = 100;
                Datagrid[r].Columns[1].Width = 40;

                tabControl3.TabPages[r].Controls.Add(Datagrid[r]);
            }



            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].Name = "Parameter";
            dataGridView1.Columns[1].Name = "Range";
            dataGridView1.Columns[2].Name = "Selector";
            var sort = Para.Keys.ToList();
            sort.Sort();


            foreach (string item in sort)
            {
                Valuse = new object[3];

                Valuse[0] = item;
                Valuse[1] = Para[item].Range;
                Valuse[2] = Para[item].Selector;

                dataGridView1.Rows.Add(Valuse);
            }




            dataGridView1.Columns[0].Width = 100;
            dataGridView1.Columns[1].Width = 40;
            dataGridView1.Columns[2].Width = 60;

            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;


            var dic = Para.OrderBy(num => num.Key);

            Datagrid2 = new DataGridView[2];

            for (int r = 0; r < 2; r++)
            {
                if (r == 0)
                {
                    string title = "Split Paraname";
                    TabPage myTabPage = new TabPage(title);
                    tabControl4.TabPages.Add(myTabPage);

                    Datagrid2[r] = new DataGridView();
                    System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                    Datagrid2[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                    Datagrid2[r].AllowUserToAddRows = false;
                    Datagrid2[r].AllowUserToDeleteRows = false;
                    Datagrid2[r].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                    Datagrid2[r].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                    Datagrid2[r].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Datagrid2[r].Location = new System.Drawing.Point(10, 10);
                    Datagrid2[r].Name = "advancedDataGridView1";
                    Datagrid2[r].RowHeadersVisible = false;
                    Datagrid2[r].RowTemplate.Height = 40;
                    Datagrid2[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    //      Datagrid2[r].Size = new System.Drawing.Size(2854, 1650);
                    Datagrid2[r].TabIndex = 19;
                    Datagrid2[r].Dock = System.Windows.Forms.DockStyle.Fill;

                    Datagrid2[r].RowHeadersVisible = false;
                    Datagrid2[r].ColumnCount = 2;
                    Datagrid2[r].BackgroundColor = Color.White;

                    Datagrid2[r].CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);


                    Valuse = new object[2];

                    DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();

                    Datagrid2[r].Columns.Add(buttonColumn);

                    buttonColumn.HeaderText = "Check";

                    //buttonColumn = new DataGridViewButtonColumn();

                    //Datagrid2[r].Columns.Add(buttonColumn);

                    //buttonColumn.HeaderText = "Box Plot";

                    foreach (KeyValuePair<string, Gross> d in dic)
                    {
                        Valuse[0] = d.Key;
                        Valuse[1] = "";
                        Datagrid2[r].Rows.Add(Valuse);
                    }

                    Datagrid2[r].Columns[0].Name = "Parameter";
                    Datagrid2[r].Columns[1].Name = "Option";

                    Datagrid2[r].Columns[0].Width = 100;
                    Datagrid2[r].Columns[1].Width = 150;

                    Datagrid2[r].Columns[0].ReadOnly = true;
                    Datagrid2[r].Columns[1].ReadOnly = false;


                    tabControl4.TabPages[r].Controls.Add(Datagrid2[r]);
                }
                else
                {
                    string title = "Spec Number";
                    TabPage myTabPage = new TabPage(title);
                    tabControl4.TabPages.Add(myTabPage);

                    Datagrid2[r] = new DataGridView();
                    System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                    Datagrid2[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                    Datagrid2[r].AllowUserToAddRows = false;
                    Datagrid2[r].AllowUserToDeleteRows = false;
                    Datagrid2[r].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                    Datagrid2[r].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                    Datagrid2[r].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Datagrid2[r].Location = new System.Drawing.Point(10, 10);
                    Datagrid2[r].Name = "advancedDataGridView1";
                    Datagrid2[r].RowHeadersVisible = false;
                    Datagrid2[r].RowTemplate.Height = 40;
                    //       Datagrid2[r].Size = new System.Drawing.Size(2854, 1650);
                    Datagrid2[r].TabIndex = 19;
                    Datagrid2[r].Dock = System.Windows.Forms.DockStyle.Fill;
                    //  Datagrid2[r].ReadOnly = true;
                    Datagrid2[r].RowHeadersVisible = false;
                    Datagrid2[r].ColumnCount = 1;
                    Datagrid2[r].BackgroundColor = Color.White;
                    Datagrid2[r].CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);

                    DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();

                    Datagrid2[r].Columns.Add(buttonColumn);

                    buttonColumn.HeaderText = "Check";

                    //buttonColumn = new DataGridViewButtonColumn();

                    //Datagrid2[r].Columns.Add(buttonColumn);

                    //buttonColumn.HeaderText = "Box Plot";

                    Valuse = new object[1];
                    foreach (string L in Spec_Number)
                    {
                        Valuse[0] = L;
                        Datagrid2[r].Rows.Add(Valuse);
                    }

                    Datagrid2[r].Columns[0].Name = "Spec Number";
                    Datagrid2[r].Columns[0].Width = 100;

                    Datagrid2[r].Columns[0].ReadOnly = true;
                    Datagrid2[r].Columns[1].ReadOnly = false;

                    tabControl4.TabPages[r].Controls.Add(Datagrid2[r]);
                }
            }


            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }

        public void Lastest_Setting_Cal(bool Flag)
        {
            Std[0].Enabled = true;

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            Already_Done_Anly = true;
            string Filename = DB_Interface.Filename.Substring(DB_Interface.Filename.LastIndexOf("\\") + 1);

            int length = DB_Interface.Filename.Length;
            Filename = DB_Interface.Filename.Substring(0, length - Filename.Length);

            Csv_Interface.Read_Open(Filename + "Inf.csv");

            int Select_bin = 0;
            long _count = 0;
            int _BinCount = 0;
            string[] _Lot = new string[0];
            bool _Flat = false;
            bool _StartFlag = false;

            List<int[]> Fail = new List<int[]>();
            List<string[]> Bin_Inf = new List<string[]>();
            string[] Bin_Int_Array = new string[0];

            


            int index = 0;

            while (!Csv_Interface.StreamReader.EndOfStream)
            {
                string[] value = Csv_Interface.Read();

                if (value[0].ToUpper() == "SAMPLECOUNT")
                {
                    _count = Convert.ToInt64(value[1]);
                }
                else if (value[0].ToUpper() == "LOT")
                {
                    _Lot = new string[value.Length - 1];

                    for (int l = 1; l < value.Length; l++)
                    {
                        _Lot[l - 1] = value[l];
                    }

                }
                else if (value[0].ToUpper() == "TOTAL_SAMPLE")
                {
                   // if (index == 0) Select_bin = 0;
                    Bin_Int_Array[0] = value[1];

                }
                else if (value[0].ToUpper() == "ANALYSIS_SAMPLE")
                {
                    Bin_Int_Array[1] = value[1];

                }
                else if (value[0].ToUpper() == "PASS")
                {
                    Bin_Int_Array[2] = value[1];

                }
                else if (value[0].ToUpper() == "FAIL")
                {
                    Bin_Int_Array[3] = value[1];

                }
                else if (value[0].ToUpper() == "PERCENT")
                {
                    Bin_Int_Array[4] = value[1];

                }
                else if (value[0].ToUpper() == "HIDDEN")
                {
                    Bin_Int_Array[5] = value[1];
                    Bin_Inf.Add(Bin_Int_Array);
                }

                else if (value[0].ToUpper() == "BINCOUNT")
                {
                    _BinCount = Convert.ToInt16(value[1]);
                    Fail = new List<int[]>();

                    Bin_Int_Array = new string[6];


                    for (int b = 0; b < Data_Interface.New_Header.Length; b++)
                    {
                        int[] List_Int = new int[_BinCount];
                        Fail.Add(List_Int);
                    }

                    DB_Interface.Yield_Test = new List<List<DB_Class.DB_Editing.RowAndPass>[]>[Data_Interface.DB_Count];
                    DB_Interface.For_Any_Yield = new List<List<int>>[Data_Interface.DB_Count];

                    for (int _i = 0; _i < Data_Interface.DB_Count; _i++)
                    {

                        DB_Interface.Yield_Test[_i] = new List<List<DB_Class.DB_Editing.RowAndPass>[]>();
                        DB_Interface.For_Any_Yield[_i] = new List<List<int>>();
                    }

                    for (int _k = 0; _k < _count; _k++)
                    {

                        List<DB_Class.DB_Editing.RowAndPass>[] test = new List<DB_Class.DB_Editing.RowAndPass>[_BinCount];
  

                        for (int _i = 0; _i < _BinCount; _i++)
                        {
                            test[_i] = new List<DB_Class.DB_Editing.RowAndPass>();
   

                        }

                        for (int rowcount = 0; rowcount < Data_Interface.DB_Count; rowcount++)
                        {
                            DB_Interface.Yield_Test[rowcount].Add(test);
                        }
                    }


                    for (int ih = 0; ih < Data_Interface.DB_Count; ih++)
                    {
                        int k = 0;
                        for (k = 0; k < 1; k++)
                        {
                            List<int> dummy = new List<int>();

                            for (k = 0; k < this.Data_Interface.Clotho_Spcc_List[1].Max.Length; k++)
                            {
                                dummy = new List<int>();

                                for (int n = 0; n < Data_Interface.Per_DB_Column_Count[ih]; n++)
                                {
                                    dummy.Add(0);

                                }

                                DB_Interface.For_Any_Yield[ih].Add(dummy);
                            }
                        }
                    }

                    _Flat = true;
                    Select_bin = 1;
                }

                if (value[0].ToUpper().Contains("BIN:" + Select_bin) && _Flat)
                {
                    _StartFlag = true;
                    value = Csv_Interface.Read();
                }

                if (_StartFlag)
                {

                    while (!Csv_Interface.StreamReader.EndOfStream)
                    {
                        string _D = value[0];
                        string _DB_SN = "";

                        if (value.Length > 1)
                        {
                            _DB_SN = value[1];
                        }


                        for (int rowcount = 2; rowcount < value.Length - 1; rowcount++)
                        {
                            int Find_DB = 0;

                            for (int k = 0; k < Data_Interface.DB_Count; k++)
                            {
                           
                                    if (Convert.ToInt16(value[rowcount]) <= Data_Interface.Per_DB_Column_Count_End[k])
                                    {
                                        Find_DB = k;
                                        Int32 N = Convert.ToInt32(_D);

                                        DB_Class.DB_Editing.RowAndPass _T = new DB_Class.DB_Editing.RowAndPass(Convert.ToInt64(_DB_SN), Convert.ToInt16(value[rowcount]) - Data_Interface.DB_Column_Limit * k, 1);
                                        DB_Interface.Yield_Test[k][N][Select_bin - 1].Add(_T);

                                     //   Fail[Convert.ToInt16(value[rowcount]) - 9][Select_bin - 1]++;

                                        int a = Data_Interface.DB_Column_Limit * k - Convert.ToInt16(value[rowcount]);

                                        int b = Convert.ToInt16(value[rowcount]) - ((Data_Interface.DB_Column_Limit) * k);

                                        DB_Interface.For_Any_Yield[k][Select_bin - 1][b]++;
                                        break;
                                    }
                                
                            }

                        }

                        //int Find_DB = 0;
                        //if (i > Data_Interface.DB_Column_Limit - 10)
                        //{
                        //    for (int k = 0; k < Data_Interface.DB_Count; k++)
                        //    {
                        //        if (i <= Data_Interface.Per_DB_Column_Count_End[k] - 9)
                        //        {
                        //            Find_DB = k;
                        //            Flag = true;
                        //            break;
                        //        }
                        //    }
                        //}

                        value = Csv_Interface.Read();

                        if (value[0].ToUpper() == "TOTAL_SAMPLE")
                        {
                            // if (index == 0) Select_bin = 0;
                            Bin_Int_Array[0] = value[1];

                        }
                        else if (value[0].ToUpper() == "ANALYSIS_SAMPLE")
                        {
                            Bin_Int_Array[1] = value[1];

                        }
                        else if (value[0].ToUpper() == "PASS")
                        {
                            Bin_Int_Array[2] = value[1];

                        }
                        else if (value[0].ToUpper() == "FAIL")
                        {
                            Bin_Int_Array[3] = value[1];

                        }
                        else if (value[0].ToUpper() == "PERCENT")
                        {
                            Bin_Int_Array[4] = value[1];

                        }
                        else if (value[0].ToUpper() == "HIDDEN")
                        {
                            Bin_Int_Array[5] = value[1];
                            Bin_Inf.Add(Bin_Int_Array);
                        }

                        if (value[0].ToUpper().Contains("BIN"))
                        {
                            string[] split = value[0].Split(':');
                            Select_bin = Convert.ToInt16(split[1]);
                            value = Csv_Interface.Read();
                        }

                    }



                }

            }

            Csv_Interface.Read_Close();

            if (Flag)
            {
                int i = 0;

                advancedDataGridView1 = new Zuby.ADGV.AdvancedDataGridView[Data_Interface.Clotho_Spcc_List[0].Max.Length];
                _dataTable = new DataTable[Data_Interface.Clotho_Spcc_List[0].Max.Length];
                bindingSource = new BindingSource[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                for (int f = 0; f < Data_Interface.Clotho_Spcc_List[0].Max.Length; f++)
                {
                    string title = "Bin" + (f + 1);
                    TabPage myTabPage = new TabPage(title);
                    tabControl1.TabPages.Add(myTabPage);

                    advancedDataGridView1[f] = new Zuby.ADGV.AdvancedDataGridView();
                    _dataTable[f] = new DataTable();
                    bindingSource[f] = new BindingSource();
                    System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
                    dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
                    advancedDataGridView1[f].VirtualMode = true;
                    advancedDataGridView1[f].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    //  advancedDataGridView1[f].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                    advancedDataGridView1[f].Anchor = (AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom);
                    advancedDataGridView1[f].AllowUserToAddRows = false;
                    advancedDataGridView1[f].AllowUserToDeleteRows = false;
                    advancedDataGridView1[f].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                    advancedDataGridView1[f].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                    advancedDataGridView1[f].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

                    //  advancedDataGridView1[f].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                    advancedDataGridView1[f].FilterAndSortEnabled = true;
                    advancedDataGridView1[f].Location = new System.Drawing.Point(10, 10);
                    advancedDataGridView1[f].Name = "advancedDataGridView1";
                    advancedDataGridView1[f].RowHeadersVisible = false;
                    //   advancedDataGridView1[f].RowTemplate.Height = 40;
                    //    advancedDataGridView1[f].Size = new System.Drawing.Size(2854, 1650);
                    advancedDataGridView1[f].TabIndex = 19;
                    advancedDataGridView1[f].SortStringChanged += new System.EventHandler(this.advancedDataGridView1_SortStringChanged);
                    advancedDataGridView1[f].FilterStringChanged += new System.EventHandler(this.advancedDataGridView1_FilterStringChanged);
                    advancedDataGridView1[f].CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellDoubleClick);
                    advancedDataGridView1[f].CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.advancedDataGridView1_CellMouseUp);
                    advancedDataGridView1[f].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellValueChanged);
                    advancedDataGridView1[f].KeyDown += new System.Windows.Forms.KeyEventHandler(this.advancedDataGridView1_KeyDown);
                    advancedDataGridView1[f].Dock = System.Windows.Forms.DockStyle.Fill;

                    bindingSource[f].DataSource = _dataTable[f];
                    advancedDataGridView1[f].DataSource = bindingSource[f];

                    advancedDataGridView1[f].DoubleBuffereds(true);



                    //_dataTable = _dataSet.Tables.Add("TableTest");
                    DataColumn[] dtkey = new DataColumn[1];

                    _dataTable[f].Columns.Add("No", typeof(int));
                    dtkey[0] = _dataTable[f].Columns["No"];
                    _dataTable[f].PrimaryKey = dtkey;

                    _dataTable[f].Columns.Add("Parameter");
                    _dataTable[f].Columns.Add("S_Min", typeof(double));
                    _dataTable[f].Columns.Add("S_Max", typeof(double));
                    _dataTable[f].Columns.Add("D_Min", typeof(double));
                    _dataTable[f].Columns.Add("Median", typeof(double));
                    _dataTable[f].Columns.Add("D_Max", typeof(double));
                    _dataTable[f].Columns.Add("CPK", typeof(double));
                    //   _dataTable.Columns.Add("H_CPK", typeof(double));
                    _dataTable[f].Columns.Add("Std", typeof(double));
                    _dataTable[f].Columns.Add("%", typeof(double));
                    _dataTable[f].Columns.Add("Fail", typeof(int));
                    // _dataTable[f].Columns.Add("");
                    //   _dataTable.Columns.Add("N_CPL", typeof(double));
                    //   _dataTable.Columns.Add("N_CPH", typeof(double));
                    //   _dataTable.Columns.Add("");
                    // _dataTable[f].Columns.Add("N_Min", typeof(double));
                    // _dataTable[f].Columns.Add("N_Max", typeof(double));

                    bindingSource[f].DataMember = _dataTable[f].TableName;

                    // Cal_Yield(Sample, out test);




                    DB_Interface.No_Index = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.Paraname = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.SpecMin = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.SpecMax = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.DataMin = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.DataMedian = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.DataMax = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.CPK = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.STD = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.Percent = new string[Data_Interface.Clotho_Spcc_List.Count];
                    DB_Interface.Fail = new string[Data_Interface.Clotho_Spcc_List.Count];



                    DB_Interface.Get_Current_Setting(Data_Interface, f);



                    double[] High = Data_Interface.Ref_New_HighSpec;
                    double[] Low = Data_Interface.Ref_New_LowSpec;

                    double Testtime = TestTime1.Elapsed.TotalMilliseconds;
                    Valuse = new object[Coulumn_Count];

                    foreach (var item in Dic)
                    {
                        if (i != 0)
                        {
                            if(DB_Interface.No_Index[i] == null)
                            {
                                Valuse[0] = i - 1;
                            }
                            else
                            {
                                Valuse[0] = DB_Interface.No_Index[i];
                            }
                      
                            Valuse[1] = DB_Interface.Paraname[i];
                            Valuse[2] = DB_Interface.SpecMin[i];
                            Valuse[3] = DB_Interface.SpecMax[i];
                            Valuse[4] = DB_Interface.DataMin[i];
                            Valuse[5] = DB_Interface.DataMedian[i];
                            Valuse[6] = DB_Interface.DataMax[i];
                            Valuse[7] = DB_Interface.CPK[i];
                            Valuse[8] = DB_Interface.STD[i];
                            Valuse[9] = DB_Interface.Percent[i];
                            Valuse[10] = Fail[i][f];

                            Stopwatch TestTime2 = new Stopwatch();
                            TestTime2.Restart();
                            TestTime2.Start();
                            _dataTable[f].Rows.Add(Valuse);

                            double Testtime2 = TestTime2.Elapsed.TotalMilliseconds;
                        }
                        i++;

                    }

                    i = 0;
                    tabControl1.TabPages[f].Controls.Add(advancedDataGridView1[f]);
                }


                this.Show();

                for (int f = 0; f < Data_Interface.Clotho_Spcc_List[0].Max.Length; f++)
                {
                    advancedDataGridView1[f].Columns[0].ReadOnly = true;
                    advancedDataGridView1[f].Columns[1].ReadOnly = true;
                    advancedDataGridView1[f].Columns[4].ReadOnly = true;
                    advancedDataGridView1[f].Columns[5].ReadOnly = true;
                    advancedDataGridView1[f].Columns[6].ReadOnly = true;
                    advancedDataGridView1[f].Columns[7].ReadOnly = true;
                    advancedDataGridView1[f].Columns[8].ReadOnly = true;
                    advancedDataGridView1[f].Columns[9].ReadOnly = true;
                    advancedDataGridView1[f].Columns[10].ReadOnly = true;

                    //advancedDataGridView1[f].Columns[0].Width = 40;
                    //advancedDataGridView1[f].Columns[1].Width = 500;
                    //advancedDataGridView1[f].Columns[2].Width = 60;
                    //advancedDataGridView1[f].Columns[3].Width = 60;
                    //advancedDataGridView1[f].Columns[4].Width = 60;
                    //advancedDataGridView1[f].Columns[5].Width = 60;
                    //advancedDataGridView1[f].Columns[6].Width = 60;
                    //advancedDataGridView1[f].Columns[7].Width = 60;
                    //advancedDataGridView1[f].Columns[8].Width = 60;
                    //advancedDataGridView1[f].Columns[9].Width = 60;
                    //advancedDataGridView1[f].Columns[10].Width = 40;

                }


                ForeColor();
                Datagrid = new DataGridView[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                for (int r = 0; r < Data_Interface.Clotho_Spcc_List[0].Max.Length; r++)
                {
                    string title = "Bin" + (r + 1);
                    TabPage myTabPage = new TabPage(title);
                    tabControl3.TabPages.Add(myTabPage);

                    Datagrid[r] = new DataGridView();
                    System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                    Datagrid[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                    Datagrid[r].AllowUserToAddRows = false;
                    Datagrid[r].AllowUserToDeleteRows = false;
                    Datagrid[r].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                    Datagrid[r].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                    Datagrid[r].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    Datagrid[r].Location = new System.Drawing.Point(10, 10);
                    Datagrid[r].Name = "advancedDataGridView1";
                    Datagrid[r].RowHeadersVisible = false;
                    Datagrid[r].RowTemplate.Height = 40;
                    //  Datagrid[r].Size = new System.Drawing.Size(2854, 1650);
                    Datagrid[r].TabIndex = 19;
                    Datagrid[r].Dock = System.Windows.Forms.DockStyle.Fill;
                    Datagrid[r].ReadOnly = true;
                    Datagrid[r].RowHeadersVisible = false;
                    Datagrid[r].ColumnCount = 2;
                    Datagrid[r].BackgroundColor = Color.White;

                    Bin_Int_Array = Bin_Inf[r];
                    for (i = 0; i < 6; i++)
                    {
                        Valuse = new object[2];

                        switch (i)
                        {
                            case 0:
                                Valuse[0] = "Total Sample";
                                Valuse[1] = Bin_Int_Array[0];
                                break;
                            case 1:
                                Valuse[0] = "Analysis Sample";
                                Valuse[1] = Bin_Int_Array[1];
                                break;
                            case 2:
                                Valuse[0] = "Pass";
                                Valuse[1] = Bin_Int_Array[2];
                                break;
                            case 3:
                                Valuse[0] = "Fail";
                                Valuse[1] = Bin_Int_Array[3];
                                break;
                            case 4:
                                Valuse[0] = "Percent";
                                Valuse[1] = Bin_Int_Array[4];
                                break;
                            case 5:
                                Valuse[0] = "Hidden";
                                Valuse[1] = Bin_Int_Array[5];
                                break;
                        }



                        Datagrid[r].Rows.Add(Valuse);

                    }

                    Datagrid[r].Columns[0].Width = 100;
                    Datagrid[r].Columns[1].Width = 40;

                    tabControl3.TabPages[r].Controls.Add(Datagrid[r]);
                }



                dataGridView1.ColumnCount = 3;
                dataGridView1.Columns[0].Name = "Parameter";
                dataGridView1.Columns[1].Name = "Range";
                dataGridView1.Columns[2].Name = "Selector";
                var sort = Para.Keys.ToList();
                sort.Sort();


                foreach (string item in sort)
                {
                    Valuse = new object[3];

                    Valuse[0] = item;
                    Valuse[1] = Para[item].Range;
                    Valuse[2] = Para[item].Selector;

                    dataGridView1.Rows.Add(Valuse);
                }




                dataGridView1.Columns[0].Width = 100;
                dataGridView1.Columns[1].Width = 40;
                dataGridView1.Columns[2].Width = 60;

                dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;


                var dic = Para.OrderBy(num => num.Key);

                Datagrid2 = new DataGridView[2];

                for (int r = 0; r < 2; r++)
                {
                    if (r == 0)
                    {
                        string title = "Split Paraname";
                        TabPage myTabPage = new TabPage(title);
                        tabControl4.TabPages.Add(myTabPage);

                        Datagrid2[r] = new DataGridView();
                        System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                        Datagrid2[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                        Datagrid2[r].AllowUserToAddRows = false;
                        Datagrid2[r].AllowUserToDeleteRows = false;
                        Datagrid2[r].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                        Datagrid2[r].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                        Datagrid2[r].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                        Datagrid2[r].Location = new System.Drawing.Point(10, 10);
                        Datagrid2[r].Name = "advancedDataGridView1";
                        Datagrid2[r].RowHeadersVisible = false;
                        Datagrid2[r].RowTemplate.Height = 40;
                        Datagrid2[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        //      Datagrid2[r].Size = new System.Drawing.Size(2854, 1650);
                        Datagrid2[r].TabIndex = 19;
                        Datagrid2[r].Dock = System.Windows.Forms.DockStyle.Fill;

                        Datagrid2[r].RowHeadersVisible = false;
                        Datagrid2[r].ColumnCount = 2;
                        Datagrid2[r].BackgroundColor = Color.White;

                        Datagrid2[r].CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);


                        Valuse = new object[2];

                        DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();

                        Datagrid2[r].Columns.Add(buttonColumn);

                        buttonColumn.HeaderText = "Check";

                        //buttonColumn = new DataGridViewButtonColumn();

                        //Datagrid2[r].Columns.Add(buttonColumn);

                        //buttonColumn.HeaderText = "Box Plot";

                        foreach (KeyValuePair<string, Gross> d in dic)
                        {
                            Valuse[0] = d.Key;
                            Valuse[1] = "";
                            Datagrid2[r].Rows.Add(Valuse);
                        }

                        Datagrid2[r].Columns[0].Name = "Parameter";
                        Datagrid2[r].Columns[1].Name = "Option";

                        Datagrid2[r].Columns[0].Width = 100;
                        Datagrid2[r].Columns[1].Width = 100;

                        Datagrid2[r].Columns[0].ReadOnly = true;
                        Datagrid2[r].Columns[1].ReadOnly = false;
                    

                        tabControl4.TabPages[r].Controls.Add(Datagrid2[r]);
                    }
                    else
                    {
                        string title = "Spec Number";
                        TabPage myTabPage = new TabPage(title);
                        tabControl4.TabPages.Add(myTabPage);

                        Datagrid2[r] = new DataGridView();
                        System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                        Datagrid2[r].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                        Datagrid2[r].AllowUserToAddRows = false;
                        Datagrid2[r].AllowUserToDeleteRows = false;
                        Datagrid2[r].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                        Datagrid2[r].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                        Datagrid2[r].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                        Datagrid2[r].Location = new System.Drawing.Point(10, 10);
                        Datagrid2[r].Name = "advancedDataGridView1";
                        Datagrid2[r].RowHeadersVisible = false;
                        Datagrid2[r].RowTemplate.Height = 40;
                        //       Datagrid2[r].Size = new System.Drawing.Size(2854, 1650);
                        Datagrid2[r].TabIndex = 19;
                        Datagrid2[r].Dock = System.Windows.Forms.DockStyle.Fill;
                        //  Datagrid2[r].ReadOnly = true;
                        Datagrid2[r].RowHeadersVisible = false;
                        Datagrid2[r].ColumnCount = 1;
                        Datagrid2[r].BackgroundColor = Color.White;
                        Datagrid2[r].CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);

                        DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();

                        Datagrid2[r].Columns.Add(buttonColumn);

                        buttonColumn.HeaderText = "Check";

                        //buttonColumn = new DataGridViewButtonColumn();

                        //Datagrid2[r].Columns.Add(buttonColumn);

                        //buttonColumn.HeaderText = "Box Plot";

                        Valuse = new object[1];
                        foreach (string L in Spec_Number)
                        {
                            Valuse[0] = L;
                            Datagrid2[r].Rows.Add(Valuse);
                        }

                        Datagrid2[r].Columns[0].Name = "Spec Number";
                        Datagrid2[r].Columns[0].Width = 100;

                        Datagrid2[r].Columns[0].ReadOnly = true;
                        Datagrid2[r].Columns[1].ReadOnly = false;

                        tabControl4.TabPages[r].Controls.Add(Datagrid2[r]);
                    }
                }

            }
            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }

        public void Re_GridView()
        {
            this.Enabled = false;

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            Csv_Interface = CSV.Open(Key);
            //Csv_Interface.Read_Open2("C:\\Automation\\Yield\\Add_option\\Option_Parameter.csv");

            //Dictionary<string, string[]> Spec_Option = new Dictionary<string, string[]>();
            //int p = 0;

            //while (!Csv_Interface.StreamReader2.EndOfStream)
            //{
            //    string[] MinAndMax = new string[2];
            //    string[] GetData = Csv_Interface.Read2();
            //    if (p != 0)
            //    {
            //        MinAndMax[0] = GetData[1];
            //        MinAndMax[1] = GetData[2];
            //        Spec_Option.Add(GetData[0], MinAndMax);
            //    }
            //    p++;
            //}

            //Csv_Interface.Read2_Close();

            for (int d = 0; d < Data_Interface.Clotho_Spcc_List[0].Max.Length; d++)
            {
                No_Index = new string[advancedDataGridView1[d].RowCount];
                Paraname = new string[advancedDataGridView1[d].RowCount];

                for (int k = 0; k < advancedDataGridView1[d].RowCount; k++)
                {
                    No_Index[k] = advancedDataGridView1[d].Rows[k].Cells[0].Value.ToString();
                    Paraname[k] = advancedDataGridView1[d].Rows[k].Cells[1].Value.ToString();
                }


                int j = 0;

                double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;

                _dataTable[d].Rows.Clear();

                #region
                int List_Count = 0;


                List_Count = DB_Interface.Yield_Test[0].Count;

                Bin_Infor = new Dictionary<string, Bin_Struct>[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                for (int k = 0; k < Bin_Infor.Length; k++)
                {
                    Bin_Infor[k] = new Dictionary<string, Bin_Struct>();
                }


                _dataTable[d].BeginLoadData();

                int Row_Index = 1;

                for (int w = 0; w < Paraname.Length; w++)
                {

                    Stopwatch TestTime2 = new Stopwatch();
                    TestTime2.Restart();
                    TestTime2.Start();

                    double Dummy = 0f;
                    Valuse = new object[Coulumn_Count];


                    int Db_Index = (Convert.ToInt16(No_Index[w]) + 10) / (Data_Interface.DB_Column_Limit);



                    int offset = 21;
                    int Test_Index = 0;

                    if (Db_Index != 0)
                    {
                        Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
                        offset = 10;
                    }
                    else
                    {
                        Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
                    }

                    if (Db_Index == 1)
                    {

                    }


                    int Test_Index2 = (Convert.ToInt16(No_Index[w]) + offset) - Test_Index;

                    if (Db_Index == 8 && Test_Index2 == 430)
                    {

                    }

                    int PerCount = DB_Interface.For_Any_Yield[Db_Index][d][Test_Index2];

               


                    Valuse[0] = Convert.ToInt16(No_Index[w]);
                    Valuse[1] = Paraname[w];
                    //   Clotho_Spcc_List

                    Valuse[2] = Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Min[d];
                    Valuse[3] = Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Max[d];

                    //Valuse[2] = Data_Class.Data_Editing.New_LowSpec[Convert.ToInt16(No_Index[j]) + 1];
                    //Valuse[3] = Data_Class.Data_Editing.New_HighSpec[Convert.ToInt16(No_Index[j]) + 1];

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Min_Data[d];
                    Valuse[4] = Dummy;

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Median_Data[d];
                    Valuse[5] = Dummy;

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Max_Data[d];
                    Valuse[6] = Dummy;

                    double L_CPK = 0f;
                    double H_CPK = 0f;
                    L_CPK = (DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Avg[d] - Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Min[d]) / (3 * DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Std[d]);
                    H_CPK = (Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Max[d] - DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Avg[d]) / (3 * DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Std[d]);

                    if (L_CPK > H_CPK) Dummy = H_CPK;
                    else Dummy = L_CPK;

                    Valuse[7] = Math.Round(Dummy, 3);

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Std[d];
                    Valuse[8] = Math.Round(Dummy, 7);

                    int Test = PerCount + List_Count;
                    int Test1 = List_Count - PerCount;


                    if (Test == List_Count)
                    {
                        try
                        {
                            Dummy = Convert.ToDouble(Test) / Convert.ToDouble(List_Count) * 100;
                        }

                        catch
                        {
                            Dummy = 0;
                        }
                        Valuse[9] = Dummy;
                    }
                    else
                    {
                        Dummy = Convert.ToDouble(Test1) / Convert.ToDouble(List_Count) * 100;
                        Valuse[9] = Dummy;

                    }

                    Valuse[10] = PerCount;
                    _dataTable[d].Rows.Add(Valuse);


                    j++;
                    double Testtime2 = TestTime2.Elapsed.TotalMilliseconds;
                    Row_Index++;
                }

                _dataTable[d].EndLoadData();

                bindingSource[d].DataSource = _dataTable[d];

                double Testtime10 = TestTime1.Elapsed.TotalMilliseconds;

                double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                #endregion

                ForeColor();

                double Testtime7 = TestTime1.Elapsed.TotalMilliseconds;

                Sample_Verify = new int[Data_Interface.Clotho_Spcc_List[0].Max.Length];

           
                //   Cal_Yield2(Sample - ForGross_Fail_Unit.Count);

                double Testtime8 = TestTime1.Elapsed.TotalMilliseconds;


            }

            Cal_No_Thread(Sample - Hidden_Sample_Count);
            Write_Inf(Sample - Hidden_Sample_Count);
            this.Enabled = true;
            double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;
        }

        public void Re_GridView2()
        {
            this.Enabled = false;

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            Csv_Interface = CSV.Open(Key);
            //Csv_Interface.Read_Open2("C:\\Automation\\Yield\\Add_option\\Option_Parameter.csv");

            //Dictionary<string, string[]> Spec_Option = new Dictionary<string, string[]>();
            //int p = 0;

            //while (!Csv_Interface.StreamReader2.EndOfStream)
            //{
            //    string[] MinAndMax = new string[2];
            //    string[] GetData = Csv_Interface.Read2();
            //    if (p != 0)
            //    {
            //        MinAndMax[0] = GetData[1];
            //        MinAndMax[1] = GetData[2];
            //        Spec_Option.Add(GetData[0], MinAndMax);
            //    }
            //    p++;
            //}

            //Csv_Interface.Read2_Close();

            for (int d = 0; d < Data_Interface.Clotho_Spcc_List[0].Max.Length; d++)
            {
                No_Index = new string[Dic.Count - 1];
                Paraname = new string[Dic.Count - 1];

                int jj = 0;
                foreach (var item in Dic)
                {
                    if (jj == 0)
                    {

                    }
                    else
                    {
                        No_Index[jj - 1] = (jj - 1).ToString();
                        Paraname[jj - 1] = item.Key;
                    }

                    jj++;
                }



                int j = 0;

                double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;

                _dataTable[d].Rows.Clear();

                #region
                int List_Count = 0;


                List_Count = DB_Interface.Yield_Test[0].Count;

                Bin_Infor = new Dictionary<string, Bin_Struct>[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                for (int k = 0; k < Bin_Infor.Length; k++)
                {
                    Bin_Infor[k] = new Dictionary<string, Bin_Struct>();
                }


                _dataTable[d].BeginLoadData();

                int Row_Index = 1;

                for (int w = 0; w < Paraname.Length; w++)
                {

                    Stopwatch TestTime2 = new Stopwatch();
                    TestTime2.Restart();
                    TestTime2.Start();

                    double Dummy = 0f;
                    Valuse = new object[Coulumn_Count];


                    int Db_Index = (Convert.ToInt16(No_Index[w]) + 10) / (Data_Interface.DB_Column_Limit);



                    int offset = 21;
                    int Test_Index = 0;

                    if (Db_Index != 0)
                    {
                        Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
                        offset = 10;
                    }
                    else
                    {
                        Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
                    }

                    if (Db_Index == 1)
                    {

                    }


                    int Test_Index2 = (Convert.ToInt16(No_Index[w]) + offset) - Test_Index;

                    int PerCount = DB_Interface.For_Any_Yield[Db_Index][d][Test_Index2];



                    Valuse[0] = Convert.ToInt16(No_Index[w]);
                    Valuse[1] = Paraname[w];
                    //   Clotho_Spcc_List

                    Valuse[2] = Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Min[d];
                    Valuse[3] = Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Max[d];

                    //Valuse[2] = Data_Class.Data_Editing.New_LowSpec[Convert.ToInt16(No_Index[j]) + 1];
                    //Valuse[3] = Data_Class.Data_Editing.New_HighSpec[Convert.ToInt16(No_Index[j]) + 1];

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Min_Data[d];
                    Valuse[4] = Dummy;

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Median_Data[d];
                    Valuse[5] = Dummy;

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Max_Data[d];
                    Valuse[6] = Dummy;

                    double L_CPK = 0f;
                    double H_CPK = 0f;
                    L_CPK = (DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Avg[d] - Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Min[d]) / (3 * DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Std[d]);
                    H_CPK = (Data_Interface.Clotho_Spcc_List[Convert.ToInt16(No_Index[j]) + 1].Max[d] - DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Avg[d]) / (3 * DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Std[d]);

                    if (L_CPK > H_CPK) Dummy = H_CPK;
                    else Dummy = L_CPK;

                    Valuse[7] = Math.Round(Dummy, 3);

                    Dummy = DB_Interface.Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[j]) + 1])].Std[d];
                    Valuse[8] = Math.Round(Dummy, 7);

                    int Test = PerCount + List_Count;
                    int Test1 = List_Count - PerCount;


                    if (Test == List_Count)
                    {
                        try
                        {
                            Dummy = Convert.ToDouble(Test) / Convert.ToDouble(List_Count) * 100;
                        }

                        catch
                        {
                            Dummy = 0;
                        }
                        Valuse[9] = Dummy;
                    }
                    else
                    {
                        Dummy = Convert.ToDouble(Test1) / Convert.ToDouble(List_Count) * 100;
                        Valuse[9] = Dummy;

                    }

                    Valuse[10] = PerCount;
                    _dataTable[d].Rows.Add(Valuse);


                    j++;
                    double Testtime2 = TestTime2.Elapsed.TotalMilliseconds;
                    Row_Index++;
                }

                _dataTable[d].EndLoadData();

                bindingSource[d].DataSource = _dataTable[d];

                double Testtime10 = TestTime1.Elapsed.TotalMilliseconds;

                double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                #endregion

                ForeColor();

                double Testtime7 = TestTime1.Elapsed.TotalMilliseconds;

                Sample_Verify = new int[Data_Interface.Clotho_Spcc_List[0].Max.Length];

          
                //   Cal_Yield2(Sample - ForGross_Fail_Unit.Count);

                double Testtime8 = TestTime1.Elapsed.TotalMilliseconds;

            }

            Cal_No_Thread(Sample - Hidden_Sample_Count);

            Write_Inf(Sample - Hidden_Sample_Count);
            this.Enabled = true;
            double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;
        }

        public void For_New_Spec_Re_GridView()
        {

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            #region
            int List_Count = 0;


            List_Count = DB_Interface.Yield_Test_New_Spec[0].Count;

            Bin_Infor = new Dictionary<string, Bin_Struct>[Data_Interface.Clotho_List[0].Max.Length];

            for (int k = 0; k < Bin_Infor.Length; k++)
            {
                Bin_Infor[k] = new Dictionary<string, Bin_Struct>();
            }

            for (int g = 0; g < MakeSpec.advanced.Length; g++)
            {

                for (int h = 0; h < Data_Interface.Reference_Header.Length - 1; h++)
                {
                    No_Index[h] = MakeSpec.advanced[g].Rows[h].Cells[0].Value.ToString();
                    Paraname[h] = MakeSpec.advanced[g].Rows[h].Cells[1].Value.ToString();
                }


                for (int w = 0; w < Paraname.Length; w++)
                {
                    DataColumn[] dtkey = new DataColumn[1];

                    dtkey[0] = MakeSpec._dataTable[g].Columns["No"];
                    MakeSpec._dataTable[g].PrimaryKey = dtkey;

                    DataRow dr = MakeSpec._dataTable[g].Rows.Find(No_Index[w]);
                    int SelRow = MakeSpec._dataTable[g].Rows.IndexOf(dr);

                    double Dummy = 0f;

                    int Db_Index = (Convert.ToInt16(No_Index[w]) + 10) / (Data_Interface.DB_Column_Limit);

                    int Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
                    int Test_Index2 = Convert.ToInt16(No_Index[w]) + 10 - Test_Index;



                    //var itemToRemove = Db_Interface.Yield_Test[Db_Index][i][RefTab].Find(r => r.Row == getNb);

                    int PerCount = DB_Interface.For_Any_Yield_For_New_Spec[Db_Index][g][Test_Index2];

                    int Index = Convert.ToInt16(No_Index[w]);

                    MakeSpec._dataTable[g].Rows[SelRow].BeginEdit();


                    Dummy = DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Min_Data[g];
                    MakeSpec._dataTable[g].Rows[SelRow][8] = Dummy;

                    Dummy = DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Median_Data[g];
                    MakeSpec._dataTable[g].Rows[SelRow][9] = Dummy;

                    Dummy = DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Max_Data[g];
                    MakeSpec._dataTable[g].Rows[SelRow][10] = Dummy;

                    double L_CPK = 0f;
                    double H_CPK = 0f;
                    L_CPK = (DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Avg[g] - Data_Interface.Customor_Clotho_List[w + 1].Min[g]) / (3 * DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Std[g]);
                    H_CPK = (Data_Interface.Customor_Clotho_List[w + 1].Max[g] - DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Avg[g]) / (3 * DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Std[g]);

                    if (L_CPK > H_CPK) Dummy = H_CPK;
                    else Dummy = L_CPK;


                    if (Double.IsInfinity(Dummy))
                    {
                        Dummy = 0;
                    }

                    MakeSpec._dataTable[g].Rows[SelRow][11] = Math.Round(Dummy, 3);

                    Dummy = DB_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].Std[g];
                    MakeSpec._dataTable[g].Rows[SelRow][12] = Math.Round(Dummy, 7);

                    int Test = PerCount + List_Count;
                    int Test1 = List_Count - PerCount;


                    if (Test == List_Count)
                    {
                        try
                        {
                            Dummy = Convert.ToDouble(Test) / Convert.ToDouble(List_Count) * 100;
                        }

                        catch
                        {
                            Dummy = 0;
                        }
                        MakeSpec._dataTable[g].Rows[SelRow][13] = Dummy;
                    }
                    else
                    {
                        Dummy = Convert.ToDouble(Test1) / Convert.ToDouble(List_Count) * 100;
                        MakeSpec._dataTable[g].Rows[SelRow][13] = Dummy;

                    }

                    MakeSpec._dataTable[g].Rows[SelRow][14] = PerCount;

                    //MakeSpec._dataTable[g].Rows[SelRow][15] = DB_Interface.DIC_IQR[Convert.ToString(Data_Class.Data_Editing.New_Header[Convert.ToInt16(No_Index[w]) + 1])].L_IQR;
                    //MakeSpec._dataTable[g].Rows[SelRow][16] = DB_Interface.DIC_IQR[Convert.ToString(Data_Class.Data_Editing.New_Header[Convert.ToInt16(No_Index[w]) + 1])].H_IQR;

                    string[] Cout = DB_Interface.DIC_IQR[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].SN;

                    if (DB_Interface.DIC_IQR[Convert.ToString(Data_Interface.Ref_New_Header[Convert.ToInt16(No_Index[w]) + 1])].SN == null) MakeSpec._dataTable[g].Rows[SelRow][15] = 0;
                    else MakeSpec._dataTable[g].Rows[SelRow][15] = Cout.Length;




                }



            }

            for (int g = 0; g < MakeSpec.advanced.Length; g++)
            {
                MakeSpec.bindingSource[g].DataSource = MakeSpec._dataTable[g];
                MakeSpec.advanced[g].Update();
            }

            double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
            #endregion

            ForeColor_New_Spec();

            Sample_Verify = new int[MakeSpec.advanced.Length];
            Cal_NewSpec_No_Thread(Sample - Hidden_Sample_Count);


            //  For_New_Spec_Cal_Yield2(Sample - ForGross_Fail_Unit.Count);


            this.Enabled = true;
        }
        private void Cal_Thread(Object index)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            int[] Bin_Arry = new int[DB_Interface.Yield_Test[0][0].Length];
            bool flag = false;
            int i = (int)index;
            int Db = 0;


            for (int n = Calculate_thread_Strat[i]; n < Calculate_thread_End[i]; n++)
            {
                for (int j = 0; j < DB_Interface.Yield_Test[Db][n].Length; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {
                        for (int m = 0; m < DB_Interface.Yield_Test[Db][n][j].Count; m++)
                        {
                            Bin_Arry[j]++;
                            Db = 0;
                            flag = true;
                            break;

                        }
                        if (flag) break;
                    }
                    flag = false;

                }
            }
            List_Sample_Verify[i].Add(Bin_Arry);
            For_Cal[i].Set();
            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }


        private void button2_Click(object sender, EventArgs e) ////////////////////// Std
        {
            //  advancedDataGridView1.CleanFilters();
            //  advancedDataGridView1.CleanSorts();
            SetSpec();

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            int index = tabControl2.SelectedIndex;

            if (index == 1 && MakeSpec.Enabled)
            {
                Outlier_List = new List<string>();
                if (DB_Interface.For_New_Spec_Cal_Value_by_rowsdata == null)
                {
                    double[] dummy_Test = new double[14];
                    DB_Interface.For_New_Spec_Cal_Value_by_rowsdata = new Dictionary<string, DB_Class.DB_Editing.Data_Calculation>();

                    for (int j = 0; j < Data_Interface.New_Header.Length; j++)
                    {
                        DB_Interface.For_New_Spec_Cal_Value_by_rowsdata.Add(Data_Interface.Ref_New_Header[j], new DB_Class.DB_Editing.Data_Calculation(Data_Interface.Clotho_List[0].Max.Length));
                    }
                }

                DB_Interface.Get_Ave_Data_For_New_Spec(Data_Interface);

                double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;


                for (int outlier = 0; outlier < DB_Interface.DIC_IQR.Count; outlier++)
                {
                    string[] dummy = DB_Interface.DIC_IQR[Data_Interface.Reference_Header[outlier]].SN;

                    for (int dummycount = 0; dummycount < dummy.Length; dummycount++)
                    {
                        if (!Outlier_List.Contains(dummy[dummycount]))
                        {
                            Outlier_List.Add(dummy[dummycount]);
                        }

                        //  Count_Test = Count.Concat(dummy).ToArray();
                    }

                }

                For_New_Spec_Re_GridView();


                string[] Count = new string[0];


            }
            else if (index == 0)
            {
                Csv_Interface = CSV.Open(Key);
                //Csv_Interface.Read_Open2("C:\\Automation\\Yield\\Add_option\\Option_Parameter.csv");

                //Dictionary<string, string[]> Spec_Option = new Dictionary<string, string[]>();
                //int p = 0;

                //while (!Csv_Interface.StreamReader2.EndOfStream)
                //{
                //    string[] MinAndMax = new string[2];
                //    string[] GetData = Csv_Interface.Read2();
                //    if (p != 0)
                //    {
                //        MinAndMax[0] = GetData[1];
                //        MinAndMax[1] = GetData[2];
                //        Spec_Option.Add(GetData[0], MinAndMax);
                //    }
                //    p++;
                //}

                //Csv_Interface.Read2_Close();

                double Testtime = TestTime1.Elapsed.TotalMilliseconds;

                DB_Interface.Get_Ave_Data(Data_Interface);

                // DB_Interface.Read_Dispose(Data_Interface);
                double Testtime10 = TestTime1.Elapsed.TotalMilliseconds;
                //
                //  DB_Interface.Set_Conn(Data_Interface);
                double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

                Re_GridView();

            }


            Already_Done_Anly = true;
            double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

        }
        private void button3_Click(object sender, EventArgs e) ////////////////////// Make a Spec
        {
            string Query = "";


            Sample = 0;
            AnalysisSample = 0;
            for (int loop = 0; loop < DB_Interface.Table_Count; loop++)
            {
                Query = "Select count(id) from data" + loop;
                string[] TotalSample = DB_Interface.Get_Data_By_Query(Query);
                Sample += Convert.ToInt64(TotalSample[0]);

            }

            for (int loop = 0; loop < DB_Interface.Table_Count; loop++)
            {
                Query = "Select count(id) from data" + loop + " where Fail not like '1'";
                string[] TotalSample = DB_Interface.Get_Data_By_Query(Query);
                AnalysisSample += Convert.ToInt64(TotalSample[0]);

            }

            for (int loop = 0; loop < DB_Interface.Table_Count; loop++)
            {
                Query = "Select count(id) from data" + loop + " where Fail not like '0'";
                string[] TotalSample = DB_Interface.Get_Data_By_Query(Query);
                Hidden_Sample_Count += Convert.ToInt64(TotalSample[0]);

            }


            if (!Enabel)
            {
                MakeSpec = new MakeSpec_Form(Data_Interface, DB_Interface, Csv_Interface, JMP_Interface, Sample, AnalysisSample, Hidden_Sample_Count, Outlier_List);
                Enabel = true;
            }
            else
            {
                MakeSpec.Show();
            }


            Analysis[1].Enabled = true;
            // MakeSpec.Show();
        }
        private void button4_Click(object sender, EventArgs e)  /////////////////////// Yield
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            string Query = "";

            Matching_Lot_data();

            Hidden_Sample_Count = 0;
            None_Sample_Count = 0;

            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                Dictionary<string, List<string>> tests = key.Value;


                foreach (KeyValuePair<string, List<string>> ts in tests)
                {
                    Query = "select count(parameter) from " + key.Key + "  where FAIL = 1";
                    Hidden_Sample_Count += DB_Interface.Get_Sample_Count(0, Query);

                    Query = "Select count(id) from " + key.Key + " where Fail = '0'";
                    None_Sample_Count += DB_Interface.Get_Sample_Count(0, Query);
                }

            }

            double Loop_Count = Convert.ToDouble(None_Sample_Count) / Convert.ToDouble(DB_Interface.Limit);
            double Temp = Math.Truncate(Loop_Count);

            if (Loop_Count > Temp) DB_Interface.Limit_Count = Convert.ToInt16(Temp) + 1;
            else DB_Interface.Limit_Count = Convert.ToInt16(Temp);


            int index = tabControl2.SelectedIndex;

            if (index == 1 && MakeSpec.Enabled)
            {
                for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
                {
                    MakeSpec._dataTable[s].DefaultView.Sort = "[No] ASC";
                    MakeSpec.bindingSource[s].Filter = "";
                }


                DB_Data_Yield_For_NewSpec();

                Std[1].Enabled = true;


            }
            else if (index == 0)
            {
                //advancedDataGridView1.CleanFilters();
                //advancedDataGridView1.CleanSorts();

                if (DB_Interface._From_Db)
                {
                    DB_Data_Yield();
                }
                else
                {
                    CSV_Data_Yield();
                }

                dataGridView1.Visible = false;
                Already_Done_Anly = true;
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
                Std[0].Enabled = true;
            }


            //  DB_Interface.Read_Dispose(Data_Interface);
            double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

            //    DB_Interface.Set_Conn(Data_Interface);
            double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DB_Interface.Close(Data_Interface);

            //DB_Interface.ForCampare_Yield = new List<List<int>[]>[7];
        }
        private void button10_Click(object sender, EventArgs e) /////////////////////// Sort
        {
            int index = tabControl1.SelectedIndex;

            Re_GridView2();

            advancedDataGridView1[index].CleanFilters();
            advancedDataGridView1[index].CleanSorts();

            _dataTable[index].DefaultView.RowFilter = "";
            _dataTable[index].DefaultView.Sort = "[No] ASC";
            bindingSource[index].DataSource = _dataTable[index];
            bindingSource[index].Filter = "";

        }
        private void button11_Click(object sender, EventArgs e)
        {

            //string Test = advancedDataGridView1.Rows[0].Cells[13].Value.ToString();
            //if (Test != "")
            //{
            //    string[] No_Index = new string[advancedDataGridView1.RowCount];

            //    for (int k = 0; k < advancedDataGridView1.RowCount; k++)
            //    {
            //        No_Index[k] = advancedDataGridView1.Rows[k].Cells[0].Value.ToString();
            //    }
            //    Data_Interface.New_HighSpec = new double[Data_Interface.New_HighSpec.Length];
            //    Data_Interface.New_LowSpec = new double[Data_Interface.New_HighSpec.Length];

            //    Data_Class.Data_Editing.ForAnl_NewMinSpec = new string[advancedDataGridView1.RowCount + Data_Interface.TheEnd_Trashes_Header_Count + Data_Interface.TheFirst_Trashes_Header_Count + 1];
            //    Data_Class.Data_Editing.ForAnl_NewMaxSpec = new string[advancedDataGridView1.RowCount + Data_Interface.TheEnd_Trashes_Header_Count + Data_Interface.TheFirst_Trashes_Header_Count + 1];

            //    Data_Class.Data_Editing.ForAnl_NewMinSpec[0] = "LOW";
            //    Data_Class.Data_Editing.ForAnl_NewMaxSpec[0] = "HIGH";

            //    for (int l = 1; l < Data_Interface.TheFirst_Trashes_Header_Count + 1; l++)
            //    {
            //        Data_Class.Data_Editing.ForAnl_NewMinSpec[l] = "0";
            //        Data_Class.Data_Editing.ForAnl_NewMaxSpec[l] = "0";
            //    }
            //    for (int l = Data_Interface.TheFirst_Trashes_Header_Count + 1; l < Data_Class.Data_Editing.ForAnl_NewMinSpec.Length - Data_Interface.TheEnd_Trashes_Header_Count; l++)
            //    {
            //        Data_Class.Data_Editing.ForAnl_NewMinSpec[Convert.ToInt16(No_Index[l - (Data_Interface.TheFirst_Trashes_Header_Count + 1)]) + Data_Interface.TheFirst_Trashes_Header_Count + 1] = Convert.ToString(advancedDataGridView1.Rows[l - (Data_Interface.TheFirst_Trashes_Header_Count + 1)].Cells[12].Value.ToString());
            //        Data_Class.Data_Editing.ForAnl_NewMaxSpec[Convert.ToInt16(No_Index[l - (Data_Interface.TheFirst_Trashes_Header_Count + 1)]) + Data_Interface.TheFirst_Trashes_Header_Count + 1] = Convert.ToString(advancedDataGridView1.Rows[l - (Data_Interface.TheFirst_Trashes_Header_Count + 1)].Cells[13].Value.ToString());
            //    }

            //    for (int l = Data_Class.Data_Editing.ForAnl_NewMinSpec.Length - Data_Interface.TheEnd_Trashes_Header_Count; l < Data_Class.Data_Editing.ForAnl_NewMaxSpec.Length; l++)
            //    {
            //        Data_Class.Data_Editing.ForAnl_NewMinSpec[l] = "0";
            //        Data_Class.Data_Editing.ForAnl_NewMaxSpec[l] = "0";
            //    }

            //    int j = 1;
            //    Data_Interface.New_HighSpec[0] = Convert.ToDouble(0);
            //    for (int i = Data_Interface.TheFirst_Trashes_Header_Count + 1; i < Data_Class.Data_Editing.ForAnl_NewMaxSpec.Length - Data_Interface.TheEnd_Trashes_Header_Count; i++)
            //    //for (int i = Data_Interface.TheFirst_Trashes_Header_Count + 1; i < Data_Interface.Getstring.Length - Data_Interface.TheEnd_Trashes_Header_Count; i++)
            //    {
            //        Data_Interface.New_HighSpec[j] = Convert.ToDouble(Data_Class.Data_Editing.ForAnl_NewMaxSpec[i]);
            //        Data_Interface.New_LowSpec[j] = Convert.ToDouble(Data_Class.Data_Editing.ForAnl_NewMinSpec[i]);
            //        j++;
            //    }

            //    DB_Interface.Delete_Spec_Data("newspec");
            //    DB_Interface.Insert_Spec_Data("newspec");
            //    checkBox2.Enabled = true;
            //}
            //else
            //{
            //    MessageBox.Show("Thres is no Spec. Please Press Save Spec button First.");
            //}

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Dialog.Reset();
            Dialog = new OpenFileDialog();
            Box_Enum = new Dictionary<int, Dictionary<int, string>>();
            //  Dialog.Filter = ".csv";
            Dialog.InitialDirectory = "C:\\Automation\\box_plot\\";
            Dialog.Multiselect = false;
            Dialog.ShowDialog();
            string[] Ignore_Spec = new string[2];

            bool flag = false;
            int Row = 0;
            int i = 0;

            if (Dialog.FileNames.Length > 0)
            {
                CSV_Class.CSV CSV = new CSV_Class.CSV();

                Csv_Interface = CSV.Open(Key);

                Csv_Interface.Read_Open(Dialog.FileNames[0]);
                while (!Csv_Interface.StreamReader.EndOfStream)
                {
                    Dictionary<int, string> Test = new Dictionary<int, string>();
                    string[] data = Csv_Interface.Read();

                    if (data[0].ToUpper() == "LABEL")
                    {
                        dataGridView3.ColumnCount = data.Length;

                        // _dataTable[f].Columns.Add("S_Min", typeof(double));
                        for (i = 0; i < data.Length; i++)
                        {
                            dataGridView3.Columns[i].Name = Convert.ToString(data[i]);
                            //  dataGridView3.Columns.Add(Convert.ToString(data[i]));
                        }

                        dataGridView3.Visible = true;

                     //   dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        dataGridView3.RowHeadersVisible = false;
                        dataGridView3.AllowUserToAddRows = false;
                      //  dataGridView3.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom);
                        //  dataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
                        //  dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

                        dataGridView3.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f =>

                        {

                            f.SortMode = DataGridViewColumnSortMode.NotSortable; // sort 막기


                                    });

                    }
                    else if (data[0] != "" && flag)
                    {

                        dataGridView3.Rows.Add("");
                   

                        for (i = 0; i < data.Length; i++)
                        {
                            dataGridView3.Rows[Row].Cells[i].Value = data[i].ToString();

                            if (i == 0)
                            {

                            }
                            else if(i == 1)
                            {
                                string[] split = data[2].Split('>');
                                int Key = 0;
                                bool Flag_Test = true;

                                for (int kk = 0; kk < split.Length; kk++)
                                {
                                    foreach (BoxPlot _i in Enum.GetValues(typeof(BoxPlot)))
                                    {
                                        string info = _i.ToString();
                                        int NB = (int)_i;

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


                            }
                            else if( i >= 2 && data[i].ToUpper().Contains("BY"))
                            {
                                string[] split1 = data[i].Split(':');

                                if(split1.Length > 1)
                                {
                                    for(int k = 1; k < split1.Length; k++)
                                    {
                                        Test.Add(888, split1[1]);
                                    }
                                }
                            }
                            else if (i >= 2 && data[i].ToUpper().Contains("IGNORE SPEC LIMIT"))
                            {
                                string[] split1 = data[i].Split(':');

                                if (split1.Length > 1)
                                {
                                    for (int k = 1; k < split1.Length; k++)
                                    {
                                        Test.Add(777, split1[1]);
                                    }
                                }
                            }

                        }
                        Row++;
                    }
                    else if (data[0].ToUpper() == "START")
                    {
                        flag = true;
                        dataGridView3.Rows.Add("");
                        for (int ii = 0; ii < data.Length; ii++)
                        {
                            dataGridView3.Rows[Row].Cells[ii].Value = data[ii].ToString();
                        }
                        Row++;
                    }

                    if (data[0].ToUpper() != "START" && data[0].ToUpper() != "LABEL" && data[0].ToUpper() != "")
                    {
                        if (!Box_Enum.ContainsKey(Convert.ToInt16(data[0])))
                        {

                            if (Test.Count != 0)
                                Box_Enum.Add(Convert.ToInt16(data[0]), Test);
                            //  Box_Enum.Add(Ignore_Spec[1]);
                        }
                    }
          
                }

            }
            Csv_Interface.Read_Close();

        }


        //if (Dialog.FileNames.Length > 0)
        //{
        //    CSV_Class.CSV CSV = new CSV_Class.CSV();

        //    Csv_Interface = CSV.Open(Key);

        //    Csv_Interface.Read_Open(Dialog.FileNames[0]);
        //    while (!Csv_Interface.StreamReader.EndOfStream)
        //    {
        //        string[] data = Csv_Interface.Read();

        //        if (data[0].ToUpper() == "LABEL")
        //        {
        //            dataGridView3.ColumnCount = data.Length;

        //            // _dataTable[f].Columns.Add("S_Min", typeof(double));
        //            for (int i = 0; i < data.Length; i++)
        //            {
        //                dataGridView3.Columns[i].Name = Convert.ToString(data[i]);
        //                //  dataGridView3.Columns.Add(Convert.ToString(data[i]));
        //            }

        //            dataGridView3.Visible = true;

        //            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        //            dataGridView3.RowHeadersVisible = false;
        //            dataGridView3.AllowUserToAddRows = false;
        //            //  dataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
        //            //  dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

        //            dataGridView3.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f =>

        //            {

        //                f.SortMode = DataGridViewColumnSortMode.NotSortable; // sort 막기


        //            });

        //        }
        //        else if (data[0] != "" && flag)
        //        {
        //            dataGridView3.Rows.Add("");


        //            for (int i = 0; i < data.Length; i++)
        //            {
        //                dataGridView3.Rows[Row].Cells[i].Value = data[i].ToString();

        //                if (data[i].ToString().ToUpper().Contains("IGNORE SPEC LIMIT"))
        //                {
        //                    string[] Split_dummy = data[i].Split(':');

        //                    Split_dummy = Split_dummy[1].Split('>');

        //                    for(int ii = 0; ii < Split_dummy.Length; ii++)
        //                    {
        //                        Ignore_Spec[ii] = Split_dummy[ii];

        //                    }

        //                    Array.Resize(ref Ignore_Spec, Split_dummy.Length);

        //                }
        //            }

        //            string[] split = data[2].Split('>');


        //            Dictionary<int, string> Test = new Dictionary<int, string>();
        //            int Key = 0;
        //            bool Flag_Test = true;

        //            for (int kk = 0; kk < split.Length; kk++)
        //            {
        //                foreach (BoxPlot i in Enum.GetValues(typeof(BoxPlot)))
        //                {
        //                    string info = i.ToString();
        //                    int NB = (int)i;

        //                    if (Flag_Test)
        //                    {
        //                        Test.Add(999, data[1]);
        //                        Flag_Test = false;
        //                    }
        //                    if (split[kk].Trim() == Convert.ToString(NB).Trim())
        //                    {
        //                        Test.Add(NB, info.Trim());
        //                        break;

        //                    }
        //                }
        //            }


        //            if (!Box_Enum.ContainsKey(Convert.ToInt16(data[0])))
        //            {
        //                Test.Add(888, Ignore_Spec[1]);
        //                Box_Enum.Add(Convert.ToInt16(data[0]), Test);
        //              //  Box_Enum.Add(Ignore_Spec[1]);
        //            }

        //         //   dataGridView3.Rows[Row].Cells[i].Value = data[i].ToString();

        //            Row++;
        //        }
        //        else if (data[0].ToUpper() == "START")
        //        {
        //            flag = true;
        //            dataGridView3.Rows.Add("");
        //            for (int i = 0; i < data.Length; i++)
        //            {
        //                dataGridView3.Rows[Row].Cells[i].Value = data[i].ToString();
        //            }
        //            Row++;
        //        }
        //    }

        //}
        //Csv_Interface.Read_Close();
        //   }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Box_Enum = new Dictionary<int, Dictionary<int, string>>();


            int Row = dataGridView3.RowCount;
            int Column = dataGridView3.ColumnCount;
            bool flag = false;
            Csv_Interface.Write_Open(Dialog.FileName);

            string ColumnName = "";
            for (int i = 0; i < Column; i++)
            {
                if (i == Column - 1)
                {
                    ColumnName += dataGridView3.Columns[i].Name;
                }
                else
                {
                    ColumnName += dataGridView3.Columns[i].Name + ",";
                }

            }

            Csv_Interface.Write(ColumnName);

            bool Flag_Test1 = false;

            string[] Ignore_Spec = new string[2];

            for (int j = 0; j < Row; j++)
            {
                Dictionary<int, string> Test = new Dictionary<int, string>();
                string RowValue = "";

                for (int i = 0; i < Column; i++)
                {

                    if (i == Column - 1)
                    {
                        if (dataGridView3.Rows[j].Cells[i].Value != null)
                            RowValue += dataGridView3.Rows[j].Cells[i].Value.ToString();
                    }
                    else
                    {
                        if (dataGridView3.Rows[j].Cells[i].Value != null)
                            RowValue += dataGridView3.Rows[j].Cells[i].Value.ToString() + ",";
                    }

                }

                string[] data = RowValue.Split(',');

                if (data[0].ToUpper() == "LABEL")
                {
                   
                }
                else if (data[0] != "" && flag)
                {

                 //   dataGridView3.Rows.Add("");


                    for (int i = 0; i < data.Length; i++)
                    {
                      //  dataGridView3.Rows[Row].Cells[i].Value = data[i].ToString();

                        if (i == 0)
                        {

                        }
                        else if (i == 1)
                        {
                            string[] split = data[2].Split('>');
                            int Key = 0;
                            bool Flag_Test = true;

                            for (int kk = 0; kk < split.Length; kk++)
                            {
                                foreach (BoxPlot _i in Enum.GetValues(typeof(BoxPlot)))
                                {
                                    string info = _i.ToString();
                                    int NB = (int)_i;

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


                        }
                        else if (i >= 2 && data[i].ToUpper().Contains("BY"))
                        {
                            string[] split1 = data[i].Split(':');

                            if (split1.Length > 1)
                            {
                                for (int k = 1; k < split1.Length; k++)
                                {
                                    Test.Add(888, split1[1]);
                                }
                            }
                        }
                        else if (i >= 2 && data[i].ToUpper().Contains("IGNORE SPEC LIMIT"))
                        {
                            string[] split1 = data[i].Split(':');

                            if (split1.Length > 1)
                            {
                                for (int k = 1; k < split1.Length; k++)
                                {
                                    Test.Add(777, split1[1]);
                                }
                            }
                        }

                    }
                  //  Row++;
                }
                else if (data[0].ToUpper() == "START")
                {
                    flag = true;
                    //dataGridView3.Rows.Add("");
                    //for (int ii = 0; ii < data.Length; ii++)
                    //{
                    //    dataGridView3.Rows[Row].Cells[ii].Value = data[ii].ToString();
                    //}
                    //Row++;
                }

                if (data[0].ToUpper() != "START" && data[0].ToUpper() != "LABEL" && data[0].ToUpper() != "")
                {
                    if (!Box_Enum.ContainsKey(Convert.ToInt16(data[0])))
                    {

                        if (Test.Count != 0)
                            Box_Enum.Add(Convert.ToInt16(data[0]), Test);
                        //  Box_Enum.Add(Ignore_Spec[1]);
                    }
                }
                Csv_Interface.Write(RowValue);
            }
            Csv_Interface.Write_Close();

        }
       
        //private void button2_Click_1(object sender, EventArgs e)
        //{
        //    Box_Enum = new Dictionary<int, Dictionary<int, string>>();


        //    int Row = dataGridView3.RowCount;
        //    int Column = dataGridView3.ColumnCount;

        //    Csv_Interface.Write_Open(Dialog.FileName);

        //    string ColumnName = "";
        //    for (int i = 0; i < Column; i++)
        //    {
        //        if (i == Column - 1)
        //        {
        //            ColumnName += dataGridView3.Columns[i].Name;
        //        }
        //        else
        //        {
        //            ColumnName += dataGridView3.Columns[i].Name + ",";
        //        }

        //    }

        //    Csv_Interface.Write(ColumnName);

        //    bool Flag_Test1 = false;
        //    bool Flag_Test = true;

        //    string[] Ignore_Spec = new string[2];

        //    for (int j = 0; j < Row; j++)
        //    {
        //        string RowValue = "";

        //        for (int i = 0; i < Column; i++)
        //        {

        //            if (i == Column - 1)
        //            {
        //                if (dataGridView3.Rows[j].Cells[i].Value != null)
        //                    RowValue += dataGridView3.Rows[j].Cells[i].Value.ToString();
        //            }
        //            else
        //            {
        //                if (dataGridView3.Rows[j].Cells[i].Value != null)
        //                    RowValue += dataGridView3.Rows[j].Cells[i].Value.ToString() + ",";
        //            }




        //        }

        //        string[] split = RowValue.Split(',');

        //        if (split[0].ToUpper() == "START")
        //        {
        //            Flag_Test1 = true;

        //        }

        //        if (Flag_Test1 && split[0].ToUpper() != "START" && split[0] != "")
        //        {
        //            string[] split2 = split[2].Split('>');

        //            Dictionary<int, string> Test = new Dictionary<int, string>();

        //            Ignore_Spec[0] = split[3].ToUpper().Trim();
        //            Ignore_Spec[1] = split[4].ToUpper().Trim();

        //            for (int kk = 0; kk < split2.Length; kk++)
        //            {
        //                foreach (BoxPlot i in Enum.GetValues(typeof(BoxPlot)))
        //                {
        //                    string info = i.ToString();
        //                    int NB = (int)i;

        //                    if (Flag_Test)
        //                    {
        //                        Test.Add(999, split[1]);
        //                        Flag_Test = false;
        //                    }
        //                    if (split2[kk].Trim() == Convert.ToString(NB).Trim())
        //                    {
        //                        Test.Add(NB, info.Trim());
        //                        break;

        //                    }
        //                }
        //            }
        //            if (!Box_Enum.ContainsKey(Convert.ToInt16(split[0])))
        //            {
        //                Test.Add(888, Ignore_Spec[1]);
        //                Box_Enum.Add(Convert.ToInt16(split[0]), Test);
        //            }

        //            Flag_Test = true;
        //        }
        //        Csv_Interface.Write(RowValue);





        //    }
        //    Csv_Interface.Write_Close();
        //}

        private void button3_Click_1(object sender, EventArgs e)
        {
            Valuse = new object[0];
            dataGridView3.Rows.Add(Valuse);
        }

        public void Cal_Yield2(int TotalCount)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            int index = tabControl1.SelectedIndex;


            bool while_Flag = true;
            int[] Yield = new int[Datagrid.Length];
            int DB = 0;

            for (int n = 0; n < DB_Interface.For_Any_Yield_Percent[0].Count; n++)
            {
                DB = 0;

                while (while_Flag)
                {
                    for (int j = 0; j < Datagrid.Length; j++)
                    {
                        for (int m = 0; m < DB_Interface.For_Any_Yield_Percent[DB][n][j].Count; m++)
                        {
                            if (DB_Interface.For_Any_Yield_Percent[DB][n][j][m] != 0)
                            {
                                Yield[j]++;
                                while_Flag = false;
                                break;
                            }
                        }

                    }
                    DB++;
                    if (DB == DB_Interface.For_Any_Yield_Percent.Length)
                    {
                        while_Flag = true;
                        break;
                    }

                }
                while_Flag = true;
            }

            for (int j = 0; j < Datagrid.Length; j++)
            {
                Datagrid[j].Rows[0].Cells[1].Value = Sample;
                Datagrid[j].Rows[1].Cells[1].Value = Sample - Hidden_Sample_Count;

                long Pass = 0;
                if ((Sample - Hidden_Sample_Count) == Yield[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (Sample - Hidden_Sample_Count) - Yield[j];
                }

                Datagrid[j].Rows[2].Cells[1].Value = Pass;
                Datagrid[j].Rows[3].Cells[1].Value = Yield[j];

                double Dummy = (Convert.ToDouble((Sample - Hidden_Sample_Count) - Yield[j]) / (Sample - Hidden_Sample_Count)) * 100;
                Datagrid[j].Rows[4].Cells[1].Value = Dummy;
                Datagrid[j].Rows[5].Cells[1].Value = Hidden_Sample_Count;

            }

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

        }
        public void For_New_Spec_Cal_Yield2(int TotalCount)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool while_Flag = true;
            int[] Yield = new int[MakeSpec.advanced.Length];
            int DB = 0;

            for (int n = 0; n < DB_Interface.For_Any_Yield_Percent_For_New_Spec[0].Count; n++)
            {
                DB = 0;

                while (while_Flag)
                {
                    for (int j = 0; j < MakeSpec.advanced.Length; j++)
                    {
                        for (int m = 0; m < DB_Interface.For_Any_Yield_Percent_For_New_Spec[DB][n][j].Count; m++)
                        {
                            if (DB_Interface.For_Any_Yield_Percent_For_New_Spec[DB][n][j][m] != 0)
                            {
                                Yield[j]++;
                                while_Flag = false;
                                break;
                            }
                        }

                    }
                    DB++;
                    if (DB == DB_Interface.For_Any_Yield_Percent_For_New_Spec.Length)
                    {
                        while_Flag = true;
                        break;
                    }

                }
                while_Flag = true;
            }

            for (int j = 0; j < MakeSpec.advanced.Length; j++)
            {
                MakeSpec.datagrid2[j].Rows[0].Cells[1].Value = Sample;
                MakeSpec.datagrid2[j].Rows[1].Cells[1].Value = Sample - Hidden_Sample_Count;

                long Pass = 0;
                if ((Sample - Hidden_Sample_Count) == Yield[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (Sample - Hidden_Sample_Count) - Yield[j];
                }

                MakeSpec.datagrid2[j].Rows[2].Cells[1].Value = Pass;
                MakeSpec.datagrid2[j].Rows[3].Cells[1].Value = Yield[j];

                double Dummy = (Convert.ToDouble((Sample - Hidden_Sample_Count) - Yield[j]) / (Sample - Hidden_Sample_Count)) * 100;
                MakeSpec.datagrid2[j].Rows[4].Cells[1].Value = Dummy;
                MakeSpec.datagrid2[j].Rows[5].Cells[1].Value = Hidden_Sample_Count;

            }

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }


        private void Cal_No_Thread(long Total)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool flag = false;
            int Db = 0;

            string Filename = DB_Interface.Filename.Substring(DB_Interface.Filename.LastIndexOf("\\") + 1);

            int length = DB_Interface.Filename.Length;
            Filename = DB_Interface.Filename.Substring(0, length - Filename.Length);


            

             //Csv_Interface.Write_Open(Filename + "Inf.csv");
             //Csv_Interface.Write("BIN:1");
             //Csv_Interface.Write("LOT:" + Lot[0]);
            // Csv_Interface.Write_Close();


            for (int n = 0; n < Total; n++)
            {
                for (int j = 0; j < Data_Interface.Clotho_Spcc_List[0].Max.Length; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {
                        for (int m = 0; m < DB_Interface.Yield_Test[Db][n][j].Count; m++)
                        {

                      
                            Sample_Verify[j]++;
                            Db = 0;
                            flag = true;
                            break;

                        }
                        if (flag) break;
                    }
                    flag = false;
                    Db = 0;
                }
            }

            for (int j = 0; j < Datagrid.Length; j++)
            {
                Datagrid[j].Rows[0].Cells[1].Value = Sample;
                Datagrid[j].Rows[1].Cells[1].Value = Sample - Hidden_Sample_Count;

                long Pass = 0;
                if ((Sample - Hidden_Sample_Count) == Sample_Verify[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (Sample - Hidden_Sample_Count) - Sample_Verify[j];
                }

                Datagrid[j].Rows[2].Cells[1].Value = Pass;
                Datagrid[j].Rows[3].Cells[1].Value = Sample_Verify[j];

                double Dummy = (Convert.ToDouble((Sample - Hidden_Sample_Count) - Sample_Verify[j]) / (Sample - Hidden_Sample_Count)) * 100;
                Datagrid[j].Rows[4].Cells[1].Value = Dummy;
                Datagrid[j].Rows[5].Cells[1].Value = Hidden_Sample_Count;

            }



            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }
        private void Cal_No_Thread_For_Delete_Unit(long Total)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool flag = false;
            int Db = 0;

            Sample_Verify = new int[Data_Interface.Clotho_Spcc_List[0].Max.Length];

            for (int n = 0; n < DB_Interface.Yield_Test[Db].Count; n++)
            {
                for (int j = 0; j < Data_Interface.Clotho_Spcc_List[0].Max.Length; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {
                        for (int m = 0; m < DB_Interface.Yield_Test[Db][n][j].Count; m++)
                        {

                            if (!Fail_Units.Contains(Convert.ToString(DB_Interface.Yield_Test[Db][n][j][m].SN)))
                            {
                                Sample_Verify[j]++;
                                Db = 0;
                                flag = true;
                                break;

                            }

                        }
                        if (flag) break;
                    }
                    flag = false;
                    Db = 0;
                }
            }

            for (int j = 0; j < Datagrid.Length; j++)
            {
                Datagrid[j].Rows[0].Cells[1].Value = Sample;
                Datagrid[j].Rows[1].Cells[1].Value = Sample - Hidden_Sample_Count;

                long Pass = 0;
                if ((Sample - Hidden_Sample_Count) == Sample_Verify[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (Sample - Hidden_Sample_Count) - Sample_Verify[j];
                }

                Datagrid[j].Rows[2].Cells[1].Value = Pass;
                Datagrid[j].Rows[3].Cells[1].Value = Sample_Verify[j];

                double Dummy = (Convert.ToDouble((Sample - Hidden_Sample_Count) - Sample_Verify[j]) / (Sample - Hidden_Sample_Count)) * 100;
                Datagrid[j].Rows[4].Cells[1].Value = Dummy;
                Datagrid[j].Rows[5].Cells[1].Value = Hidden_Sample_Count;

            }



            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }
        private void Cal_No_Thread_For_Lot(int Total)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool flag = false;
            int Db = 0;

            for (int n = 0; n < Total; n++)
            {
                for (int j = 0; j < 1; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {

                        for (int m = 0; m < DB_Interface.Yield_Test[Db][n][Selected_Bin].Count; m++)
                        {
                            int L = DB_Interface.Lot_Dic[databylot[n]];

                            Sample_Verify_Lot[L]++;
                            Db = 0;
                            flag = true;
                            break;

                        }
                        if (flag) break;
                    }
                    flag = false;
                    Db = 0;
                }
            }

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }
        private void Cal_No_Thread_For_Site(int Total)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool flag = false;
            int Db = 0;

            for (int n = 0; n < Total; n++)
            {
                for (int j = 0; j < 1; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {

                        for (int m = 0; m < DB_Interface.Yield_Test[Db][n][Selected_Bin].Count; m++)
                        {
                            int L = DB_Interface.Site_Dic[databylot[n]];

                            Sample_Verify_Lot[L]++;
                            Db = 0;
                            flag = true;
                            break;

                        }
                        if (flag) break;
                    }
                    flag = false;
                    Db = 0;
                }
            }

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }
        private void Cal_NewSpec_No_Thread(long Total)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool flag = false;
            int Db = 0;



            for (int n = 0; n < DB_Interface.Yield_Test_New_Spec[Db].Count; n++)
            {
                for (int j = 0; j < Data_Interface.Clotho_List[0].Max.Length; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {
                        for (int m = 0; m < DB_Interface.Yield_Test_New_Spec[Db][n][j].Count; m++)
                        {
                            Sample_Verify[j]++;
                            Db = 0;
                            flag = true;
                            break;

                        }
                        if (flag) break;
                    }
                    flag = false;
                    Db = 0;
                }
            }



            for (int j = 0; j < MakeSpec.datagrid.Length; j++)
            {
                MakeSpec.datagrid2[j].Rows[0].Cells[1].Value = Sample;
                MakeSpec.datagrid2[j].Rows[1].Cells[1].Value = Sample - Hidden_Sample_Count;

                long Pass = 0;
                if ((Sample - Hidden_Sample_Count) == Sample_Verify[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (Sample - Hidden_Sample_Count) - Sample_Verify[j];
                }

                MakeSpec.datagrid2[j].Rows[2].Cells[1].Value = Pass;
                MakeSpec.datagrid2[j].Rows[3].Cells[1].Value = Sample_Verify[j];

                double Dummy = (Convert.ToDouble((Sample - Hidden_Sample_Count) - Sample_Verify[j]) / (Sample - Hidden_Sample_Count)) * 100;
                MakeSpec.datagrid2[j].Rows[4].Cells[1].Value = Dummy;
                MakeSpec.datagrid2[j].Rows[5].Cells[1].Value = Hidden_Sample_Count;
                MakeSpec.datagrid2[j].Rows[6].Cells[1].Value = Outlier_List.Count;
                MakeSpec.datagrid2[j].Update();
                MakeSpec.Outlier_List = Outlier_List;

            }



            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }

        public void Cal_Yield_For_Lot_Variation(out int[] Total_Count, out int[] Pass_Count, out int[] Fail_Count)
        {
            int Count = 0;

            Total_Count = new int[LOT.Length];
            Pass_Count = new int[LOT.Length];
            Fail_Count = new int[LOT.Length];

            bool Flag = false;

            foreach (List<List<int>[]> items in DB_Interface.ForCampare_Yield)
            {
                Total_Count[Count] = items.Count;
                foreach (List<int>[] item in items)
                {

                    foreach (List<int> ite in item)
                    {
                        foreach (int it in ite)
                        {
                            if (it == 1)
                            {
                                Fail_Count[Count]++;
                                Flag = true;
                                break;
                            }


                            if (Flag) break;
                        }
                        if (Flag) break;
                    }
                    Count++;
                    Flag = false;
                    if (Count == Total_Count.Length)
                    {
                        Count = 0;
                    }
                }


            }

            for (int i = 0; i < LOT.Length; i++)
            {
                Pass_Count[i] = Total_Count[i] - Fail_Count[i];
            }

        }
        public void CSV_Data_Yield()
        {
            DB_Interface.ForCampare_Yield_List1 = new List<List<int>[]>();

            string Key = "YIELD";
            Data_Count = 1;

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            Csv_Interface = CSV.Open(Key);
            //Data_Interface = Data_Edit.Open(Key);

            Csv_Interface.Read_Open(CSV_File_Path);

           Data_Interface.Ref_New_Header = new string[1];
            Data_Interface.Ref_New_HighSpec = new double[1];
            Data_Interface.Ref_New_LowSpec = new double[1];

            #region Find_First_and_Spec_Row

            // Find First Row

            while (!Csv_Interface.StreamReader.EndOfStream)
            {
                Csv_Interface.Read();
                bool Flag = Data_Interface.Find_First_Row(Csv_Interface.Get_String);
                if (Flag) break;
            }

            // Find Spec High

            ForNewSpec = true;

            SetSpec();

            GetString_length = Csv_Interface.Get_String.Length;

            while (!Csv_Interface.StreamReader.EndOfStream)
            {
                Csv_Interface.Read();
                bool Flag = false;

                Flag = Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, ForNewSpec);

                if (Flag) break;
            }

            // Find Spec Low

            while (!Csv_Interface.StreamReader.EndOfStream)
            {
                Csv_Interface.Read();
                bool Flag = false;

                Flag = Data_Interface.Find_Spec_Row(Csv_Interface.Get_String, ForNewSpec);

                if (Flag) break;
            }



            #endregion

            DB_Interface.Insert_ThreadFlags = new ManualResetEvent[2];
            DB_Interface.Insert_Thread_Wait = new bool[2];

            for (int thread_i = 0; thread_i < 2; thread_i++)
            {
                DB_Interface.Insert_ThreadFlags[thread_i] = new ManualResetEvent(false);
            }

            string[] GetData = Csv_Interface.Read_Test();
            Data_Interface.Getstring = GetData;


            while (!Csv_Interface.StreamReader.EndOfStream)
            {
                if (!ForGross_Fail_Unit.Contains(Convert.ToString(Data_Count)))
                {

                    for (int thread_i = 0; thread_i < 2; thread_i++)
                    {
                        DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                    }
                    ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { CSV_Data_Yield_Cal(); }));

                    GetData = Csv_Interface.Read_Test();


                    DB_Interface.Insert_ThreadFlags[1].Set();

                    DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                    DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();

                    Data_Interface.Getstring = GetData;
                }
                else
                {
                    GetData = Csv_Interface.Read_Test();
                    Data_Interface.Getstring = GetData;
                }
                Data_Count++;
            }
            if (!ForGross_Fail_Unit.Contains(Convert.ToString(Data_Count)))
            {
                for (int thread_i = 0; thread_i < 2; thread_i++)
                {
                    DB_Interface.Insert_ThreadFlags[thread_i].Reset();
                }
                ThreadPool.QueueUserWorkItem(new WaitCallback((object state) => { CSV_Data_Yield_Cal(); }));

                DB_Interface.Insert_ThreadFlags[1].Set();

                DB_Interface.Insert_Thread_Wait[0] = DB_Interface.Insert_ThreadFlags[0].WaitOne();
                DB_Interface.Insert_Thread_Wait[1] = DB_Interface.Insert_ThreadFlags[1].WaitOne();
            }
            else
            {

            }


            Csv_Interface.Read_Close();

            double Testime = TestTime1.Elapsed.TotalMilliseconds;

            TestResult_Cal();

            List = DB_Interface.ForCampare_Yield_List1;
            //    Re_GridView(TestResult_Dic);


        }
        public void DB_Data_Yield()
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            ForNewSpec = true;
            Data_Interface._From_DB = true;
            SetSpec();
            string Query = "";

            LOT = new string[300];

            double Testtime54 = TestTime1.Elapsed.TotalMilliseconds;

            int Site_Count = 0;
            SITE = new string[300];

            By_Lot = new List<List<List<int>>[]>[Data_Interface.DB_Count];

            DB_Interface.Lot_Dic = new Dictionary<string, int>();
            DB_Interface.Site_Dic = new Dictionary<string, int>();

            Sample_Verify_Lot = new int[Matching_Lots.Count];

            int Lot_Dic_Indext = 0;


            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                Dictionary<string, List<string>> tests = key.Value;


                foreach (KeyValuePair<string, List<string>> ts in tests)
                {
                    Query = "Select DISTINCT SITEID from " + key.Key;
                    string[] Site_Inf = DB_Interface.Get_Data_By_Query(Query);

                    for (int f = 0; f < Site_Inf.Length; f++)
                    {
                        if (!SITE.Contains(Site_Inf[f]))
                        {
                            SITE[Site_Count] = Site_Inf[f];
                            Site_Count++;
                        }
                    }

                }

            }

            Array.Resize(ref SITE, Site_Count);
            Array.Sort(SITE);


            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                Dictionary<string, List<string>> tests = key.Value;


                foreach (KeyValuePair<string, List<string>> ts in tests)
                {

                    DB_Interface.Lot_Dic.Add(ts.Key, Lot_Dic_Indext);
                    Lot_Dic_Indext++;
                }

            }

            double Testtim51 = TestTime1.Elapsed.TotalMilliseconds;

            BIN = new string[Data_Interface.SWBIN_Dic.Count];

            int dummy = 0;
            foreach (KeyValuePair<string, Data_Class.Data_Editing.SWBIN> a in Data_Interface.SWBIN_Dic)
            {
                BIN[dummy] = a.Key.ToString(); dummy++;
            }



            for (int Site_Dic_Indext = 0; Site_Dic_Indext < SITE.Length; Site_Dic_Indext++)
            {
                DB_Interface.Site_Dic.Add(SITE[Site_Dic_Indext], Site_Dic_Indext);

            }

            TestTime1.Restart();
            TestTime1.Start();


            DB_Interface.Matching_Lots = Matching_Lots;


            DB_Interface.Get_From_Db_Data_for_Anly(Data_Interface);


            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

            TestResult_Cal();

            double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;


            Re_GridView();

            double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void DB_Data_Yield_For_NewSpec()
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            No_Index = new string[Data_Interface.Reference_Header.Length - 1];
            Paraname = new string[Data_Interface.Reference_Header.Length - 1];

            if (Data_Interface.Customor_Clotho_List == null || Data_Interface.Customor_Clotho_List.Count == 0)
            {

                double[] Spec_Min = new double[MakeSpec.advanced.Length];
                double[] Spec_Max = new double[MakeSpec.advanced.Length];


                for (int h = 0; h < MakeSpec.advanced.Length; h++)
                {
                    Spec_Min[h] = -9999;
                    Spec_Max[h] = 9999;
                }

                Data_Class.Data_Editing.Clotho_Spec New_Spec = new Data_Class.Data_Editing.Clotho_Spec(Spec_Min, Spec_Max);

                Data_Interface.Clotho_List.Add(New_Spec);


                for (int k = 0; k < Data_Interface.Reference_Header.Length - 1; k++)
                {
                    No_Index[k] = MakeSpec.advanced[0].Rows[k].Cells[0].Value.ToString();
                    Paraname[k] = MakeSpec.advanced[0].Rows[k].Cells[1].Value.ToString();

                    Spec_Min = new double[MakeSpec.advanced.Length];
                    Spec_Max = new double[MakeSpec.advanced.Length];

                    for (int h = 0; h < MakeSpec.advanced.Length; h++)
                    {
                        Spec_Min[h] = Convert.ToDouble(MakeSpec.advanced[h].Rows[k].Cells[4].Value);
                        Spec_Max[h] = Convert.ToDouble(MakeSpec.advanced[h].Rows[k].Cells[5].Value);
                    }

                    New_Spec = new Data_Class.Data_Editing.Clotho_Spec(Spec_Min, Spec_Max);

                    Data_Interface.Clotho_List.Add(New_Spec);

                }
            }


            ForNewSpec = true;
            Data_Interface._From_DB = true;
            // SetSpec();
            string Query = "";
            LOT = new string[0];

            for (int k = 0; k < DB_Interface.Table_Count; k++)
            {
                Query = "Select DISTINCT LOT_ID from data" + k;
                string[] datas = DB_Interface.Get_Data_By_Query(Query);

                LOT = LOT.Concat(datas).ToArray();
            }
            LOT = LOT.Distinct().ToArray();
            Array.Sort(LOT);


            SITE = new string[0];

            for (int k = 0; k < DB_Interface.Table_Count; k++)
            {
                Query = "Select DISTINCT SITE from data" + k;
                string[] datas = DB_Interface.Get_Data_By_Query(Query);

                SITE = SITE.Concat(datas).ToArray();
            }
            SITE = LOT.Distinct().ToArray();
            Array.Sort(SITE);



            double Testtime0 = TestTime1.Elapsed.TotalMilliseconds;

            DB_Interface.Get_From_Db_Data_for_Anly_For_New_Spec(Data_Interface);

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

            For_New_Spec_TestResult_Cal();

            double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

            For_New_Spec_Re_GridView();

            double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void CSV_Data_Yield_Cal()
        {
            DB_Interface.ForCampare_Yield_List = new List<int>[Data_Interface.DB_Count];

            for (int i = 0; i < Data_Interface.DB_Count; i++)
            {
                DB_Interface.ForCampare_Yield_List[i] = new List<int>();
            }

            for (int i = 0; i < DB_Interface.ForCampare_Yield_List.Length; i++)
            {
                for (int j = 0; j < Data_Interface.Per_DB_Column_Count[i]; j++)
                {
                    DB_Interface.ForCampare_Yield_List[i].Add(0);
                }
            }

            for (int i = 0; i < Data_Interface.DB_Count; i++)
            {
                DB_Interface.ThreadFlags[i] = new ManualResetEvent(false);
                ThreadPool.QueueUserWorkItem(new WaitCallback(CSV_Data_Yield_Cal_Thread), i);
            }

            for (int i = 0; i < Data_Interface.DB_Count; i++)
            {
                DB_Interface.Wait[i] = DB_Interface.ThreadFlags[i].WaitOne();
            }

            DB_Interface.ForCampare_Yield_List1.Add(DB_Interface.ForCampare_Yield_List);

            DB_Interface.Insert_ThreadFlags[0].Set();
        }
        public void CSV_Data_Yield_Cal_Thread(Object threadContext)
        {
            int i = (int)threadContext;

            DB_Interface.TestTime1[i].Restart();
            DB_Interface.TestTime1[i].Start();

            int k = 0;

            if (i == 0)
            {
                DB_Interface.ForCampare_Yield_List[0][0] = 0;
            }
            else
            {
                if (Data_Interface.New_HighSpec[Data_Interface.DB_Column_Limit * i] < Convert.ToDouble(Data_Interface.Getstring[(Data_Interface.DB_Column_Limit * i) + DB_Interface.TheFirst_Trashes_Header_Count]) || Data_Interface.New_LowSpec[Data_Interface.DB_Column_Limit * i] > Convert.ToDouble(Data_Interface.Getstring[(Data_Interface.DB_Column_Limit * i) + DB_Interface.TheFirst_Trashes_Header_Count]))
                {
                    DB_Interface.ForCampare_Yield_List[i][0] = 1;
                }
            }

            for (k = 1; k < Data_Interface.Per_DB_Column_Count[i] - 1; k++)
            {

                if (Data_Interface.New_HighSpec[Data_Interface.DB_Column_Limit * i + k] < Convert.ToDouble(Data_Interface.Getstring[(Data_Interface.DB_Column_Limit * i) + DB_Interface.TheFirst_Trashes_Header_Count + k]) || Data_Interface.New_LowSpec[Data_Interface.DB_Column_Limit * i + k] > Convert.ToDouble(Data_Interface.Getstring[(Data_Interface.DB_Column_Limit * i) + DB_Interface.TheFirst_Trashes_Header_Count + k]))
                {
                    DB_Interface.ForCampare_Yield_List[i][k] = 1;
                }

            }

            if (Data_Interface.New_HighSpec[Data_Interface.DB_Column_Limit * i + k] < Convert.ToDouble(Data_Interface.Getstring[(Data_Interface.DB_Column_Limit * i) + DB_Interface.TheFirst_Trashes_Header_Count + k]) || Data_Interface.New_LowSpec[Data_Interface.DB_Column_Limit * i + k] > Convert.ToDouble(Data_Interface.Getstring[(Data_Interface.DB_Column_Limit * i) + DB_Interface.TheFirst_Trashes_Header_Count + k]))
            {
                DB_Interface.ForCampare_Yield_List[i][Data_Interface.Per_DB_Column_Count[i] - 1] = 1;
            }


            DB_Interface.Testtime[i] = DB_Interface.TestTime1[i].Elapsed.TotalMilliseconds;

            DB_Interface.ThreadFlags[i].Set();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //  string text = comboBox1.Text;
            //  double Persent = Gross_Check[text];


            //   DB_Interface.List_Gross_Values = new List<Dictionary<string, double[]>[]>();
            //   DB_Interface.Gross_Values1 = new Dictionary<string, double[]>[Data_Interface.DB_Count];

            //   for (int i = 0; i < Data_Interface.DB_Count; i++)
            //   {
            //       DB_Interface.Gross_Values1[i] = new Dictionary<string, double[]>();
            //   }

            // DB_Interface.Get_Gross_Check_Para(Data_Interface, text, Persent, ForGross_Fail_Unit);

            //   int Count = 0;
            //   foreach (Dictionary<string, double[]>[] item in DB_Interface.List_Gross_Values)
            //   {
            //       foreach (Dictionary<string, double[]> items in item)
            //       {
            //           foreach (KeyValuePair<string, double[]> o in items)
            //           {
            //               Count++;
            //           }
            //       }

            //   }
            //   if (Count != 0)
            //   {
            //       button7.Enabled = true;

            //       Csv_Interface.Write_Open("C:\\temp\\dummy\\Data.csv");
            //       Csv_Interface.Write(DB_Interface.ID, DB_Interface.List_Gross_Values);
            //       Csv_Interface.Write_Close();

            //       JMP_Draw_For_Gross("C:\\temp\\dummy\\Data.csv");
            //   }


        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text = comboBox2.Text;

            if (comboBox3.Text.ToString() == "")
            {
                MessageBox.Show("Please, Select Bin Number");
            }
            else
            {

                #region

                if (DB_Interface._From_Db)
                {
                    if (text.ToUpper() == "LOT")
                    {
                        Dictionary<string, int>[] Update_Dic = new Dictionary<string, int>[Matching_Lots.Count];
                        Stopwatch TestTime1 = new Stopwatch();
                        TestTime1.Restart();
                        TestTime1.Start();

                        string Query = "";
                        int[] Total = new int[Matching_Lots.Count];
                        int v = 0;

                        databylot = new string[0];

                        //foreach (KeyValuePair<string, string> t in Matching_Lots)
                        //{

                        //    //  Query = "Select LOTID from " + t.Value + " where FAIL not like '1' and LOTID = '" + LOT[v] + "'";
                        //    Query = "Select LOTID from " + t.Value + " where FAIL not like '1'";
                        //    string[] datas = DB_Interface.Get_Data_By_Query(Query);

                        //    databylot = databylot.Concat(datas).ToArray();
                        //    Total[v] = datas.Length;
                        //    v++;
                        //}



                        //for (int v = 0; v < Matching_Lots.Count; v++)
                        //{
                        //    databylot = new string[0];

                        //    for (int loop = 0; loop < DB_Interface.Table_Count; loop++)
                        //    {
                        //        Query = "Select LOTID from data" + loop + " where FAIL not like '1' and LOTID = '" + LOT[v] + "'";
                        //        string[] datas = DB_Interface.Get_Data_By_Query(Query);

                        //        databylot = databylot.Concat(datas).ToArray();

                        //    }

                        //    Total[v] = databylot.Length;
                        //}

                        databylot = new string[0];
                        int i = 0;
                        foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                        {
                            Dictionary<string, List<string>> tests = key.Value;


                            foreach (KeyValuePair<string, List<string>> ts in tests)
                            {


                                Query = "Select LOTID from " + key.Key + " where FAIL not like '1'";
                                string[] datas = DB_Interface.Get_Data_By_Query(Query);

                                databylot = databylot.Concat(datas).ToArray();
                                Total[i] = datas.Length;
                                i++;
                            }
                          
                        }
                    


                        Selected_Bin = Convert.ToInt16(comboBox3.Text.ToString()) - 1;

                        for (int Lot_Dic_Indext = 0; Lot_Dic_Indext < Matching_Lots.Count; Lot_Dic_Indext++)
                        {

                            Update_Dic[Lot_Dic_Indext] = new Dictionary<string, int>();
                        }


                        Sample_Verify_Lot = new int[Matching_Lots.Count];
                        Cal_No_Thread_For_Lot(DB_Interface.Yield_Test[0].Count());


                        for (int j = 0; j < Data_Interface.DB_Count; j++)
                        {
                            int Length = Data_Interface.Per_DB_Column_Count[j];
                            int offset = 0;
                            int offset2 = 9;

                            if (Data_Interface.DB_Count - 1 == j)
                            {
                                Length = Length;
                                offset = 9;
                                offset2 = 0;
                            }
                            else if (j == 0)
                            {
                                Length = Length - 9;
                                offset = 0;
                                offset2 = 9;
                            }
                            else
                            {
                                Length = Length;
                                offset = 9;
                                offset2 = 0;
                            }


                            for (int k = 0; k < Length; k++)
                            {
                                //  for (int l = 0; l < 1; l++)
                                for (int l = 0; l < Matching_Lots.Count; l++)
                                {

                                    Update_Dic[l].Add(Data_Interface.Ref_New_Header[(j * Data_Interface.DB_Column_Limit - offset) + k], DB_Interface.For_Any_Yield_For_Lot[j][Selected_Bin][l][k + offset2]);

                                }

                            }

                        }

                        int q = 0;
                        LOT = new string[this.Matching_Lots.Count];
                        foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                        {
                            Dictionary<string, List<string>> tests = key.Value;


                            foreach (KeyValuePair<string, List<string>> ts in tests)
                            {

                                LOT[q] = key.Key.ToString(); q++;
                            }
                        }

     

                        Array.Resize(ref LOT, q);
                        double Testime = TestTime1.Elapsed.TotalMilliseconds;

                        Lot_Variation_Form Lot_form = new Lot_Variation_Form(LOT, Total, Update_Dic, Sample_Verify_Lot, DB_Interface.Filename);
                        Lot_form.Text = "Lot Variation";
                        Lot_form.Show();

                    }
                    else if (text.ToUpper() == "SITE")
                    {
                        Dictionary<string, int>[] Update_Dic = new Dictionary<string, int>[SITE.Length];
                        Stopwatch TestTime1 = new Stopwatch();
                        TestTime1.Restart();
                        TestTime1.Start();

                        string Query = "";
                        int[] Total = new int[SITE.Length];

                        for (int v = 0; v < SITE.Length; v++)
                        {
                            databylot = new string[0];

                            for (int loop = 0; loop < DB_Interface.Table_Count; loop++)
                            {
                                Query = "Select SITEID from data" + loop + " where FAIL not like '1' and SITEID = '" + SITE[v] + "'";
                                string[] datas = DB_Interface.Get_Data_By_Query(Query);

                                databylot = databylot.Concat(datas).ToArray();
                            }
                            Total[v] = databylot.Length;

                        }


                        databylot = new string[0];
                        for (int loop = 0; loop < DB_Interface.Table_Count; loop++)
                        {
                            Query = "Select SITEID from data" + loop + " where FAIL not like '1'";
                            string[] datas = DB_Interface.Get_Data_By_Query(Query);

                            databylot = databylot.Concat(datas).ToArray();

                        }


                        Selected_Bin = Convert.ToInt16(comboBox3.Text.ToString()) - 1;


                        for (int Site_Dic_Indext = 0; Site_Dic_Indext < SITE.Length; Site_Dic_Indext++)
                        {
                            Update_Dic[Site_Dic_Indext] = new Dictionary<string, int>();
                        }

                        Sample_Verify_Lot = new int[Data_Interface.Clotho_Spcc_List[0].Max.Length];
                        Cal_No_Thread_For_Site(DB_Interface.Yield_Test[0].Count());


                        for (int j = 0; j < Data_Interface.DB_Count; j++)
                        {
                            int Length = Data_Interface.Per_DB_Column_Count[j];

                            if (Data_Interface.DB_Count - 1 == j)
                            {
                                Length = Length - 10;
                            }
                            for (int k = 0; k < Length; k++)
                            {
                                for (int l = 0; l < SITE.Length; l++)
                                {
                                    Update_Dic[l].Add(Data_Interface.Ref_New_Header[j * Data_Interface.DB_Column_Limit + k], DB_Interface.For_Any_Yield_For_SITE[j][Selected_Bin][l][k]);
                                }

                            }

                        }

                        By_Site = new List<List<List<int>>[]>[Data_Interface.DB_Count];

                        double Testime = TestTime1.Elapsed.TotalMilliseconds;

                        Lot_Variation_Form Lot_form = new Lot_Variation_Form(SITE, Total, Update_Dic, Sample_Verify_Lot, CSV_File_Path);
                        Lot_form.Text = "Site Variation";
                        Lot_form.Show();
                    }
                    else if (text.ToUpper() == "BIN")
                    {
                        Dictionary<string, int>[] Update_Dic = new Dictionary<string, int>[BIN.Length];
                        Stopwatch TestTime1 = new Stopwatch();
                        TestTime1.Restart();
                        TestTime1.Start();

                        string Query = "";

                        Query = "Select bin from data where FAIL not like '1'";
                        databybin = DB_Interface.Get_Data_By_Query(Query);


                        Query = "Select DISTINCT bin from data";
                        string[] Bin_Count = DB_Interface.Get_Data_By_Query(Query);

                        int[] Total = new int[Bin_Count.Length];

                        Array.Sort(Total);
                        Array.Sort(Bin_Count);

                        for (int v = 0; v < Total.Length; v++)
                        {
                            Query = "Select count(bin) from data where bin like " + Bin_Count[v];
                            string[] dummy = DB_Interface.Get_Data_By_Query(Query);
                            Total[v] = Convert.ToInt16(dummy[0]);

                        }


                        Selected_Bin = Convert.ToInt16(comboBox3.Text.ToString()) - 1;

                        By_Bin = new List<List<List<int>>[]>[Data_Interface.DB_Count];

                        Bin_Dic = new Dictionary<string, int>();

                        Bin_Yield = new int[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                        string[] databin = new string[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                        for (int Bin_Dic_Indext = 0; Bin_Dic_Indext < Bin_Yield.Length; Bin_Dic_Indext++)
                        {

                            Update_Dic[Bin_Dic_Indext] = new Dictionary<string, int>();
                        }


                        Sample_Verify = new int[Data_Interface.Clotho_Spcc_List[0].Max.Length];
                        Cal_No_Thread(DB_Interface.Yield_Test[0].Count());

                        double Testime3 = TestTime1.Elapsed.TotalMilliseconds;

                        for (int j = 0; j < Data_Interface.DB_Count; j++)
                        {
                            for (int k = 0; k < Data_Interface.Per_DB_Column_Count[j]; k++)
                            {
                                for (int l = 0; l < Data_Interface.Clotho_Spcc_List[0].Max.Length; l++)
                                {
                                    Update_Dic[l].Add(Data_Interface.Ref_New_Header[j * Data_Interface.DB_Column_Limit + k], DB_Interface.For_Any_Yield[j][l][k]);
                                }

                            }

                        }


                        double Testime4 = TestTime1.Elapsed.TotalMilliseconds;

                        Dictionary<string, Data_Class.Data_Editing.Clotho_Spec> For_Bin = new Dictionary<string, Data_Class.Data_Editing.Clotho_Spec>();

                        for (int index = 0; index < Data_Interface.Clotho_List.Count; index++)
                        {
                            For_Bin.Add(Data_Interface.Reference_Header[index], Data_Interface.Clotho_List[index]);
                        }

                        int sample = 0;

                        for (int index = 0; index < Total.Length; index++)
                        {
                            sample += Total[index];
                        }


                        Lot_Variation_Form Lot_form = new Lot_Variation_Form(BIN, sample, Total, Update_Dic, Bin_Yield, For_Bin);
                        Lot_form.Text = "Bin Variation";
                        Lot_form.Show();

                        GC.Collect(0, GCCollectionMode.Forced);
                        GC.WaitForFullGCComplete();
                    }

                }
                #endregion
            }
        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string text = comboBox3.Text;

            //Selected_Bin = Convert.ToInt16(text) - 1;
            //Data_Interface.New_LowSpec = new double[Data_Interface.Clotho_Spcc_List.Count];
            //Data_Interface.New_HighSpec = new double[Data_Interface.Clotho_Spcc_List.Count];

            //for (int i = 0; i < Data_Interface.Clotho_Spcc_List.Count; i++)
            //{
            //    Data_Interface.New_LowSpec[i] = Data_Interface.Clotho_Spcc_List[i].Min[Convert.ToInt16(text) - 1];
            //    Data_Interface.New_HighSpec[i] = Data_Interface.Clotho_Spcc_List[i].Max[Convert.ToInt16(text) - 1];
            //}


            //int scrollPosition = advancedDataGridView1.FirstDisplayedScrollingRowIndex;

            //Stopwatch TestTime1 = new Stopwatch();
            //TestTime1.Restart();
            //TestTime1.Start();

            //string[] No_Index = new string[advancedDataGridView1.RowCount];
            //string[] Paraname = new string[advancedDataGridView1.RowCount];

            //for (int k = 0; k < advancedDataGridView1.RowCount; k++)
            //{
            //    No_Index[k] = advancedDataGridView1.Rows[k].Cells[0].Value.ToString();
            //    Paraname[k] = advancedDataGridView1.Rows[k].Cells[1].Value.ToString();
            //}



            //_dataTable.Rows.Clear();

            //double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;
            //#region
            //object[] Valuse = new object[Coulumn_Count];

            //_dataTable.BeginLoadData();
            //for (int i = 0; i < Bin_Infor[Convert.ToInt16(text) - 1].Count; i++)
            //{
            //    Valuse[0] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].No;
            //    Valuse[1] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Para;
            //    Valuse[2] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Spec_min;
            //    Valuse[3] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Spec_max;
            //    Valuse[4] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Data_min;
            //    Valuse[5] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Data_median;
            //    Valuse[6] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Data_max;
            //    Valuse[7] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].CPK;
            //    Valuse[8] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Std;
            //    Valuse[9] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Pecent;
            //    Valuse[10] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].Fail;
            //    Valuse[11] = null;
            //    Valuse[12] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].N_Spec_min;
            //    Valuse[13] = Bin_Infor[Convert.ToInt16(text) - 1][Paraname[i]].N_Spec_max;

            //    _dataTable.Rows.Add(Valuse);
            //}
            //_dataTable.EndLoadData();




            //double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
            //#endregion
            //ForeColor();
            //Cal_Yield2(Sample - ForGross_Fail_Unit.Count);
            //double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;
            //bindingSource.DataSource = _dataTable;

            //advancedDataGridView1.FirstDisplayedScrollingRowIndex = scrollPosition;

        }
  
        private void JMP_Draw(string FilePaht, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<int, Dictionary<int, string>> OrderbySequence ,string Key, string[] X, string[] By)
        {
            JMP_Interface.Open_Session(true);

            Distribition_Select_Spec_Method();
            JMP_File = FilePaht;

            JMP_Interface.Open_Document(FilePaht);
            JMP_Interface.GetDataTable();

            JMP_Class.Script Distribution_Script;

            Distribution_Script = null;
            List<string>[] Para_Test = new List<string>[OrderbySequence.Count];
            Dictionary<int, Dictionary<int, string>> dummy = new Dictionary<int, Dictionary<int, string>>();

            switch (Key)
            {
                case "Fit Y X Lot":
                case "Fit Y X Site":
                case "Distributions":

                    //  DB_Interface.Variation = DupCheck<object>(DB_Interface.Variation);
                    Distribution_Script = JMP_Interface.Make_Script(Key, Data, null, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value ,X, By);
                    break;
                case "SITE":
                //    DB_Interface.Variation = DupCheck<object>(DB_Interface.Variation);
              //      Distribution_Script = JMP_Interface.Make_Script("DISTRIBUTION", Parameter, Parameter2, DB_Interface.Value, DB_Interface.Variation, null, "SITE", Customer_enable, NPI_enable, CPK_enable, CPK_Value);
                    break;
                case "BIN":
              //      DB_Interface.Variation = DupCheck<object>(DB_Interface.Variation);
              //      Distribution_Script = JMP_Interface.Make_Script("DISTRIBUTION", Parameter, Parameter2, DB_Interface.Value, DB_Interface.Variation, Spec, "BIN", Customer_enable, NPI_enable, CPK_enable, CPK_Value);
                    break;

                case "BoxPlot":
                    //  DB_Interface.Variation = DupCheck<object>(DB_Interface.Variation);

                    Distribution_Script = JMP_Interface.Make_Script(Key, Data, null, FilePaht, OrderbySequence, false, ref Para_Test, Customer_enable, NPI_enable, CPK_enable, CPK_Value);
                    break;

                case "Distribution":

                    Distribution_Script = JMP_Interface.Make_Script(Key, Data, null, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X, By);

                    break;

            }


            //if (ForGross_Fail_Unit.Count != 0)
            //{
            //  //  JMP_Class.Script Distribution_HideAndExclude = JMP_Interface.Distribution_HideAndExclude_1("DISTRIBUTION_HideAndExclude_1", ForGross_Fail_Unit);

            //    Csv_Interface.Write_Open("C:\\temp\\dummy\\dummy3.jsl");
            ////    Csv_Interface.WriteScript(Distribution_HideAndExclude.Scrip_Data);
            //    Csv_Interface.Write_Close();

            //    JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy3.jsl");

            //}

            string Script = Distribution_Script.Scrip_Data;
            string[] Split = Script.Split('#');

            for (int k = 0; k < Split.Length; k++)
            {

                Csv_Interface.Write_Open("C:\\temp\\dummy\\dummy.jsl");
                //Csv_Interface.WriteScript(Distribution_Script.Scrip_Data);
                Csv_Interface.WriteScript(Split[k]);
                Csv_Interface.Write_Close();

                JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy.jsl");
            }



        }

        private void JMP_Draw_For_Gross(string Filepath, Dictionary<string, DB_Class.DB_Editing.Gross> Dic_for_Gross)
        {

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            JMP_Interface.Open_Session(true);


            JMP_File = "C:\\temp\\dummy\\Data.csv";

            double Testtime0 = TestTime1.Elapsed.TotalMilliseconds;

            JMP_Interface.Open_Document(JMP_File);

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

            JMP_Interface.GetDataTable();

            double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;
            //     
            //JMP_Class.Script Transpose_Script = JMP_Interface.Transpose("Transpose", JMP_File, DB_Interface.List_Gross_Values, DB_Interface.ID);

            //Csv_Interface.Write_Open("C:\\temp\\dummy\\dummy1.jsl");
            //Csv_Interface.WriteScript(Transpose_Script.Scrip_Data);
            //Csv_Interface.Write_Close();

            //JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy1.jsl");

            var Dsec = Dic_for_Gross.OrderByDescending(num => num.Value.STD);

            Dictionary<string, DB_Class.DB_Editing.Gross> Dic_for_Gro = Dsec.ToDictionary(t => t.Key, t => t.Value);

            // List<string> List_Para = Dsec.ToList();

            double Testtim3 = TestTime1.Elapsed.TotalMilliseconds;

            JMP_Class.Script Distribution_for_Gross_Script = JMP_Interface.Distribution_for_Gross("DISTRIBUTION_FOR_GROSS", Dic_for_Gro, DB_Interface.ID);

            double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;
            //    JMP_Interface.Open_Document("C:\\temp\\dummy\\dummy.jmp");

            //if (Dic_for_Gro.Count != 0)
            //{
            //    JMP_Class.Script Distribution_HideAndExclude = JMP_Interface.Distribution_HideAndExclude("DISTRIBUTION_HideAndExclude", Dic_for_Gro);

            //    Csv_Interface.Write_Open("C:\\temp\\dummy\\dummy3.jsl");
            //    Csv_Interface.WriteScript(Distribution_HideAndExclude.Scrip_Data);
            //    Csv_Interface.Write_Close();

            //    JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy3.jsl");

            //}


            Csv_Interface.Write_Open("C:\\temp\\dummy\\dummy2.jsl");
            double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;
            Csv_Interface.WriteScript(Distribution_for_Gross_Script.Scrip_Data);
            double Testtime6 = TestTime1.Elapsed.TotalMilliseconds;
            Csv_Interface.Write_Close();

            double Testtime7 = TestTime1.Elapsed.TotalMilliseconds;
            JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy2.jsl");

            double Testtime8 = TestTime1.Elapsed.TotalMilliseconds;
        }

        private void dataGridView1_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;


        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int Row = e.RowIndex;
            int Col = e.ColumnIndex;

            if (Col == 1)
            {
                string P = dataGridView1.Rows[Row].Cells[Col - 1].Value.ToString();
                string Value = dataGridView1.Rows[Row].Cells[Col].Value.ToString();

                this.Para[P].Range = Convert.ToDouble(Value);
            }



        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            string text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            double Persent = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Value);
            string Selector = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 2].Value.ToString();

            DB_Interface.List_Gross_Values = new List<Dictionary<string, DB_Class.DB_Editing.Gross>[]>();
            DB_Interface.Gross_Values1 = new Dictionary<string, DB_Class.DB_Editing.Gross>[Data_Interface.DB_Count];

            for (int i = 0; i < Data_Interface.DB_Count; i++)
            {
                DB_Interface.Gross_Values1[i] = new Dictionary<string, DB_Class.DB_Editing.Gross>();
            }

            Selected_Bin = Convert.ToInt16(comboBox3.Text) - 1;

            DB_Interface.Get_Gross_Check_Para(Data_Interface, text, Persent, Selector, Selected_Bin);

            double Testtime8 = TestTime1.Elapsed.TotalMilliseconds;

            int Count = 0;
            Dictionary<string, DB_Class.DB_Editing.Gross> SortbySTD = new Dictionary<string, DB_Class.DB_Editing.Gross>();

            foreach (Dictionary<string, DB_Class.DB_Editing.Gross>[] item in DB_Interface.List_Gross_Values)
            {
                foreach (Dictionary<string, DB_Class.DB_Editing.Gross> items in item)
                {
                    foreach (KeyValuePair<string, DB_Class.DB_Editing.Gross> o in items)
                    {
                        SortbySTD.Add(o.Key.ToString(), o.Value);
                        Count++;
                    }
                }

            }

            var Dsec = SortbySTD.OrderByDescending(num => num.Value.STD);

            double Testtime0 = TestTime1.Elapsed.TotalMilliseconds;

            if (Count != 0)
            {
                StringBuilder sb = new StringBuilder();

                string key = "";

                Csv_Interface.Write_Open("C:\\temp\\dummy\\Data.csv");

                int Count_Row = 0;

                sb.Append("Label,");

                foreach (KeyValuePair<string, DB_Class.DB_Editing.Gross> item in Dsec)
                {

                    if (SortbySTD.Count - 1 != Count_Row)
                    {
                        sb.Append(item.Key.ToString() + ",");
                    }
                    else if (SortbySTD.Count - 1 == Count_Row)
                    {
                        key = item.Key.ToString();
                        sb.Append(item.Key.ToString());
                        Csv_Interface.Write(sb.ToString(), "", 0);
                        break;
                    }
                    Count_Row++;

                }
                sb = new StringBuilder();
                for (int k = 0; k < DB_Interface.ID.Length; k++)
                {
                    sb = new StringBuilder();
                    Count_Row = 0;

                    sb.Append(DB_Interface.ID[k] + ",");

                    foreach (KeyValuePair<string, DB_Class.DB_Editing.Gross> item in Dsec)
                    {

                        if (SortbySTD.Count - 1 != Count_Row)
                        {
                            sb.Append(item.Value.Data[k] + ",");
                        }
                        else if (SortbySTD.Count - 1 == Count_Row)
                        {
                            sb.Append(item.Value.Data[k]);
                            Csv_Interface.Write(sb.ToString(), "", 0);
                            break;
                        }
                        Count_Row++;

                    }

                }

                Csv_Interface.Write_Close();

                double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
                JMP_Draw_For_Gross("C:\\temp\\dummy\\Data.csv", SortbySTD);

                double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;
            }
        }

        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
        
            int index = tabControl1.SelectedIndex;

            if (e.ColumnIndex == 1 && e.RowIndex != -1)
            {
                Dictionary<string, CSV_Class.For_Box> Dic_Test = new Dictionary<string, CSV_Class.For_Box>();

                DB_Interface.Line = new string[1];

                DB_Interface.Line[0] = advancedDataGridView1[index].Rows[e.RowIndex].Cells[1].Value.ToString();

                for (int i = 0; i < Data_Interface.Reference_Header.Length; i++)
                {
                    if (DB_Interface.Line[0] == Data_Interface.Reference_Header[i])
                    {
                        int Find_DB = 0;
                        if (i > Data_Interface.DB_Column_Limit)
                        {
                            for (int k = 0; k < Data_Interface.DB_Count; k++)
                            {
                                if (i <= Data_Interface.Per_DB_Column_Count_End[k])
                                {
                                    Find_DB = k;
                                    break;
                                }
                            }
                        }

                        DB_Interface.ID = new object[0];
                        DB_Interface.WAFER_ID = new object[0];
                        DB_Interface.LOT_ID = new object[0];
                        DB_Interface.SITE_ID = new object[0];
                        DB_Interface.Value = new object[0];

                        DB_Interface.Dic_Test_For_Spec_Gen = new Dictionary<string, CSV_Class.For_Box>();
                        DB_Interface.Dic_Test = new Dictionary<string, CSV_Class.For_Box>[Data_Interface.DB_Count];

                        DB_Interface.Get_Selected_Para(Data_Interface);

                        foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                        {
                            foreach (KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                            {
                                DB_Interface.Dic_Test_For_Spec_Gen.Add(test2.Key, test2.Value);
                            }

                        }


                        CSV_Class.CSV CSV = new CSV_Class.CSV();
                        CSV_Class.CSV.INT CSV_Interface = CSV.Open(Key);

                        CSV_Interface.Write_Open("C:\\temp\\dummy\\" + Data_Interface.Reference_Header[i] + ".csv");
                        CSV_Interface.Write(DB_Interface.Dic_Test_For_Spec_Gen);
                        CSV_Interface.Write_Close();




                        Dictionary<int, Dictionary<int, string>> OrderbySequence = new Dictionary<int, Dictionary<int, string>>();

                        int Paralen = 0;
                        bool Falg = true;
                        string TextName = "";
                       
                            string[] split = Data_Interface.New_Header[i].Split('_');

                            int kk = 0;
                            foreach (KeyValuePair<int, Dictionary<int, string>> D in Box_Enum)
                            {
                                foreach (KeyValuePair<int, string> S in D.Value)
                                {
                                    string[] dummy1 = new string[0];

                                    if (S.Value == null)
                                    {
                                        //   OrderbySequence.Add(Convert.ToInt16(D.Key), D.Value);
                                    }
                                    else
                                    {
                                        dummy1 = S.Value.Split('_');

                                        if (kk == 0)
                                        {
                                            if (Falg)
                                            {
                                                Paralen = dummy1.Length;

                                                TextName = "";
                                                if (Paralen != 1)
                                                {
                                                    TextName = S.Value;
                                                }
                                                Falg = false;
                                            }

                                            Text = S.Value;
                                        }
                                        if (dummy1.Length == 1)
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
                        N_Spec_Min = new string[1];
                        N_Spec_Max = new string[1];
                        C_Spec_Min = new string[1];
                        C_Spec_Max = new string[1];

                        N_Spec_Min[0] = Convert.ToString(Data_Interface.Clotho_Spcc_List[i].Min[index]);
                        N_Spec_Max[0] = Convert.ToString(Data_Interface.Clotho_Spcc_List[i].Max[index]);
                        C_Spec_Min[0] = Convert.ToString(Data_Interface.Customor_Clotho_List[i].Min[index]);
                        C_Spec_Max[0] = Convert.ToString(Data_Interface.Customor_Clotho_List[i].Max[index]);

                       
                        

                        JMP_Draw("C:\\temp\\dummy\\" + Data_Interface.Reference_Header[i] + ".csv", DB_Interface.Dic_Test_For_Spec_Gen, OrderbySequence, "Distribution" , null, null);

                        DB_Interface.ID = new object[0];
                        DB_Interface.Value = new object[0];
                    }
                }
            }
        }

        private void advancedDataGridView1_SortStringChanged(object sender, EventArgs e)
        {
            int index = tabControl1.SelectedIndex;
            _dataTable[index].DefaultView.Sort = advancedDataGridView1[index].SortString;

            _dataTable[index] = _dataTable[index].DefaultView.ToTable();
            _dataTable[index].PrimaryKey = new DataColumn[] { _dataTable[index].Columns["No"] };
            bindingSource[index].DataSource = _dataTable[index];

            advancedDataGridView1[index].CleanSorts();

        }

        private void advancedDataGridView1_FilterStringChanged(object sender, EventArgs e)
        {
            int index = tabControl1.SelectedIndex;
            bindingSource[index].Filter = advancedDataGridView1[index].FilterString;
        }

        private void advancedDataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int tabControl1_index = tabControl1.SelectedIndex;

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            if (Already_Done_Anly)
            {
                if (ForCtrlz_List_count == 0)
                {
                    ForCtrlz_Min = new List<double[]>();
                    ForCtrlz_Max = new List<double[]>();
                    ForCtrlz_List = new List<forctrlz>[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                    for (int a = 0; a < Data_Interface.Clotho_Spcc_List[0].Max.Length; a++)
                    {
                        ForCtrlz_List[a] = new List<forctrlz>();
                        double[] Min = new double[Data_Interface.Clotho_Spcc_List.Count];
                        double[] Max = new double[Data_Interface.Clotho_Spcc_List.Count];

                        for (int o = 0; o < Data_Interface.Clotho_Spcc_List.Count; o++)
                        {
                            Min[o] = Data_Interface.Clotho_Spcc_List[o].Min[a];
                            Max[o] = Data_Interface.Clotho_Spcc_List[o].Max[a];
                        }
                        ForCtrlz_Min.Add(Min);
                        ForCtrlz_Max.Add(Max);

                    }
                    ForCtrlz_List_count++;
                }
                if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
                {
                    double Testtime9 = TestTime1.Elapsed.TotalMilliseconds;

                    string Falg = advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    if (Falg != "")
                    {
                        int Index = Convert.ToInt16(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[0].Value);
                        string Parameter = Convert.ToString(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[1].Value);
                        double ChangedData = Convert.ToDouble(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                        int MinOrMax = 0; if (e.ColumnIndex == 2) MinOrMax = 1; else MinOrMax = 2;

                        int Header_Length = Data_Interface.Reference_Header.Length;
                        for (int i = 0; i < Header_Length; i++)
                        {
                            if (Parameter == Data_Interface.Reference_Header[i])
                            {
                                int Find_DB = 0;
                                int ColumnLimit = Data_Interface.DB_Column_Limit;
                                if (i + 9 >= ColumnLimit)
                                {
                                    int Db_count = Data_Interface.DB_Count;
                                    for (int k = 0; k < Db_count; k++)
                                    {
                                        int Column_end = Data_Interface.Per_DB_Column_Count_End[k];
                                        if (i + 9 <= Column_end)
                                        {
                                            Find_DB = k;

                                            break;
                                        }
                                    }
                                }

                                double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;
                                //  DB_Interface.Chnaged_Spec_Update_Data(Find_DB, Index, Data_Interface.New_Header[i], ChangedData, MinOrMax);
                                Dictionary<string, double[]> Dic_For_Changed_Spec = DB_Interface.Chnaged_Spec_Anl_Yield(Find_DB, Index, Data_Interface.New_Header[i]);

                                double[] N_Spec = new double[2];

                                for (int a = 0; a < 1; a++)
                                {
                                    N_Spec[0] = Convert.ToDouble(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[2].Value);
                                    N_Spec[1] = Convert.ToDouble(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[3].Value);
                                }

                                double[] N_Data = Dic_For_Changed_Spec["DATA"];


                         //       Dic[Data_Interface.New_Header[i]] = new List<int>();

                                Dic[Data_Interface.Ref_New_Header[i]] = new List<int>();

                                List<int> Count = new List<int>();
                                int List_Count = 0;

                                int Data_Length = N_Data.Length;
                                int Fail_Count = 0;

                                int offset_Raw = 0;


                                offset_Raw = 10;

                                int getNb = Index - ((Data_Interface.DB_Column_Limit) * Find_DB) + offset_Raw;
                                DB_Interface.For_Any_Yield[Find_DB][tabControl1_index][getNb] = N_Data.Length;

                                for (i = 0; i < Data_Length; i++)
                                {

                                    if (N_Spec[1] <= N_Data[i] || N_Spec[0] >= N_Data[i])
                                    {
                                        Fail_Count++;
                                        List_Count = 1;
                                    }


                                    if (List_Count == 0)
                                    {


                                        var itemToRemove = DB_Interface.Yield_Test[Find_DB][i][tabControl1_index].Find(r => r.Row == getNb);
                                        if (itemToRemove != null)
                                        {
                                            DB_Interface.Yield_Test[Find_DB][i][tabControl1_index].Remove(itemToRemove);
                                        }

                                    }
                                    else
                                    {
                                        DB_Interface.For_Any_Yield[Find_DB][tabControl1_index][getNb]--;

                                        var itemToRemove = DB_Interface.Yield_Test[Find_DB][i][tabControl1_index].Find(r => r.Row == getNb);
                                        if (itemToRemove == null)
                                        {
                                            DB_Class.DB_Editing.RowAndPass ss = new DB_Class.DB_Editing.RowAndPass(i, getNb, 1);
                                            DB_Interface.Yield_Test[Find_DB][i][tabControl1_index].Add(ss);
                                        }

                                    }

                                    List_Count = 0;

                                }
                                double Testtime10 = TestTime1.Elapsed.TotalMilliseconds;
                                Count.Add(Fail_Count);


                                TestResult_Cal();

                                //TestResult_Dic[Data_Interface.New_Header[Index + 1]] = Count;
                                //Dic = TestResult_Dic;


                                DataRow dr = _dataTable[tabControl1_index].Rows.Find(Index);
                                int SelRow = _dataTable[tabControl1_index].Rows.IndexOf(dr);

                                double Min = 0f;
                                double Max = 0f;
                                double Avg = 0f;
                                double L_CPK = 0f;
                                double H_CPK = 0f;
                                double Worst_CPK = 0f;
                                double Median = 0f;
                                double Stdev = 0f;
                                //       int Fail = 0;

                                double Testtime0 = TestTime1.Elapsed.TotalMilliseconds;

                                STD(N_Data, N_Spec, out Min, out Max, out Avg, out L_CPK, out H_CPK, out Median, out Stdev);

                                double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;

                                if (L_CPK > H_CPK) Worst_CPK = H_CPK;
                                else Worst_CPK = L_CPK;


                                _dataTable[tabControl1_index].Rows[SelRow][4] = Min;
                                _dataTable[tabControl1_index].Rows[SelRow][5] = Median;
                                _dataTable[tabControl1_index].Rows[SelRow][6] = Max;
                                _dataTable[tabControl1_index].Rows[SelRow][7] = Worst_CPK;
                                _dataTable[tabControl1_index].Rows[SelRow][8] = Stdev;


                                double Testtime11 = TestTime1.Elapsed.TotalMilliseconds;
                                //  Cal_Yield2(Sample - ForGross_Fail_Unit.Count);


                                Sample_Verify = new int[Data_Interface.Clotho_Spcc_List[0].Max.Length];

                                Cal_No_Thread(N_Data.Length - ForGross_Fail_Unit.Count);

                                double Testtime12 = TestTime1.Elapsed.TotalMilliseconds;


                                double Yiled = 0f;
                                try
                                {
                                    Yiled = ((Convert.ToDouble(Data_Length) - Convert.ToDouble(Fail_Count)) / Convert.ToDouble(Data_Length)) * 100;
                                }
                                catch
                                {
                                    Yiled = 0;
                                }
                                _dataTable[tabControl1_index].BeginLoadData();
                                _dataTable[tabControl1_index].Rows[SelRow][9] = Math.Round(Yiled, 2);
                                _dataTable[tabControl1_index].Rows[SelRow][10] = Fail_Count;

                                double Value = 0f;


                                if (e.ColumnIndex == 2)
                                {
                                    Value = ChangedData;

                                    ForCtrlz_Min[tabControl1_index][Index] = Convert.ToDouble(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                                    Cz = new forctrlz(Convert.ToInt16(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[0].Value), e.ColumnIndex, e.RowIndex, Data_Interface.Clotho_Spcc_List[Index].Min[tabControl1_index]);
                                    Data_Interface.Clotho_Spcc_List[Index + 1].Min[tabControl1_index] = Value;
                                    ForCtrlz_List[tabControl1_index].Add(Cz);
                                }
                                else
                                {
                                    Value = ChangedData;
                                    ForCtrlz_Max[tabControl1_index][Index] = Convert.ToDouble(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                                    Cz = new forctrlz(Convert.ToInt16(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[0].Value), e.ColumnIndex, e.RowIndex, Data_Interface.Clotho_Spcc_List[Index].Max[tabControl1_index]);
                                    Data_Interface.Clotho_Spcc_List[Index + 1].Max[tabControl1_index] = Value;
                                    ForCtrlz_List[tabControl1_index].Add(Cz);
                                }


                                _dataTable[tabControl1_index].Rows[SelRow][e.ColumnIndex] = Convert.ToDouble(advancedDataGridView1[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

                                bindingSource[tabControl1_index].DataSource = _dataTable[tabControl1_index];
                                advancedDataGridView1[tabControl1_index].Update();

                                _dataTable[tabControl1_index].EndLoadData();

                                double Testtime15 = TestTime1.Elapsed.TotalMilliseconds;
                                break;

                            }
                        }
                    }
                }
            }
            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }

        private void advancedDataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            if (e.Control && e.KeyCode == Keys.C)
            {
                DataObject Do = advancedDataGridView1[index].GetClipboardContent();
                Clipboard.SetDataObject(Do);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int row = advancedDataGridView1[index].CurrentCell.RowIndex;
                int col = advancedDataGridView1[index].CurrentCell.ColumnIndex;
                foreach (string line in lines)
                {
                    if (row < advancedDataGridView1[index].RowCount && line.Length > 0)
                    {
                        string[] cells = line.Split('\t');
                        for (int i = 0; i < cells.GetLength(0); ++i)
                        {
                            if (col + i < advancedDataGridView1[index].ColumnCount)
                            {
                                advancedDataGridView1[index][col + i, row].Value =
                                Convert.ChangeType(cells[i], advancedDataGridView1[index][col + i, row].ValueType);
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
                bindingSource[index].DataSource = _dataTable[index];
                advancedDataGridView1[index].Update();
            }
            else if (e.Control && e.KeyCode == Keys.Z)
            {
                if (ForCtrlz_List != null)
                {
                    if (ForCtrlz_List[index].Count != 0)
                    {

                        Cz = ForCtrlz_List[index][ForCtrlz_List[index].Count - 1];

                        DataRow dr = _dataTable[index].Rows.Find(Cz.No);
                        int SelRow = _dataTable[index].Rows.IndexOf(dr);

                        advancedDataGridView1[index][Cz.Col, Cz.Row].Value = Cz.Ref_Value;

                        _dataTable[index].Rows[SelRow][Cz.Col] = advancedDataGridView1[index][Cz.Col, Cz.Row].Value;
                        bindingSource[index].DataSource = _dataTable[index];
                        advancedDataGridView1[index].Update();
                        ForCtrlz_List[index].RemoveAt(ForCtrlz_List[index].Count - 1);

                    }
                }

            }
            else if (e.KeyCode == Keys.F1)
            {
                DataGridViewCell currentCell;

                int Tab = tabControl1.SelectedIndex;
                int Tab1 = tabControl1.SelectedIndex;

                currentCell = advancedDataGridView1[Tab].CurrentCell;


                if (Tab != BIN.Length - 1)
                {
                    if (Tab == _dataTable.Length - 1)
                    {
                        Tab = -1;
                    }

                    tabControl1.SelectedIndex = Tab + 1;
                    advancedDataGridView1[Tab1].CurrentCell = currentCell;
                    advancedDataGridView1[tabControl1.SelectedIndex].Visible = true;
                    advancedDataGridView1[tabControl1.SelectedIndex].Focus();
                }
                else
                {

                    tabControl1.SelectedIndex = 0;
                    advancedDataGridView1[Tab1].CurrentCell = currentCell;
                    advancedDataGridView1[tabControl1.SelectedIndex].Visible = true;
                    advancedDataGridView1[tabControl1.SelectedIndex].Focus();
                }

            }
        }

        private void advancedDataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            if (e.Button == MouseButtons.Right && e.ColumnIndex == 1)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                clickedCell = (sender as DataGridView).Rows[e.RowIndex].Cells[e.ColumnIndex];
                this.advancedDataGridView1[index].CurrentCell = clickedCell;
                var relativeMousePosition = advancedDataGridView1[index].PointToClient(Cursor.Position);

                m.Items.Add("Delete Units");
                m.Items.Add("Close All Windows");
                m.Items.Add(new ToolStripSeparator());

                m.Items.Add("Distributions");
                m.Items.Add("Fit Y X Lot");
                m.Items.Add("Fit Y X Site");
                m.Items.Add("BoxPlot");
                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Analyze");


             


                m.ItemClicked += new ToolStripItemClickedEventHandler(m_ItemClicked);

                m.Show(advancedDataGridView1[index], relativeMousePosition);

            }
        }

        public void m_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            CSV_Class.CSV CSV = new CSV_Class.CSV();
            CSV_Class.CSV.INT CSV_Interface = CSV.Open(Key);
            Distribution_Form Dist;
            bool Delete_Flag = false;
            string Query = "";
            string CellValue = "";
       

            DB_Interface.ID = new object[0];
            DB_Interface.WAFER_ID = new object[0];
            DB_Interface.LOT_ID = new object[0];
            DB_Interface.SITE_ID = new object[0];
            DB_Interface.Value = new object[0];

            Dic_Test = new Dictionary<string, CSV_Class.For_Box>[Data_Interface.DB_Count];

            DataObject Do = advancedDataGridView1[index].GetClipboardContent();
            Clipboard.SetDataObject(Do);
            string s = Clipboard.GetText();
            string Text = "";
            string[] X = new string[1];
            char[] option = new char[2];
            bool Flag = false;
            string[] split = new string[0];


            option[1] = '\n';
            option[0] = '\r';

            string[] lines = new string[0];

            switch (e.ClickedItem.Text)
            {
                #region Close Windows

                case "Close All Windows":

                    JMP_Interface.CloseWindowas();

                    break;

                #endregion

                #region Delete Units

                case "Delete Units":

                    bool falg = JMP_Interface.CheckDoc();
                    JMP_Interface.GetSelect_DataTable("dummy");

                    object Units = JMP_Interface.GetSelected_Gross_Row();


                    if (Units != null && Units != "")
                    {


                        List<string> Sb = new List<string>();

                        foreach (object nb in (Array)Units)
                        {
                            Sb.Add(Convert.ToString(nb));
                        }
                        DB_Interface.Gross_Update_Datas(Sb);

                        Hidden_Sample_Count = 0;

                        foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                        {

                            Dictionary<string, List<string>> tests = key.Value;

                            foreach (KeyValuePair<string, List<string>> ts in tests)
                            {
                                Query = "Select count(id) from " + key.Key + " where Fail like '1'";
                                string[] HiddenSample = DB_Interface.Get_Data_By_Query(Query);

                                Hidden_Sample_Count += Convert.ToInt16(HiddenSample[0]);

                            }
                        }


                        JMP_Interface.Close_Dt("");

                        int Len = ((Array)Units).Length;
                        Fail_Units = new string[0];

                        foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
                        {

                            Dictionary<string, List<string>> tests = key.Value;

                            foreach (KeyValuePair<string, List<string>> ts in tests)
                            {
                                Query = "Select id from " + key.Key + " where Fail like '1'";

                                string[] datas = DB_Interface.Get_Data_By_Query(Query);


                                Fail_Units = Fail_Units.Concat(datas).ToArray();

                            }
                        }



                        Cal_No_Thread_For_Delete_Unit(Sample - Fail_Units.Length);

                    }

             //       JMP_Interface.CloseWindowas();
                    break;

                    #endregion

                #region Distribution


                case "Fit Y X Lot":
                case "Fit Y X Site":
                case "Distributions":

                    Flag = false;
                    CellValue = advancedDataGridView1[index].Rows[clickedCell.RowIndex].Cells[1].Value.ToString();

                    Do = advancedDataGridView1[index].GetClipboardContent();
                    Clipboard.SetDataObject(Do);
                    s = Clipboard.GetText();
                    Text = "";

                    option = new char[2];
                    option[1] = '\n';
                    option[0] = '\r';

          


                    if (e.ClickedItem.Text == "Fit Y X:Lot")
                    {
                        X[0] = ":LOT";
                    }
                    else
                    {
                        X[0] = ":SITE";
                    }

                    CSV_Interface.Write_Open("C:\\temp\\dummy\\" + e.ClickedItem.Text + ".csv");

                    DB_Interface.Line = s.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);


                    DB_Interface.Get_Selected_Para(Data_Interface);

                    DB_Interface.Dic_Test_For_Spec_Gen = new Dictionary<string, CSV_Class.For_Box>();


                    foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                    {
                        foreach(KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                        {
                            DB_Interface.Dic_Test_For_Spec_Gen.Add(test2.Key, test2.Value);
                        }

                    }
 
         


                    CSV_Interface.Write(DB_Interface.Dic_Test_For_Spec_Gen);

                    CSV_Interface.Write_Close();

                 //   DB_Interface.Dic_Test_For_Spec_Gen = Dic_Test_For_Spec_Gen;

                    Ordersequence_Method();

                    JMP_Draw("C:\\temp\\dummy\\" + e.ClickedItem.Text + ".csv", DB_Interface.Dic_Test_For_Spec_Gen, OrderbySequence, e.ClickedItem.Text, X, null);



                    break;
                #endregion

                #region BoxPlot

                case "BoxPlot":

                    Do = advancedDataGridView1[index].GetClipboardContent();
                    Clipboard.SetDataObject(Do);
                    s = Clipboard.GetText();
                    Text = "";

                    option = new char[2];
                    option[1] = '\n';
                    option[0] = '\r';


                    DB_Interface.Line = s.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

     
                    CSV_Interface.Write_Open("C:\\temp\\dummy\\Boxplot.csv");


                    DB_Interface.Get_Selected_Para(Data_Interface);

                    DB_Interface.Dic_Test_For_Spec_Gen = new Dictionary<string, CSV_Class.For_Box>();


                    foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                    {
                        foreach (KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                        {
                            DB_Interface.Dic_Test_For_Spec_Gen.Add(test2.Key, test2.Value);
                        }

                    }

                    DB_Interface.Dic_Test = Dic_Test;

                    Ordersequence_Method();


                    if (OrderbySequence.Count != 0)
                    {
                        CSV_Interface.ForBoxplotWrite("", DB_Interface.ID, DB_Interface.Dic_Test_For_Spec_Gen, "");
                        CSV_Interface.Write_Close();

                   //     JMP_Draw("C:\\temp\\dummy\\Boxplot1.csv","", "", "", "","","", null, Dic_Test, OrderbySequence, "BOX" , false);

                        JMP_Draw("C:\\temp\\dummy\\" + e.ClickedItem.Text + ".csv", DB_Interface.Dic_Test_For_Spec_Gen, OrderbySequence, e.ClickedItem.Text, X ,null);

                        DB_Interface.ID = new object[0];
                        DB_Interface.Value = new object[0];
                    }
                    else
                    {
                        CSV_Interface.Write_Close();
                    }

                    PPTX_Interface = PPTX.Opened("YIELD");

                    //PPTX_Interface.Open("C:\\Automation\\PPTX\\PPTX.pptx");

                    //PPTX_Interface.Slide(0);

                    //PPTX_Interface.Title(Text, "", 40);
                    break;
                #endregion

                #region Anlayze

                case "Analyze":

          
                    Distribition_Select_Spec_Method();
                    Dist = new Distribution_Form(Data_Interface, DB_Interface, JMP_Interface, Box_Enum, Data_Interface.Reference_Header, Data_Interface.New_Header, Customer_enable, NPI_enable, CPK_enable, CPK_Value, ref Delete_Flag, "Dist");


                    break;

                    #endregion


            }


        }

        public void Event(object units)
        {

        }

        public void ForeColor()
        {
            for (int s = 0; s < Data_Interface.Clotho_Spcc_List[0].Max.Length; s++)
            {
                string[] No_Index = new string[advancedDataGridView1[s].RowCount];

                int advancedDataGridView1_RowCount = advancedDataGridView1[s].RowCount;
                for (int k = 0; k < advancedDataGridView1_RowCount; k++)
                {
                    No_Index[k] = advancedDataGridView1[s].Rows[k].Cells[1].Value.ToString();

                    string[] ParanameSplit = No_Index[k].Split('_');


                    if (ParanameSplit[ParanameSplit.Length - 1].Contains('-'))
                    {
                        DataGridViewRow rowStyle = advancedDataGridView1[s].Rows[k];

                        rowStyle.DefaultCellStyle.BackColor = Color.Red;
                        //   advancedDataGridView1.Rows[k].DefaultCellStyle.BackColor = Color.Red;

                    }
                }
            }
        }

        public void ForeColor_New_Spec()
        {


            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                tabControl1.SelectedIndex = s;
                MakeSpec.advanced[s].Visible = false;

                string[] No_Index = new string[MakeSpec.advanced[s].RowCount];

                int advancedDataGridView1_RowCount = MakeSpec.advanced[s].RowCount;
                for (int k = 0; k < advancedDataGridView1_RowCount; k++)
                {
                    No_Index[k] = MakeSpec.advanced[s].Rows[k].Cells[1].Value.ToString();

                    string[] ParanameSplit = No_Index[k].Split('_');


                    if (ParanameSplit[ParanameSplit.Length - 1].Contains('-'))
                    {
                        DataGridViewRow rowStyle = MakeSpec.advanced[s].Rows[k];

                        //    rowStyle.DefaultCellStyle.BackColor = Color.Red;
                        MakeSpec.advanced[s].Rows[k].DefaultCellStyle.BackColor = Color.Red;
                        //  MakeSpec.advanced[s].Update();
                    }
                }

                MakeSpec.advanced[s].Visible = true;
                //  MakeSpec.advanced[s].EndEdit();

            }
        }

        public void For_New_Spec_ForeColor()
        {
            //string[] No_Index = new string[advancedDataGridView1.RowCount];

            //int advancedDataGridView1_RowCount = advancedDataGridView1.RowCount;
            //for (int k = 0; k < advancedDataGridView1_RowCount; k++)
            //{
            //    No_Index[k] = advancedDataGridView1.Rows[k].Cells[1].Value.ToString();

            //    string[] ParanameSplit = No_Index[k].Split('_');


            //    if (ParanameSplit[ParanameSplit.Length - 1].Contains('-'))
            //    {
            //        DataGridViewRow rowStyle = advancedDataGridView1.Rows[k];

            //        rowStyle.DefaultCellStyle.BackColor = Color.Red;
            //        //   advancedDataGridView1.Rows[k].DefaultCellStyle.BackColor = Color.Red;

            //    }
            //}
        }

        private void SetSpec()
        {

            //if (checkBox2.Checked)
            //{
            //    ForNewSpec = true;
            //    DB_Interface.Get_Saved_Spec(Data_Interface);

            //    Data_Class.Data_Editing.ForAnl_NewMinSpec = new string[advancedDataGridView1.RowCount + Data_Interface.TheEnd_Trashes_Header_Count + Data_Interface.TheFirst_Trashes_Header_Count + 1];
            //    Data_Class.Data_Editing.ForAnl_NewMaxSpec = new string[advancedDataGridView1.RowCount + Data_Interface.TheEnd_Trashes_Header_Count + Data_Interface.TheFirst_Trashes_Header_Count + 1];

            //    Data_Class.Data_Editing.ForAnl_NewMinSpec[0] = "LOW";
            //    Data_Class.Data_Editing.ForAnl_NewMaxSpec[0] = "HIGH";

            //    for (int l = 1; l < Data_Interface.TheFirst_Trashes_Header_Count + 1; l++)
            //    {
            //        Data_Class.Data_Editing.ForAnl_NewMinSpec[l] = "0";
            //        Data_Class.Data_Editing.ForAnl_NewMaxSpec[l] = "0";
            //    }

            //    int p = Data_Interface.TheFirst_Trashes_Header_Count + 1;
            //    int p1 = 1;
            //    for (int ki = 0; ki < DB_Interface.DataSet_Value.Length; ki++)
            //    {
            //        for (int l = 0; l < DB_Interface.DataSet_Value[ki][0].Length - 5; l++)
            //        {
            //            if (ki == 0)
            //            {
            //                Data_Class.Data_Editing.ForAnl_NewMinSpec[p] = DB_Interface.DataSet_Value[ki][0][l + 1];
            //                Data_Class.Data_Editing.ForAnl_NewMaxSpec[p] = DB_Interface.DataSet_Value[ki][1][l + 1];
            //                if (p1 == Data_Interface.DB_Column_Limit - 1)
            //                {
            //                    p++;
            //                    p1++;
            //                    break;
            //                }
            //            }
            //            else
            //            {
            //                Data_Class.Data_Editing.ForAnl_NewMinSpec[p] = DB_Interface.DataSet_Value[ki][0][l];
            //                Data_Class.Data_Editing.ForAnl_NewMaxSpec[p] = DB_Interface.DataSet_Value[ki][1][l];
            //            }
            //            p++;
            //            p1++;

            //        }
            //    }

            //    for (int l = Data_Class.Data_Editing.ForAnl_NewMinSpec.Length - Data_Interface.TheEnd_Trashes_Header_Count; l < Data_Class.Data_Editing.ForAnl_NewMaxSpec.Length; l++)
            //    {
            //        Data_Class.Data_Editing.ForAnl_NewMinSpec[l] = "0";
            //        Data_Class.Data_Editing.ForAnl_NewMaxSpec[l] = "0";
            //    }

            //    int j = 1;
            //    Data_Interface.New_HighSpec[0] = Convert.ToDouble(0);
            //    for (int i = Data_Interface.TheFirst_Trashes_Header_Count + 1; i < Data_Class.Data_Editing.ForAnl_NewMaxSpec.Length - Data_Interface.TheEnd_Trashes_Header_Count; i++)
            //    {

            //        Data_Interface.New_HighSpec[j] = Convert.ToDouble(Data_Class.Data_Editing.ForAnl_NewMaxSpec[i]);
            //        Data_Interface.New_LowSpec[j] = Convert.ToDouble(Data_Class.Data_Editing.ForAnl_NewMinSpec[i]);
            //        j++;
            //    }
            //}

            //else ForNewSpec = false;
        }

        public void STD(double[] Data, double[] Spec, out double Min, out double Max, out double Ave, out double L_CPK, out double H_CPK, out double Median, out double Stdev)
        {
            Min = Data.Min();
            Max = Data.Max();
            Ave = Data.Average();
            L_CPK = 0f;
            H_CPK = 0f;
            Median = 0f;
            Stdev = 0f;

            if (Data.Length % 2 == 0)
            {
                Array.Sort(Data);
                double GetMedian_i = Data[(Data.Length / 2) - 1];
                double GetMedian_j = Data[(Data.Length / 2)];

                Median = (GetMedian_i + GetMedian_j) / 2;
            }
            else
            {
                Array.Sort(Data);
                int GetMedian_i = (Data.Length) / 2;
                Median = Data[GetMedian_i];
            }

            double minusSquareSummary = 0.0;

            foreach (double source in Data)
            {
                minusSquareSummary += (source - Ave) * (source - Ave);
            }

            Stdev = Math.Sqrt(minusSquareSummary / (Data.Length - 1));

            L_CPK = (Ave - Spec[0]) / (3 * Stdev);
            H_CPK = (Spec[1] - Ave) / (3 * Stdev);

        }

        public void TestResult_Cal()
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            DB_Interface.Cal_Value_by_rowsdata = new Dictionary<string, DB_Class.DB_Editing.Data_Calculation>();

            double[] dummy_Test = new double[14];

            for (int j = 0; j < Data_Interface.New_Header.Length; j++)
            {
                DB_Interface.Cal_Value_by_rowsdata.Add(Data_Interface.Ref_New_Header[j], new DB_Class.DB_Editing.Data_Calculation(Data_Interface.SWBIN_Dic.Count));
            }

            double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;
        }

        public void For_New_Spec_TestResult_Cal()
        {

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            if (DB_Interface.For_New_Spec_Cal_Value_by_rowsdata == null)
            {
                DB_Interface.For_New_Spec_Cal_Value_by_rowsdata = new Dictionary<string, DB_Class.DB_Editing.Data_Calculation>();
                double[] dummy_Test = new double[14];

                for (int j = 0; j < Data_Interface.New_Header.Length; j++)
                {
                    DB_Interface.For_New_Spec_Cal_Value_by_rowsdata.Add(Data_Interface.Ref_New_Header[j], new DB_Class.DB_Editing.Data_Calculation(Data_Interface.Clotho_List[0].Max.Length));
                }
            }


            double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;

        }

        public static T[] DupCheck<T>(T[] dupArray)
        {
            List<T> result = new List<T>();

            for (int i = 0; i < dupArray.Length; i++)
            {
                if (result.Contains(dupArray[i])) continue;
                result.Add(dupArray[i]);
            }
            return result.ToArray();
        }

        public void Anlyzer_By_Lot_thread(Object i)
        {
            int DB = (int)i;
            int Count = Data_Interface.Per_DB_Column_Count[DB];

            for (int j = 0; j < databylot.Length; j++)
            {
                for (int k = 0; k < Count; k++)
                {
                    if (DB_Interface.For_Any_Yield[DB][Selected_Bin][k] != 0)
                    {
                        int L = Lot_Dic[databylot[j]];
                        By_Lot[DB][0][L][0][k] += 1;
                        // break;
                    }
                }
            }
            ThreadFlags[DB].Set();

        }

        public void Anlyzer_By_Site_thread(Object i)
        {
            int DB = (int)i;
            int Count = Data_Interface.Per_DB_Column_Count[DB];

            for (int j = 0; j < databylot.Length; j++)
            {
                for (int k = 0; k < Count; k++)
                {
                    if (DB_Interface.For_Any_Yield[DB][Selected_Bin][k] == 1)
                    {
                        int L = Site_Dic[databylot[j]];
                        By_Site[DB][0][L][0][k] += 1;
                    }
                }
            }
            ThreadFlags[DB].Set();

        }

        public void Anlyzer_By_Bin_thread(Object i)
        {

            int DB = (int)i;
            int Count = Data_Interface.Per_DB_Column_Count[DB];

            int Samplecount = databybin.Length;

            for (int q = 0; q < Samplecount; q++)
            {
                for (int j = 0; j < BIN.Length; j++)
                {
                    if (DB_Interface.For_Any_Yield[DB][j][q] == 1)
                    {
                        int B = Convert.ToInt16(databybin[q]);

                        for (int g = 0; g < BIN.Length; g++)
                        {
                            if (Convert.ToString(B) == BIN[g])
                            {
                                By_Bin[DB][0][j][0][q] += 1;
                            }
                        }

                    }
                }

            }
            ThreadFlags[DB].Set();

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        public void Button_Set()
        {
            Analysis = new System.Windows.Forms.Button[2];
            Std = new System.Windows.Forms.Button[2];
            Sort = new System.Windows.Forms.Button[2];
            Save = new System.Windows.Forms.Button[2];


            for (int r = 0; r < 2; r++)
            {
                int Location_width = 10;
                int Location_height = 10;

                int Size_width = 110;
                int Size_height = 25;


                string title = "";
                if (r == 0) title = "Yield";
                else title = "N_Yield";
                TabPage myTabPage = new TabPage(title);
                tabControl2.TabPages.Add(myTabPage);

                Analysis[r] = new System.Windows.Forms.Button();
                Analysis[r].Location = new System.Drawing.Point(Location_width, Location_height);
                Analysis[r].Name = "Yield Analysis";
                Analysis[r].Size = new System.Drawing.Size(Size_width, Size_height);
                Analysis[r].TabIndex = 6;
                Analysis[r].Text = "Redo Yield Analysis";
                Analysis[r].UseVisualStyleBackColor = true;
                Analysis[r].Click += new System.EventHandler(this.button4_Click);

                tabControl2.TabPages[r].Controls.Add(Analysis[r]);


                Location_height = Location_height + 30;

                Std[r] = new System.Windows.Forms.Button();
                Std[r].Location = new System.Drawing.Point(Location_width, Location_height);
                Std[r].Name = "button2";
                Std[r].Size = new System.Drawing.Size(Size_width, Size_height);
                Std[r].TabIndex = 3;
                Std[r].Text = "Std Analysis";
                Std[r].UseVisualStyleBackColor = true;
                Std[r].Click += new System.EventHandler(this.button2_Click);


                tabControl2.TabPages[r].Controls.Add(Std[r]);

                Location_height = Location_height + 30;

                Sort[r] = new System.Windows.Forms.Button();
                Sort[r].Location = new System.Drawing.Point(Location_width, Location_height);
                Sort[r].Name = "button10";
                Sort[r].Size = new System.Drawing.Size(Size_width, Size_height);
                Sort[r].TabIndex = 20;
                Sort[r].Text = "Reset Filter and Sort";
                Sort[r].UseVisualStyleBackColor = true;
                Sort[r].Click += new System.EventHandler(this.button10_Click);

                tabControl2.TabPages[r].Controls.Add(Sort[r]);

                Location_height = Location_height + 30;
                Save[r] = new System.Windows.Forms.Button();
                Save[r].Location = new System.Drawing.Point(Location_width, Location_height);
                Save[r].Name = "button11";
                Save[r].Size = new System.Drawing.Size(Size_width, Size_height);
                Save[r].TabIndex = 21;
                Save[r].Text = "Save Spec";
                Save[r].UseVisualStyleBackColor = true;
                Save[r].Click += new System.EventHandler(this.button11_Click);
                Save[r].Enabled = false;

                tabControl2.TabPages[r].Controls.Add(Save[r]);

                if (r == 1)
                {
                    Location_height = Location_height + 30;

                    MakeSpecB = new Button();
                    MakeSpecB.Location = new System.Drawing.Point(Location_width, Location_height);
                    MakeSpecB.Name = "button3";
                    MakeSpecB.Size = new System.Drawing.Size(Size_width, Size_height);
                    MakeSpecB.TabIndex = 27;
                    MakeSpecB.Text = "Make Spec";
                    MakeSpecB.UseVisualStyleBackColor = true;
                    MakeSpecB.Click += new System.EventHandler(this.button3_Click);
                    //  MakeSpecB.Enabled = false;
                    tabControl2.TabPages[r].Controls.Add(MakeSpecB);


                }
            }


            //  tabControl2.Controls.Remove(tabPage2);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = tabControl1.SelectedIndex;
            tabControl3.SelectedIndex = index;
        }

        private void tabControl3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = tabControl3.SelectedIndex;
            tabControl1.SelectedIndex = index;
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
                else if (Lot[i].ToUpper().Contains("CHAN"))
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

        public void Dic_Method(out int _Outcount, string Query)
        {
            _Outcount = 0;

            string d = "";
            foreach (KeyValuePair<string, Dictionary<string, List<string>>> key in this.Matching_Lots)
            {
                Dictionary<string, List<string>> tests = key.Value;


                foreach (KeyValuePair<string, List<string>> ts in tests)
                {
                    //if (ts.Key == selecteditem.ToString())
                    //{
                    //    d = key.Key;
                    //    DB_Interface.Lot_ID = selecteditem.ToString();
                    //    DB_Interface.Matching_Lot = tests;
                    //    break;
                    //}
                }

            }
        }

        public void Gridview2()
        {

            dataGridView2.ColumnCount = 1;
            dataGridView2.Columns[0].Name = "Bin";


            DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();


            dataGridView2.Columns.Add(chk);

            chk.HeaderText = "Check";


            for (int i = 0; i < Data_Interface.SWBIN_Dic.Count; i++)
            {
                dataGridView2.Rows.Add("");
                dataGridView2.Rows[i].Cells[1].Value = true;
            }


            for (int i = 0; i < Data_Interface.SWBIN_Dic.Count; i++)
            {
                dataGridView2.Rows[i].Cells[0].Value = "Bin " + (i + 1);
            }

            dataGridView2.Visible = true;

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView2.RowHeadersVisible = false;
            dataGridView2.AllowUserToAddRows = false;
            //  dataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            //  dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dataGridView2.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f =>

            {

                f.SortMode = DataGridViewColumnSortMode.NotSortable; // sort 막기


            });

        }

        public void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = tabControl4.SelectedIndex;

            if (index == 0)
            {
                #region
                if (e.ColumnIndex == 2)
                {
                    index = tabControl1.SelectedIndex;

                    advancedDataGridView1[index].CleanFilter();
                    advancedDataGridView1[index].CleanSort();


                    _dataTable[index].DefaultView.Sort = advancedDataGridView1[index].SortString;
                    //  _dataTable[index] = _dataTable[index].DefaultView.ToTable();
                    _dataTable[index].PrimaryKey = new DataColumn[] { _dataTable[index].Columns["No"] };
                    bindingSource[index].DataSource = _dataTable[index];
                    bindingSource[index].Filter = "";

                    advancedDataGridView1[index].DataSource = bindingSource[index];
                    advancedDataGridView1[index].Update();

                    StringBuilder ForFilter = new StringBuilder();


                    if (Datagrid2[index].Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != null && Datagrid2[index].Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != "")
                    {
                        string[] split = Datagrid2[index].Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString().Split(',');

                        for (int i = 0; i < split.Length; i++)
                        {
                            if (i == 0)
                            {
                                ForFilter.Append("([Parameter] LIKE '%_" + Datagrid2[index].Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());
                                ForFilter.Append("_%'");
                                ForFilter.Append(" and [Parameter] LIKE '%_" + split[0] + "_%'");
                            }
                            else if (i == split.Length - 1)
                            {
                                ForFilter.Append(" and [Parameter] LIKE '%_" + split[i].ToString() + "_%'");
                            }
                            else
                            {
                                ForFilter.Append(" and [Parameter] LIKE '%_" + split[i].ToString() + "_%'");
                            }
                        }
                        ForFilter.Append(")");
                        bindingSource[index].Filter = ForFilter.ToString();

                        advancedDataGridView1[index].DataSource = bindingSource[index];
                        advancedDataGridView1[index].Update();
                        ForFilter = new StringBuilder();
                    }
                    else
                    {


                        ForFilter.Append("([Parameter] LIKE '%" + Datagrid2[index].Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());
                        ForFilter.Append("%'");


                        ForFilter.Append(")");

                        bindingSource[index].Filter = ForFilter.ToString(); ;

                        advancedDataGridView1[index].DataSource = bindingSource[index];
                        advancedDataGridView1[index].Update();
                        ForFilter = new StringBuilder();


                    }


                }
                else if (e.ColumnIndex == 3)
                {

                }
                #endregion
            }
            else
            {
                #region
                if (e.ColumnIndex == 1)
                {
                    index = tabControl1.SelectedIndex;

                    advancedDataGridView1[index].CleanFilter();
                    advancedDataGridView1[index].CleanSort();


                    _dataTable[index].DefaultView.Sort = advancedDataGridView1[index].SortString;
                    //  _dataTable[index] = _dataTable[index].DefaultView.ToTable();
                    _dataTable[index].PrimaryKey = new DataColumn[] { _dataTable[index].Columns["No"] };
                    bindingSource[index].DataSource = _dataTable[index];
                    bindingSource[index].Filter = "";

                    advancedDataGridView1[index].DataSource = bindingSource[index];
                    advancedDataGridView1[index].Update();

                    StringBuilder ForFilter = new StringBuilder();


                    if (Datagrid2[1].Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != null)
                    {


                        ForFilter.Append("([Parameter] LIKE '%" + Datagrid2[1].Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString());
                        ForFilter.Append("%'");


                        ForFilter.Append(")");

                        bindingSource[index].Filter = ForFilter.ToString(); ;

                        advancedDataGridView1[index].DataSource = bindingSource[index];
                        advancedDataGridView1[index].Update();
                        ForFilter = new StringBuilder();


                    }


                }
                #endregion
            }

        }

        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {

            Stopwatch TestTime2 = new Stopwatch();
            TestTime2.Restart();
            TestTime2.Start();

            if (e.Control && e.KeyCode == Keys.C)
            {

                DataObject Do = dataGridView3.GetClipboardContent();
                Clipboard.SetDataObject(Do);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                string[] A = new string[lines.Length];


                for (int k = 0; k < dataGridView3.SelectedCells.Count; k++)
                {
                    // advanced[index].CellBeginEdit(true);

                    int row = dataGridView3.SelectedCells[k].RowIndex;
                    int col = dataGridView3.SelectedCells[k].ColumnIndex;

                    int a = 0;

                    string[] d = s.Split(new char[] { '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                    string[] t = s.Split('\t');
                    string[] r = s.Split('\r');
                    string[] n = s.Split('\n');


                    for (int i = 0; i < d.Length; i++)
                    {
                        dataGridView3[col, row].Value = d[i];
                    }
                    //foreach (string line in lines)
                    //{




                    //    if (line.Contains('\t'))
                    //    {

                    //        dataGridView3[col + ColOffset, row].Value = d[a];
                    //        ColOffset++;
                    //    }
                    //    if (line.Contains('\r'))
                    //    {
                    //        RowOffset++;
                    //    }




                    //    for(int i = 0; i < d.Length; i ++)
                    //    {

                    //    }

                    //    a++;
                    //}
                }

            }
            double Testtime77 = TestTime2.Elapsed.TotalMilliseconds;


        }

        public void FindOrderbyParameter(Dictionary<string, CSV_Class.For_Box> Dic_Test, KeyValuePair<int, Dictionary<int, string>> OrderbySequence)
        {


        }

        private void Yield_Cal_Form_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            this.Enabled = false;

            int index = tabControl1.SelectedIndex;


            Re_GridView2();
            

            advancedDataGridView1[index].CleanFilters();
            advancedDataGridView1[index].CleanSorts();

            _dataTable[index].DefaultView.RowFilter = "";
            _dataTable[index].DefaultView.Sort = "[No] ASC";
            bindingSource[index].DataSource = _dataTable[index];
            bindingSource[index].Filter = "";


            Data_Interface.Data_Table = "Current_Setting";
            DB_Interface.DropTable(Data_Interface, "drop table Current_Setting");

            for (int d = 0; d < Data_Interface.Clotho_Spcc_List[0].Max.Length; d++)
            {
                DB_Interface.No_Index = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.Paraname = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.SpecMin = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.SpecMax = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.DataMin = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.DataMedian = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.DataMax = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.CPK = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.STD = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.Percent = new string[advancedDataGridView1[d].RowCount + 1];
                DB_Interface.Fail = new string[advancedDataGridView1[d].RowCount + 1];

                DB_Interface.No_Index[0] = "0";
                DB_Interface.Paraname[0] = "0";
                DB_Interface.SpecMin[0] = "0";
                DB_Interface.SpecMax[0] = "0";
                DB_Interface.DataMin[0] = "0";
                DB_Interface.DataMedian[0] = "0";
                DB_Interface.DataMax[0] = "0";
                DB_Interface.CPK[0] = "0";
                DB_Interface.STD[0] = "0";
                DB_Interface.Percent[0] = "0";
                DB_Interface.Fail[0] = "0";


                for (int k = 1; k < advancedDataGridView1[d].RowCount + 1 ; k++)
                {
                    DB_Interface.No_Index[k] = advancedDataGridView1[d].Rows[k - 1].Cells[0].Value.ToString();
                    DB_Interface.Paraname[k] = advancedDataGridView1[d].Rows[k - 1].Cells[1].Value.ToString();
                    DB_Interface.SpecMin[k] = advancedDataGridView1[d].Rows[k - 1].Cells[2].Value.ToString();
                    DB_Interface.SpecMax[k] = advancedDataGridView1[d].Rows[k - 1].Cells[3].Value.ToString();
                    DB_Interface.DataMin[k] = advancedDataGridView1[d].Rows[k - 1].Cells[4].Value.ToString();
                    DB_Interface.DataMedian[k] = advancedDataGridView1[d].Rows[k - 1].Cells[5].Value.ToString();
                    DB_Interface.DataMax[k] = advancedDataGridView1[d].Rows[k - 1].Cells[6].Value.ToString();
                    DB_Interface.CPK[k] = advancedDataGridView1[d].Rows[k - 1].Cells[7].Value.ToString();
                    DB_Interface.STD[k] = advancedDataGridView1[d].Rows[k - 1].Cells[8].Value.ToString();
                    DB_Interface.Percent[k] = advancedDataGridView1[d].Rows[k - 1].Cells[9].Value.ToString();
                    DB_Interface.Fail[k] = advancedDataGridView1[d].Rows[k - 1].Cells[10].Value.ToString();
                }

                if(d == 0) DB_Interface.Insert_Current_Setting(Data_Interface);

      

                DB_Interface.Insert_Current_Setting_Data(Data_Interface, Convert.ToString(d));
                Lastest_Setting_Cal(false);
            }
        }

        public void Write_Inf(long SampleCount)
        {
            bool flag = false;
            int Db = 0;

            string Filename = DB_Interface.Filename.Substring(DB_Interface.Filename.LastIndexOf("\\") + 1);

            int length = DB_Interface.Filename.Length;
            Filename = DB_Interface.Filename.Substring(0, length - Filename.Length);

            Csv_Interface.Write_Open(Filename + "Inf.csv");
            Csv_Interface.Write("SampleCount," + SampleCount);

            string lot = "";

            for(int h = 0; h < Lot.Length; h ++)
            {
                if( h == Lot.Length - 1)
                {
                    lot += Lot[h];
                }
                else
                {
                    lot += Lot[h] + ",";
                }
             
            }
            Csv_Interface.Write("LOT," + lot);
            Csv_Interface.Write("BinCount," + Data_Interface.Clotho_Spcc_List[0].Max.Length);

            for (int j = 0; j < Data_Interface.Clotho_Spcc_List[0].Max.Length; j++)
            {

                Csv_Interface.Write("Total_Sample," + Datagrid[j].Rows[0].Cells[1].Value.ToString());
                Csv_Interface.Write("Analysis_Sample," + Datagrid[j].Rows[1].Cells[1].Value.ToString());
                Csv_Interface.Write("Pass," + Datagrid[j].Rows[2].Cells[1].Value.ToString());
                Csv_Interface.Write("Fail," + Datagrid[j].Rows[3].Cells[1].Value.ToString());
                Csv_Interface.Write("Percent," + Datagrid[j].Rows[4].Cells[1].Value.ToString());
                Csv_Interface.Write("Hidden," + Datagrid[j].Rows[5].Cells[1].Value.ToString());


                Csv_Interface.Write("BIN:" + (j + 1));

                for (int n = 0; n < SampleCount; n++)
                {
                    //    StringBuilder Apped = new StringBuilder();
                    bool flag_Test = false;
                    bool db_flag = true;
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {
                        for (int m = 0; m < DB_Interface.Yield_Test[Db][n][j].Count; m++)
                        {
                            List<DB_Class.DB_Editing.RowAndPass> s = new List<DB_Class.DB_Editing.RowAndPass>();
                            s = DB_Interface.Yield_Test[Db][n][j];
                            if (db_flag)
                            {
                                Csv_Interface.Write_For_Result(n + "," + s[0].SN + ",");
                                db_flag = false;
                            }

                            if (n == 257)
                            {

                            }

                            for (int k = 0; k < s.Count; k++)
                            {
                                if(k != s.Count - 1)
                                {
                                 //   Csv_Interface.Write_For_Result(Convert.ToString(s[k].Row)  + ",");
                                    Csv_Interface.Write_For_Result(Convert.ToString(s[k].Row + Data_Interface.DB_Column_Limit * Db) + ",");
                 
                                }
                                else
                                {
                                 //   Csv_Interface.Write_For_Result(Convert.ToString(s[k].Row)  + ",");
                                    Csv_Interface.Write_For_Result(Convert.ToString(s[k].Row + (Data_Interface.DB_Column_Limit * Db)) + ",");
               
                                }
                               
                            }
                            flag_Test = true;
                            break;

                        }
                    }
                    if(flag_Test)
                    {
                        Csv_Interface.Write("");
                    }

                }
             
            }

            Csv_Interface.Write_Close();
        }

        public void Ordersequence_Method()
        {
            OrderbySequence = new Dictionary<int, Dictionary<int, string>>();

            bool Falg = true;
            int Paralen = 0;
            string TextName = "";

            foreach (KeyValuePair<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test_For_Spec_Gen)
            {
                bool Flag_T = false;
                string[] split = test.Key.Split('_');

                int kk = 0;
                foreach (KeyValuePair<int, Dictionary<int, string>> D in Box_Enum)
                {
                    foreach (KeyValuePair<int, string> S in D.Value)
                    {
                        string[] dummy1 = new string[0];

                        if (S.Value == null)
                        {
                            //   OrderbySequence.Add(Convert.ToInt16(D.Key), D.Value);
                        }
                        else
                        {
                            dummy1 = S.Value.Split('_');

                            if (kk == 0)
                            {
                                if (Falg)
                                {
                                    Paralen = dummy1.Length;

                                    TextName = "";
                                    if (Paralen != 1)
                                    {
                                        TextName = S.Value;
                                    }
                                    Falg = false;
                                }

                                Text = S.Value;
                            }
                            if (dummy1.Length == 1)
                            {
                                if (S.Value.ToString().ToUpper() == split[1].ToUpper())
                                {
                                    if (!OrderbySequence.ContainsKey(Convert.ToInt16(D.Key)))
                                    {
                                        OrderbySequence.Add(Convert.ToInt16(D.Key), D.Value);
                                    }

                                    Flag_T = true;
                                }
                            }
                            else
                            {

                                if (S.Value.ToString().ToUpper() == split[1].ToUpper() + "_" + split[2].ToUpper())
                                {
                                    if (!OrderbySequence.ContainsKey(Convert.ToInt16(D.Key)))
                                        OrderbySequence.Add(Convert.ToInt16(D.Key), D.Value);
                                }
                            }
                        }
                        if (kk == 0)
                            break;
                        kk++;
                    }
                    if (Flag_T)
                        break;
                }

            }
        }

        public void Distribition_Select_Spec_Method()
        {
            Customer_enable = false;
            NPI_enable = false;
            CPK_enable = false;
            CPK_Value = 0f;

            if (tabControl5.SelectedIndex == 0)
            {
                if (radioButton1.Checked) Customer_enable = true;
                else if (radioButton2.Checked) NPI_enable = true;
                else if (radioButton3.Checked)
                {
                    CPK_enable = true;
                    if (textBox1.Text == "")
                        CPK_Value = Convert.ToDouble(1.5);
                    else
                        CPK_Value = Convert.ToDouble(textBox1.Text);
                }
            }
        }

        public class forctrlz
        {
            public int No;
            public int Col;
            public int Row;
            public double Ref_Value;

            public forctrlz(int No, int Col, int Row, double Ref_Value)
            {
                this.No = No;
                this.Col = Col;
                this.Row = Row;
                this.Ref_Value = Ref_Value;
            }
        }
        public class Gross
        {
            public double Range;
            public string Selector;
            public Gross(double Range, string Selector)
            {
                this.Range = Range;
                this.Selector = Selector;
            }
        }
        struct Bin_Struct
        {
            public int No;
            public string Para;
            public double Spec_min;
            public double Spec_max;
            public double Data_min;
            public double Data_median;
            public double Data_max;
            public double CPK;
            public double Std;
            public double Pecent;
            public int Fail;
            public string Null;
            public double N_Spec_min;
            public double N_Spec_max;

            public Bin_Struct(int No, string Para, double Spec_min, double Spec_Max, double Data_min, double Data_median, double Data_max, double CPK, double Std, double Pecent, int Fail, string Null, double N_Spec_min, double N_Spec_max)
            {
                this.No = No;
                this.Para = Para;
                this.Spec_min = Spec_min;
                this.Spec_max = Spec_Max;
                this.Data_min = Data_min;
                this.Data_median = Data_median;
                this.Data_max = Data_max;
                this.CPK = CPK;
                this.Std = Std;
                this.Pecent = Pecent;
                this.Fail = Fail;
                this.Null = Null;
                this.N_Spec_min = N_Spec_min;
                this.N_Spec_max = N_Spec_max;
            }
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

        private void button5_Click(object sender, EventArgs e)
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
                if(!Trace[k].Contains("CHAN"))
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

            for(int kk = 0; kk< Trace.Length; kk++)
            {
                _Trace[kk] = Convert.ToInt16(Trace[kk]);
            }

            Array.Sort(_Trace);
            //    double[] doubles = Array.ConvertAll<object, double>(DataValue, Convert.ToDouble);

            ATE.SPARA_Form t = new ATE.SPARA_Form(DB_Interface, _Trace);
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }
    }

    public static class ExtensionMethod
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.SetProperty);
            pi.SetValue(dgv, setting, null);
        }
    }


}
