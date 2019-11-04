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
using System.Diagnostics;
using System.Threading;

namespace TestApplication
{
    public partial class MakeSpec_Form : Form
    {

        public DataTable[] _dataTable;
        DataSet _dataSet = new DataSet();
        public BindingSource[] bindingSource;

        public DataTable[] _dataTable_Spec;
        public BindingSource[] bindingSource_Spec;

        DataTable _dataTable_Para = new DataTable();
        DataSet _dataSet_Para = new DataSet();
        BindingSource bindingSource_Para = new BindingSource();

        int i = 0;

        object[] Valuse;
        Data_Class.Data_Editing.INT Data_Interface;
        DB_Class.DB_Editing.INT Db_Interface;
        CSV_Class.CSV.INT Csv_Interface;
        JMP_Class.JMP_Editing.INT JMP_Interface;


        List<string> Split_Para;
        string JMP_File;
        string Key = "Yield";

        public Zuby.ADGV.AdvancedDataGridView[] advanced;
        public DataGridView[] datagrid;
        public DataGridView[] datagrid2;

        Dictionary<string, List<forctrlz>> A_ForCtrlz_Dic = new Dictionary<string, List<forctrlz>>();

        int A_ForCtrlz_List_count = 0;


        Dictionary<string, List<forctrlz>> ForCtrlz_Dic = new Dictionary<string, List<forctrlz>>();
        List<forctrlz>[] ForCtrlz_List;
        List<double[]> ForCtrlz_Min;
        List<double[]> ForCtrlz_Max;
        //   forctrlz Cz;
        int ForCtrlz_List_count = 0;


        long Total;
        long Any_Total;
        long Hidden_Total;
        DataGridViewCell currentCell;
        DataGridViewCell clickedCell;

        int Bin_Length;
        int Col;
        int Row;
        bool Flag;

        StringBuilder ForFilter;
        delegate void SetComboBoxCellType(int iRowIndex);

        bool bIsComboBox = false;

        long[] Calculate_thread_Strat;
        long[] Calculate_thread_End;
        int[] Sample_Verify;

        ManualResetEvent[] For_Cal;
        bool[] Wait;
        List<int[]>[] List_Sample_Verify;

        //   string[] Define_Spec;

        ManualResetEvent[] ThreadFlags;

        static long Hidden_Sample_Count;
        string[] Fail_Units;

        public List<string> Outlier_List;

        public Dictionary<string, List<int>> Dic_For_Bin = new Dictionary<string, List<int>>();

        public MakeSpec_Form(Data_Class.Data_Editing.INT Data_Interface, DB_Class.DB_Editing.INT Db_Interface, CSV_Class.CSV.INT Csv_Interface, JMP_Class.JMP_Editing.INT JMP_Interface, long Total, long Any_Total, long Hidden_Total, List<string> Outlier_List)
        {

            InitializeComponent();
            this.Total = Total;
            this.Any_Total = Any_Total;
            this.Hidden_Total = Hidden_Total;
            this.Data_Interface = Data_Interface;
            this.Db_Interface = Db_Interface;
            this.Csv_Interface = Csv_Interface;
            this.JMP_Interface = JMP_Interface;
            this.Bin_Length = Data_Interface.Clotho_List[0].Max.Length;


            if (MessageBox.Show("Do you want to load existing Spec?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {

                For_New_Spec_TestResult_Cal();


                for (int i = 0; i < Data_Interface.Customor_Clotho_List[0].Max.Length; i++)
                {
                    Data_Interface.Data_Table = "Table" + i;
                    Db_Interface.Insert_Spec_Get_From_DB(Data_Interface);
                }
                LoadSpec();
            }
            else
            {
                for (int i = 0; i < Data_Interface.Customor_Clotho_List[0].Max.Length; i++)
                {
                    Db_Interface.DropTable(Data_Interface, "drop table Table" + i);
                }


                for (int i = 0; i < Data_Interface.Customor_Clotho_List[0].Max.Length; i++)
                {
                    Data_Interface.Data_Table = "Table" + i;
                    Db_Interface.Insert_Spec_Header(Data_Interface);
                }

                MakeSpec();
            }


            this.Show();
            View();


        }

        public void MakeSpec()
        {

            #region

            Split_Para = new List<string>();

            for (i = 1; i < Data_Interface.New_Header.Length; i++)
            {
                string[] Split = Data_Interface.Reference_Header[i].Split('_');

                // if (Split[0].ToUpper().ToString() != "M")
                // {
                if (!Split_Para.Contains(Split[1].ToUpper()))
                {
                    Split_Para.Add(Split[1].ToUpper());
                }
                // }
            }
            Split_Para.Sort();

            datagrid = new DataGridView[Data_Interface.Clotho_List[0].Max.Length];

            for (int s = 0; s < 1; s++)
            {

                string title = "Set";
                TabPage myTabPage = new TabPage(title);
                tabControl2.TabPages.Add(myTabPage);

                datagrid[s] = new DataGridView();

                datagrid[s].AllowUserToAddRows = false;
                datagrid[s].AllowUserToDeleteRows = false;
                datagrid[s].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                datagrid[s].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                datagrid[s].Location = new System.Drawing.Point(10, 10);
                datagrid[s].Name = "dataGridView1";
                datagrid[s].RowHeadersVisible = false;
                datagrid[s].RowTemplate.Height = 37;
                datagrid[s].Size = new System.Drawing.Size(1004, 1776);
                datagrid[s].TabIndex = 1;
                datagrid[s].Dock = System.Windows.Forms.DockStyle.Fill;
                datagrid[s].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
                datagrid[s].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
                datagrid[s].EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
                //   datagrid[s].KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
                datagrid[s].EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
                datagrid[s].CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
                // datagrid[s].Font = new Font("Tahoma", 6, FontStyle.Regular);
                //   datagrid[s].ReadOnly = true;
                datagrid[s].DoubleBuffered2(true);


                datagrid[s].ColumnCount = Bin_Length;

                datagrid[s].Columns[0].Name = "Parameter";
                datagrid[s].Columns[0].Frozen = true;

                int b = 0;
                for (b = 0; b < Bin_Length - 1; b++)
                {
                    datagrid[s].Columns[b + 1].Name = "Bin" + (b + 2);
                }


                datagrid[s].Columns[0].Width = 100;

                for (b = 0; b < Bin_Length - 1; b++)
                {
                    datagrid[s].Columns[b + 1].Width = 50;
                }


                for (i = 0; i < Split_Para.Count; i++)
                {
                    Valuse = new object[Bin_Length];

                    Valuse[0] = Split_Para[i];

                    for (b = 0; b < Bin_Length - 1; b++)
                    {
                        Valuse[b + 1] = null;
                    }

                    datagrid[s].Rows.Add(Valuse);

                }

                i = 0;

                foreach (var item in Split_Para)
                {
                    DataGridViewComboBoxCell comboBoxColumn;

                    for (b = 0; b < Bin_Length - 1; b++)
                    {
                        comboBoxColumn = new DataGridViewComboBoxCell();
                        comboBoxColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;

                        for (int k = 0; k < Bin_Length - 1; k++)
                        {
                            comboBoxColumn.Items.Add((k + 1).ToString());
                        }
                        comboBoxColumn.Items.Add("9999");
                        datagrid[s][b + 1, i] = comboBoxColumn;

                        if (b < Bin_Length - 3)
                        {
                            datagrid[s][b + 1, i].Value = (b + 1).ToString();
                        }
                        else
                        {
                            datagrid[s][b + 1, i].Value = "9999";
                        }

                    }

                    i++;
                }


                datagrid[s].Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;

                tabControl2.TabPages[s].Controls.Add(datagrid[s]);
            }


            if (Dic_For_Bin.Count == 0)
            {
                List<int> List_For_Bin = new List<int>();
                string Para = "";

                for (int datagrid_Row_Count = 0; datagrid_Row_Count < datagrid[0].RowCount; datagrid_Row_Count++)
                {
                    List_For_Bin = new List<int>();
                    for (int datagrid_Column_Count = 0; datagrid_Column_Count < datagrid[0].ColumnCount; datagrid_Column_Count++)
                    {

                        if (datagrid_Column_Count == 0)
                        {
                            Para = datagrid[0][datagrid_Column_Count, datagrid_Row_Count].Value.ToString();
                        }
                        else
                        {

                            string data = datagrid[0][datagrid_Column_Count, datagrid_Row_Count].Value.ToString();
                            List_For_Bin.Add(Convert.ToInt16(data));
                        }


                    }
                    Dic_For_Bin.Add(Para, List_For_Bin);
                }

            }
            #endregion


            advanced = new Zuby.ADGV.AdvancedDataGridView[Data_Interface.Clotho_List[0].Max.Length];
            _dataTable = new DataTable[Data_Interface.Clotho_List[0].Max.Length];
            bindingSource = new BindingSource[Data_Interface.Clotho_List[0].Max.Length];

            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                _dataTable[s] = new DataTable();
                bindingSource[s] = new BindingSource();

                string title = "Bin" + (s + 1);
                TabPage myTabPage = new TabPage(title);
                tabControl1.TabPages.Add(myTabPage);

                System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
                advanced[s] = new Zuby.ADGV.AdvancedDataGridView();
                advanced[s].AllowUserToAddRows = false;
                advanced[s].AllowUserToDeleteRows = false;
                advanced[s].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advanced[s].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                advanced[s].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                advanced[s].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                advanced[s].FilterAndSortEnabled = true;
                advanced[s].Location = new System.Drawing.Point(10, 10);
                advanced[s].Name = "grid" + s;
                advanced[s].RowHeadersVisible = false;
                advanced[s].RowTemplate.Height = 40;
                advanced[s].Size = new System.Drawing.Size(2000, 1695);
                advanced[s].TabIndex = 1;
                advanced[s].VirtualMode = false;
                advanced[s].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advanced[s].Dock = System.Windows.Forms.DockStyle.Fill;
                advanced[s].SortStringChanged += new System.EventHandler(this.advancedDataGridView1_SortStringChanged);
                advanced[s].FilterStringChanged += new System.EventHandler(this.advancedDataGridView1_FilterStringChanged);
                advanced[s].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellValueChanged);
                advanced[s].KeyDown += new System.Windows.Forms.KeyEventHandler(this.advancedDataGridView1_KeyDown);
                advanced[s].CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellDoubleClick);
                advanced[s].CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellEnter);
                advanced[s].CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.advancedDataGridView1_CellMouseUp);

                bindingSource[s].DataSource = _dataTable[s];
                advanced[s].DataSource = bindingSource[s];
                advanced[s].DoubleBuffered2(true);

                DataColumn[] dtkey = new DataColumn[1];
                _dataTable[s].Clear();
                _dataTable[s].Columns.Add("No", typeof(int));
                dtkey[0] = _dataTable[s].Columns["No"];
                _dataTable[s].PrimaryKey = dtkey;

                _dataTable[s].Columns.Add("Parameter");

                _dataTable[s].Columns.Add("Min_Selector", typeof(string));
                _dataTable[s].Columns.Add("Max_Selector", typeof(string));

                _dataTable[s].Columns.Add("Min", typeof(double));
                _dataTable[s].Columns.Add("Max", typeof(double));

                _dataTable[s].Columns.Add("S_Min", typeof(double));
                _dataTable[s].Columns.Add("S_Max", typeof(double));

                _dataTable[s].Columns.Add("D_Min", typeof(double));
                _dataTable[s].Columns.Add("D_Median", typeof(double));
                _dataTable[s].Columns.Add("D_Max", typeof(double));

                _dataTable[s].Columns.Add("CPK", typeof(double));
                _dataTable[s].Columns.Add("Std", typeof(double));
                _dataTable[s].Columns.Add("%", typeof(double));
                _dataTable[s].Columns.Add("Fail", typeof(int));

                //   _dataTable[s].Columns.Add("L_IQR", typeof(double));
                //   _dataTable[s].Columns.Add("H_IQR", typeof(double));
                _dataTable[s].Columns.Add("Outlier", typeof(double));




                bindingSource[s].DataMember = _dataTable[s].TableName;

                _dataTable[s].BeginLoadData();


                for (i = 0; i < Data_Interface.New_Header.Length - 1; i++)
                {
                    Valuse = new object[16];

                    Valuse[0] = i;
                    Valuse[1] = Data_Interface.Reference_Header[i + 1];

                    string[] ParanameSplit = Data_Interface.Reference_Header[i + 1].Split('_');

                    if (ParanameSplit[ParanameSplit.Length - 1].Contains('-'))
                    {
                        Valuse[2] = "CUSTOMER";
                        Valuse[3] = "CUSTOMER";
                        Valuse[4] = Data_Interface.Clotho_List[i + 1].Min[0];
                        Valuse[5] = Data_Interface.Clotho_List[i + 1].Max[0];
                        Valuse[6] = Data_Interface.Clotho_List[i + 1].Min[0];
                        Valuse[7] = Data_Interface.Clotho_List[i + 1].Max[0];


                    }
                    else
                    {
                        Valuse[2] = "CPK";
                        Valuse[3] = "CPK";
                        Valuse[4] = Data_Interface.Clotho_List[i + 1].Min[0];
                        Valuse[5] = Data_Interface.Clotho_List[i + 1].Max[0];
                        Valuse[6] = Data_Interface.Clotho_List[i + 1].Min[0];
                        Valuse[7] = Data_Interface.Clotho_List[i + 1].Max[0];

                    }




                    Valuse[8] = Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Min_Data[s];
                    Valuse[9] = Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Median_Data[s];
                    Valuse[10] = Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Max_Data[s];

                    double L_CPK = (Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Avg[s] - Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Min_Data[s]) / (3 * Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Std[s]);
                    double H_CPK = (Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Max_Data[s] - Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Avg[s]) / (3 * Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Std[s]);

                    double Data = 0f;
                    if (L_CPK > H_CPK) Data = H_CPK;
                    else Data = L_CPK;

                    Valuse[11] = 0;
                    Valuse[12] = Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Std[s];
                    Valuse[13] = 0;
                    Valuse[14] = 0;
                    //   Valuse[15] = 1.5;
                    //   Valuse[16] = 1.5;
                    Valuse[15] = 0;


                    _dataTable[s].Rows.Add(Valuse);
                }

                _dataTable[s].EndLoadData();
                bindingSource[s].DataSource = _dataTable[s];

                tabControl1.TabPages[s].Controls.Add(advanced[s]);

            }

            #region

            datagrid2 = new DataGridView[Data_Interface.Clotho_List[0].Max.Length];

            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                string title = "Bin" + (s + 1);
                TabPage myTabPage = new TabPage(title);
                tabControl3.TabPages.Add(myTabPage);

                datagrid2[s] = new DataGridView();

                datagrid2[s].AllowUserToAddRows = false;
                datagrid2[s].AllowUserToDeleteRows = false;
                datagrid2[s].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                datagrid2[s].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                datagrid2[s].Location = new System.Drawing.Point(10, 10);
                datagrid2[s].Name = "dataGridView1";
                datagrid2[s].RowHeadersVisible = false;
                datagrid2[s].RowTemplate.Height = 37;
                datagrid2[s].Size = new System.Drawing.Size(1004, 1776);
                datagrid2[s].TabIndex = 1;
                datagrid2[s].Dock = System.Windows.Forms.DockStyle.Fill;
                datagrid2[s].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
                datagrid2[s].KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
                datagrid2[s].EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
                //  datagrid2[s].EditMode = DataGridViewEditMode.EditOnEnter;
                datagrid2[s].DoubleBuffered2(true);
                datagrid2[s].RowHeadersVisible = false;
                datagrid2[s].ColumnCount = 2;
                datagrid2[s].BackgroundColor = Color.White;
                datagrid2[s].ReadOnly = false;


                for (i = 0; i < 7; i++)
                {
                    Valuse = new object[2];

                    switch (i)
                    {
                        case 0:
                            Valuse[0] = "Total Sample";
                            Valuse[1] = this.Total;
                            break;
                        case 1:
                            Valuse[0] = "Analysis Sample";
                            Valuse[1] = this.Any_Total;
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
                            Valuse[1] = this.Hidden_Total;
                            break;
                        case 6:
                            Valuse[0] = "Outlier";
                            Valuse[1] = 0;
                            break;
                    }



                    datagrid2[s].Rows.Add(Valuse);

                }


                datagrid2[s].Columns[0].Width = 100;
                datagrid2[s].Columns[1].Width = 40;


                tabControl3.TabPages[s].Controls.Add(datagrid2[s]);


            }

            #endregion

            if (A_ForCtrlz_List_count == 0)
            {
                ForCtrlz_Min = new List<double[]>();
                ForCtrlz_Max = new List<double[]>();
                ForCtrlz_List = new List<forctrlz>[Data_Interface.Clotho_List[0].Max.Length];

                for (int a = 0; a < Data_Interface.Clotho_List[0].Max.Length; a++)
                {
                    ForCtrlz_List[a] = new List<forctrlz>();
                    double[] Min = new double[Data_Interface.Clotho_List.Count];
                    double[] Max = new double[Data_Interface.Clotho_List.Count];

                    for (int o = 0; o < Data_Interface.Clotho_List.Count; o++)
                    {
                        Min[o] = Data_Interface.Clotho_List[o].Min[a];
                        Max[o] = Data_Interface.Clotho_List[o].Max[a];
                    }
                    ForCtrlz_Min.Add(Min);
                    ForCtrlz_Max.Add(Max);

                }
                ForCtrlz_List_count++;
            }


            Calculate_thread_Strat = new long[Data_Interface.DB_Count];
            Calculate_thread_End = new long[Data_Interface.DB_Count];
            Sample_Verify = new int[Bin_Length];

            List_Sample_Verify = new List<int[]>[Data_Interface.DB_Count];

            double ThreadCount = Convert.ToDouble(this.Total) / Convert.ToDouble(Data_Interface.DB_Count);
            double Temp = Math.Truncate(ThreadCount);

            if (ThreadCount > Temp) ThreadCount = Convert.ToInt16(Temp) + 1;
            else ThreadCount = Convert.ToInt16(Temp);

            int dummy = Convert.ToInt16(ThreadCount);
            int dummy2 = Convert.ToInt16(ThreadCount);
            for (int u = 0; u < Data_Interface.DB_Count; u++)
            {
                if (u == 0)
                {
                    Calculate_thread_Strat[u] = 0;
                    Calculate_thread_End[u] = dummy;
                }
                else if (u == Data_Interface.DB_Count - 1)
                {
                    Calculate_thread_Strat[u] = dummy;
                    Calculate_thread_End[u] = Total;

                }
                else
                {
                    Calculate_thread_Strat[u] = dummy;
                    Calculate_thread_End[u] = dummy + Convert.ToInt16(ThreadCount);
                    dummy += dummy2;

                }
            }

            for (int j = 0; j < Data_Interface.Clotho_List[0].Max.Length; j++)
            {

                advanced[j].Update();
            }
        }

        public void LoadSpec()
        {

            #region

            Split_Para = new List<string>();

            for (i = 1; i < Data_Interface.New_Header.Length; i++)
            {
                string[] Split = Data_Interface.Reference_Header[i].Split('_');

                // if (Split[0].ToUpper().ToString() != "M")
                // {
                if (!Split_Para.Contains(Split[1].ToUpper()))
                {
                    Split_Para.Add(Split[1].ToUpper());
                }
                // }
            }
            Split_Para.Sort();

            datagrid = new DataGridView[Data_Interface.Clotho_List[0].Max.Length];

            for (int s = 0; s < 1; s++)
            {

                string title = "Set";
                TabPage myTabPage = new TabPage(title);
                tabControl2.TabPages.Add(myTabPage);

                datagrid[s] = new DataGridView();

                datagrid[s].AllowUserToAddRows = false;
                datagrid[s].AllowUserToDeleteRows = false;
                datagrid[s].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                datagrid[s].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                datagrid[s].Location = new System.Drawing.Point(10, 10);
                datagrid[s].Name = "dataGridView1";
                datagrid[s].RowHeadersVisible = false;
                datagrid[s].RowTemplate.Height = 37;
                datagrid[s].Size = new System.Drawing.Size(1004, 1776);
                datagrid[s].TabIndex = 1;
                datagrid[s].Dock = System.Windows.Forms.DockStyle.Fill;
                datagrid[s].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
                datagrid[s].CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
                datagrid[s].EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
                //   datagrid[s].KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
                datagrid[s].EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
                datagrid[s].CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
                // datagrid[s].Font = new Font("Tahoma", 6, FontStyle.Regular);
                //   datagrid[s].ReadOnly = true;
                datagrid[s].DoubleBuffered2(true);


                datagrid[s].ColumnCount = Bin_Length;

                datagrid[s].Columns[0].Name = "Parameter";
                datagrid[s].Columns[0].Frozen = true;

                int b = 0;
                for (b = 0; b < Bin_Length - 1; b++)
                {
                    datagrid[s].Columns[b + 1].Name = "Bin" + (b + 2);
                }


                datagrid[s].Columns[0].Width = 100;

                for (b = 0; b < Bin_Length - 1; b++)
                {
                    datagrid[s].Columns[b + 1].Width = 50;
                }


                for (i = 0; i < Split_Para.Count; i++)
                {
                    Valuse = new object[Bin_Length];

                    Valuse[0] = Split_Para[i];

                    for (b = 0; b < Bin_Length - 1; b++)
                    {
                        Valuse[b + 1] = null;
                    }

                    datagrid[s].Rows.Add(Valuse);

                }

                i = 0;

                foreach (var item in Split_Para)
                {
                    DataGridViewComboBoxCell comboBoxColumn;

                    for (b = 0; b < Bin_Length - 1; b++)
                    {
                        comboBoxColumn = new DataGridViewComboBoxCell();
                        comboBoxColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;

                        for (int k = 0; k < Bin_Length - 1; k++)
                        {
                            comboBoxColumn.Items.Add((k + 1).ToString());
                        }
                        comboBoxColumn.Items.Add("9999");
                        datagrid[s][b + 1, i] = comboBoxColumn;

                        if (b < Bin_Length - 3)
                        {
                            datagrid[s][b + 1, i].Value = (b + 1).ToString();
                        }
                        else
                        {
                            datagrid[s][b + 1, i].Value = "9999";
                        }

                    }

                    i++;
                }


                datagrid[s].Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                datagrid[s].Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;

                tabControl2.TabPages[s].Controls.Add(datagrid[s]);
            }


            if (Dic_For_Bin.Count == 0)
            {
                List<int> List_For_Bin = new List<int>();
                string Para = "";

                for (int datagrid_Row_Count = 0; datagrid_Row_Count < datagrid[0].RowCount; datagrid_Row_Count++)
                {
                    List_For_Bin = new List<int>();
                    for (int datagrid_Column_Count = 0; datagrid_Column_Count < datagrid[0].ColumnCount; datagrid_Column_Count++)
                    {

                        if (datagrid_Column_Count == 0)
                        {
                            Para = datagrid[0][datagrid_Column_Count, datagrid_Row_Count].Value.ToString();
                        }
                        else
                        {

                            string data = datagrid[0][datagrid_Column_Count, datagrid_Row_Count].Value.ToString();
                            List_For_Bin.Add(Convert.ToInt16(data));
                        }


                    }
                    Dic_For_Bin.Add(Para, List_For_Bin);
                }

            }
            #endregion


            advanced = new Zuby.ADGV.AdvancedDataGridView[Data_Interface.Clotho_List[0].Max.Length];
            _dataTable = new DataTable[Data_Interface.Clotho_List[0].Max.Length];
            bindingSource = new BindingSource[Data_Interface.Clotho_List[0].Max.Length];

            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                _dataTable[s] = new DataTable();
                bindingSource[s] = new BindingSource();

                string title = "Bin" + (s + 1);
                TabPage myTabPage = new TabPage(title);
                tabControl1.TabPages.Add(myTabPage);

                System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();

                dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
                advanced[s] = new Zuby.ADGV.AdvancedDataGridView();
                advanced[s].AllowUserToAddRows = false;
                advanced[s].AllowUserToDeleteRows = false;
                advanced[s].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advanced[s].AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
                advanced[s].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                advanced[s].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                advanced[s].FilterAndSortEnabled = true;
                advanced[s].Location = new System.Drawing.Point(10, 10);
                advanced[s].Name = "grid" + s;
                advanced[s].RowHeadersVisible = false;
                advanced[s].RowTemplate.Height = 40;
                advanced[s].Size = new System.Drawing.Size(2000, 1695);
                advanced[s].TabIndex = 1;
                advanced[s].VirtualMode = false;
                advanced[s].AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advanced[s].Dock = System.Windows.Forms.DockStyle.Fill;
                advanced[s].SortStringChanged += new System.EventHandler(this.advancedDataGridView1_SortStringChanged);
                advanced[s].FilterStringChanged += new System.EventHandler(this.advancedDataGridView1_FilterStringChanged);
                advanced[s].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellValueChanged);
                advanced[s].KeyDown += new System.Windows.Forms.KeyEventHandler(this.advancedDataGridView1_KeyDown);
                advanced[s].CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellDoubleClick);
                advanced[s].CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.advancedDataGridView1_CellEnter);
                advanced[s].CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.advancedDataGridView1_CellMouseUp);

                bindingSource[s].DataSource = _dataTable[s];
                advanced[s].DataSource = bindingSource[s];
                advanced[s].DoubleBuffered2(true);

                DataColumn[] dtkey = new DataColumn[1];
                _dataTable[s].Clear();
                _dataTable[s].Columns.Add("No", typeof(int));
                dtkey[0] = _dataTable[s].Columns["No"];
                _dataTable[s].PrimaryKey = dtkey;

                _dataTable[s].Columns.Add("Parameter");

                _dataTable[s].Columns.Add("Min_Selector", typeof(string));
                _dataTable[s].Columns.Add("Max_Selector", typeof(string));

                _dataTable[s].Columns.Add("Min", typeof(double));
                _dataTable[s].Columns.Add("Max", typeof(double));

                _dataTable[s].Columns.Add("S_Min", typeof(double));
                _dataTable[s].Columns.Add("S_Max", typeof(double));

                _dataTable[s].Columns.Add("D_Min", typeof(double));
                _dataTable[s].Columns.Add("D_Median", typeof(double));
                _dataTable[s].Columns.Add("D_Max", typeof(double));

                _dataTable[s].Columns.Add("CPK", typeof(double));
                _dataTable[s].Columns.Add("Std", typeof(double));
                _dataTable[s].Columns.Add("%", typeof(double));
                _dataTable[s].Columns.Add("Fail", typeof(int));

                //   _dataTable[s].Columns.Add("L_IQR", typeof(double));
                //   _dataTable[s].Columns.Add("H_IQR", typeof(double));
                _dataTable[s].Columns.Add("Outlier", typeof(double));




                bindingSource[s].DataMember = _dataTable[s].TableName;

                _dataTable[s].BeginLoadData();

                for (i = 0; i < Data_Interface.New_Header.Length - 1; i++)
                {
                    Valuse = new object[16];

                    Valuse[0] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].No[s];
                    Valuse[1] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Parameter[s];

                    string[] ParanameSplit = Data_Interface.Reference_Header[i + 1].Split('_');


                    Valuse[2] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Min_Selector[s];
                    Valuse[3] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Max_Selector[s];
                    Valuse[4] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Min_Spec_Control[s];
                    Valuse[5] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Max_Spec_Control[s];
                    Valuse[6] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Min_Spec[s];
                    Valuse[7] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Max_Spec[s];

                    Data_Interface.Customor_Clotho_List[i + 1].Min[s] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Min_Spec[s];
                    Data_Interface.Customor_Clotho_List[i + 1].Max[s] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Max_Spec[s];


                    Valuse[8] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Min_Data[s];
                    Valuse[9] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Median_Data[s];
                    Valuse[10] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Max_Data[s];


                    Valuse[11] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].CPK[s];
                    Valuse[12] = Db_Interface.Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Std[s];
                    Valuse[13] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Persent[s];
                    Valuse[14] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Fail_Count[s];
                    //   Valuse[15] = 1.5;
                    //   Valuse[16] = 1.5;
                    Valuse[15] = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[i + 1]].Outlier[s];


                    _dataTable[s].Rows.Add(Valuse);
                }
                _dataTable[s].EndLoadData();
                bindingSource[s].DataSource = _dataTable[s];

                tabControl1.TabPages[s].Controls.Add(advanced[s]);

            }

            #region

            datagrid2 = new DataGridView[Data_Interface.Clotho_List[0].Max.Length];

            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                string title = "Bin" + (s + 1);
                TabPage myTabPage = new TabPage(title);
                tabControl3.TabPages.Add(myTabPage);

                datagrid2[s] = new DataGridView();

                datagrid2[s].AllowUserToAddRows = false;
                datagrid2[s].AllowUserToDeleteRows = false;
                datagrid2[s].AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
                datagrid2[s].ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                datagrid2[s].Location = new System.Drawing.Point(10, 10);
                datagrid2[s].Name = "dataGridView1";
                datagrid2[s].RowHeadersVisible = false;
                datagrid2[s].RowTemplate.Height = 37;
                datagrid2[s].Size = new System.Drawing.Size(1004, 1776);
                datagrid2[s].TabIndex = 1;
                datagrid2[s].Dock = System.Windows.Forms.DockStyle.Fill;
                datagrid2[s].CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
                datagrid2[s].KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyDown);
                datagrid2[s].EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
                //  datagrid2[s].EditMode = DataGridViewEditMode.EditOnEnter;
                datagrid2[s].DoubleBuffered2(true);
                datagrid2[s].RowHeadersVisible = false;
                datagrid2[s].ColumnCount = 2;
                datagrid2[s].BackgroundColor = Color.White;
                datagrid2[s].ReadOnly = false;


                for (i = 0; i < 7; i++)
                {
                    Valuse = new object[2];

                    switch (i)
                    {
                        case 0:
                            Valuse[0] = "Total Sample";
                            Valuse[1] = this.Total;
                            break;
                        case 1:
                            Valuse[0] = "Analysis Sample";
                            Valuse[1] = this.Any_Total;
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
                            Valuse[1] = this.Hidden_Total;
                            break;
                        case 6:
                            Valuse[0] = "Outlier";
                            Valuse[1] = 0;
                            break;
                    }



                    datagrid2[s].Rows.Add(Valuse);

                }


                datagrid2[s].Columns[0].Width = 100;
                datagrid2[s].Columns[1].Width = 40;


                tabControl3.TabPages[s].Controls.Add(datagrid2[s]);


            }

            #endregion

            if (A_ForCtrlz_List_count == 0)
            {
                ForCtrlz_Min = new List<double[]>();
                ForCtrlz_Max = new List<double[]>();
                ForCtrlz_List = new List<forctrlz>[Data_Interface.Clotho_List[0].Max.Length];

                for (int a = 0; a < Data_Interface.Clotho_List[0].Max.Length; a++)
                {
                    ForCtrlz_List[a] = new List<forctrlz>();
                    double[] Min = new double[Data_Interface.Clotho_List.Count];
                    double[] Max = new double[Data_Interface.Clotho_List.Count];

                    for (int o = 0; o < Data_Interface.Clotho_List.Count; o++)
                    {
                        Min[o] = Data_Interface.Clotho_List[o].Min[a];
                        Max[o] = Data_Interface.Clotho_List[o].Max[a];
                    }
                    ForCtrlz_Min.Add(Min);
                    ForCtrlz_Max.Add(Max);

                }
                ForCtrlz_List_count++;
            }


            Calculate_thread_Strat = new long[Data_Interface.DB_Count];
            Calculate_thread_End = new long[Data_Interface.DB_Count];
            Sample_Verify = new int[Bin_Length];

            List_Sample_Verify = new List<int[]>[Data_Interface.DB_Count];

            double ThreadCount = Convert.ToDouble(this.Total) / Convert.ToDouble(Data_Interface.DB_Count);
            double Temp = Math.Truncate(ThreadCount);

            if (ThreadCount > Temp) ThreadCount = Convert.ToInt16(Temp) + 1;
            else ThreadCount = Convert.ToInt16(Temp);

            int dummy = Convert.ToInt16(ThreadCount);
            int dummy2 = Convert.ToInt16(ThreadCount);
            for (int u = 0; u < Data_Interface.DB_Count; u++)
            {
                if (u == 0)
                {
                    Calculate_thread_Strat[u] = 0;
                    Calculate_thread_End[u] = dummy;
                }
                else if (u == Data_Interface.DB_Count - 1)
                {
                    Calculate_thread_Strat[u] = dummy;
                    Calculate_thread_End[u] = Total;

                }
                else
                {
                    Calculate_thread_Strat[u] = dummy;
                    Calculate_thread_End[u] = dummy + Convert.ToInt16(ThreadCount);
                    dummy += dummy2;

                }
            }

            for (int j = 0; j < Data_Interface.Clotho_List[0].Max.Length; j++)
            {

                advanced[j].Update();
            }
        }
        private void advancedDataGridView1_SortStringChanged(object sender, EventArgs e)
        {
            int index = tabControl1.SelectedIndex;
            string SortString = advanced[index].SortString;

            Stopwatch TestTime2 = new Stopwatch();
            TestTime2.Restart();
            TestTime2.Start();

            for (int length = 0; length < 1; length++)
            {
                string Value = datagrid[0][Col + length, Row].Value.ToString();
                string Test = datagrid[0][TabIndex + 1, Row].Value.ToString();

                DataView View = new DataView(_dataTable[index]);
                View.Sort = advanced[index].SortString;

                DataColumn[] dtkey = new DataColumn[1];

                dtkey[0] = _dataTable[index].Columns["No"];
                _dataTable[index].PrimaryKey = dtkey;

                _dataTable[index].PrimaryKey = new DataColumn[] { _dataTable[index].Columns["No"] };
                _dataTable[index] = View.ToTable();
                bindingSource[index].DataSource = _dataTable[index];

                advanced[index].FirstDisplayedScrollingRowIndex = 0;

                double Testtime5 = TestTime2.Elapsed.TotalMilliseconds;


                Flag = true;
                for (int w = tabControl1.SelectedIndex; w < Bin_Length - 1; w++)
                {
                    Test = datagrid[0][w + 1, Row].Value.ToString();
                    if (Test != "9999")
                    {
                        View = new DataView(_dataTable[Convert.ToInt16(Test)]);
                        View.Sort = SortString;

                        dtkey = new DataColumn[1];

                        dtkey[0] = _dataTable[Convert.ToInt16(Test)].Columns["No"];
                        _dataTable[Convert.ToInt16(Test)].PrimaryKey = dtkey;

                        _dataTable[Convert.ToInt16(Test)].PrimaryKey = new DataColumn[] { _dataTable[Convert.ToInt16(Test)].Columns["No"] };
                        _dataTable[Convert.ToInt16(Test)] = View.ToTable();
                        bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];

                        advanced[Convert.ToInt16(Test)].FirstDisplayedScrollingRowIndex = 0;

                        double Testtime3 = TestTime2.Elapsed.TotalMilliseconds;
                    }
                }
                Flag = false;
            }
            ForeColor();
        }
        private void advancedDataGridView1_FilterStringChanged(object sender, EventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            bindingSource[index].Filter = advanced[index].FilterString;
        }
        private void advancedDataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int tabControl1_index = tabControl1.SelectedIndex;

            if (advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {

                if (!Flag)
                {
                    if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
                    {
                        Stopwatch TestTime1 = new Stopwatch();
                        TestTime1.Restart();
                        TestTime1.Start();

                        EditColumnIndex2and3(tabControl1_index, e.ColumnIndex, e.RowIndex, advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());

                        double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;


                    }
                    else if (e.ColumnIndex == 4 || e.ColumnIndex == 5)
                    {
                        string result = advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                        double numChk = 0;
                        bool isNum = double.TryParse(result, out numChk);

                        if (!isNum)
                        {
                            MessageBox.Show("Please Check Input String");
                        }
                        else
                        {

                            Stopwatch TestTime1 = new Stopwatch();
                            TestTime1.Restart();
                            TestTime1.Start();

                            EditColumnIndex4and5(tabControl1_index, e.ColumnIndex, e.RowIndex, advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());

                            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
                        }
                    }
                    else if (e.ColumnIndex == 15 || e.ColumnIndex == 16)
                    {
                        EditColumnIndex15and16(tabControl1_index, e.ColumnIndex, e.RowIndex, advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());


                    }
                }
            }

        }
        private void advancedDataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            Stopwatch TestTime2 = new Stopwatch();
            TestTime2.Restart();
            TestTime2.Start();

            if (e.Control && e.KeyCode == Keys.C)
            {

                DataObject Do = advanced[index].GetClipboardContent();
                Clipboard.SetDataObject(Do);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');

                for (int k = 0; k < advanced[index].SelectedCells.Count; k++)
                {
                    // advanced[index].CellBeginEdit(true);

                    int row = advanced[index].SelectedCells[k].RowIndex;
                    int col = advanced[index].SelectedCells[k].ColumnIndex;


                    foreach (string line in lines)
                    {
                        if (row < advanced[index].RowCount && line.Length > 0)
                        {

                            Stopwatch TestTime1 = new Stopwatch();
                            TestTime1.Restart();
                            TestTime1.Start();

                            string[] cells = line.Split('\t');
                            int Cells_Length = cells.GetLength(0);
                            for (int i = 0; i < Cells_Length; ++i)
                            {
                                if (col + i < advanced[index].ColumnCount)
                                {
                                    double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
                                    if (advanced[index][col + i, row].Value == "CPK" || advanced[index][col + i, row].Value == "MANUAL" || advanced[index][col + i, row].Value == "RANGE" || advanced[index][col + i, row].Value == "9999" || advanced[index][col + i, row].Value == "-9999")
                                    {
                                        string result = cells[i].ToString();
                                        double numChk = 0;
                                        bool isNum = double.TryParse(result, out numChk);

                                        if (isNum)
                                        {
                                            MessageBox.Show("Wrong Copy and Paste");
                                        }
                                        else
                                        {
                                            advanced[index][col + i, row].Value = Convert.ChangeType(cells[i], advanced[index][col + i, row].ValueType);
                                            double Testtime52 = TestTime1.Elapsed.TotalMilliseconds;
                                        }

                                    }
                                    else if (advanced[index][col + i, row].Value != "CPK" && advanced[index][col + i, row].Value != "MANUAL" && advanced[index][col + i, row].Value != "RANGE" && advanced[index][col + i, row].Value != "9999" && advanced[index][col + i, row].Value == "-9999")
                                    {

                                        advanced[index][col + i, row].Value = Convert.ChangeType(cells[i], advanced[index][col + i, row].ValueType);
                                        double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                                    }
                                    else
                                    {
                                        if (Cells_Length == 2)
                                        {
                                            if (i == 0) advanced[index][2, row].Value = cells[i];
                                            else advanced[index][3, row].Value = cells[i];

                                        }
                                        else advanced[index][col + i, row].Value = cells[i];



                                        double Testtime6 = TestTime1.Elapsed.TotalMilliseconds;


                                    }

                                }
                                else
                                {
                                    break;
                                }
                                double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;
                            }
                            row++;
                        }
                        else
                        {
                            break;
                        }



                    }

                }


                for (int j = 0; j < Data_Interface.Clotho_List[0].Max.Length; j++)
                {
                    // _dataTable[index].AcceptChanges();
                    bindingSource[index].DataSource = _dataTable[index];
                    advanced[index].DataSource = bindingSource[index];
                    advanced[index].Update();
                    // advanced[index].FirstDisplayedScrollingRowIndex = 0;

                }


                double Testtime77 = TestTime2.Elapsed.TotalMilliseconds;
            }
            else if (e.Control && e.KeyCode == Keys.Z)
            {
                //if (A_ForCtrlz_List[index].Count != 0)
                //{

                //    A_Cz = A_ForCtrlz_List[index][A_ForCtrlz_List[index].Count - 1];

                //    DataRow dr = _dataTable[index].Rows.Find(A_Cz.No);
                //    int SelRow = _dataTable[index].Rows.IndexOf(dr);

                //    advanced[index][A_Cz.Col, A_Cz.Row].Value = A_Cz.Ref_Value;

                //    _dataTable[index].Rows[SelRow][A_Cz.Col] = advanced[index][A_Cz.Col, A_Cz.Row].Value;
                //    bindingSource[index].DataSource = _dataTable[index];
                //    advanced[index].Update();
                //    A_ForCtrlz_List[index].RemoveAt(A_ForCtrlz_List[index].Count - 1);

                //}

            }

            else if (e.KeyCode == Keys.F1)
            {
                //  advanced[tabControl1.SelectedIndex].Visible = false;
                int Tab = tabControl1.SelectedIndex;
                int Tab1 = tabControl1.SelectedIndex;

                currentCell = advanced[Tab].CurrentCell;

                int Row = datagrid[0].CurrentCell.RowIndex;
                int Bin_Defi = Bin_Length - 1;


                string dummy = datagrid[0].Rows[Row].Cells[Tab + 1].Value.ToString();

                if (dummy != "9999")
                {
                    if (Tab == _dataTable.Length - 1)
                    {
                        Tab = -1;
                    }

                    tabControl1.SelectedIndex = Tab + 1;
                    advanced[Tab1].CurrentCell = currentCell;
                    advanced[tabControl1.SelectedIndex].Visible = true;
                    advanced[tabControl1.SelectedIndex].Focus();
                }
                else
                {

                    tabControl1.SelectedIndex = 0;
                    advanced[Tab1].CurrentCell = currentCell;
                    advanced[tabControl1.SelectedIndex].Visible = true;
                    advanced[tabControl1.SelectedIndex].Focus();
                }

            }
            else if (e.KeyCode == Keys.F2)
            {
                //DataGridViewCell Cell = datagrid[0].CurrentCell;
                //string Name = datagrid[0].Rows[Cell.RowIndex].Cells[Cell.ColumnIndex].Value.ToString();
                //datagrid[0].Rows[Cell.RowIndex].Cells[Cell.ColumnIndex].

            }
        }
        private void advancedDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            int index = tabControl1.SelectedIndex;

            if (e.ColumnIndex == 1)
            {
                string CellValue = advanced[index].Rows[e.RowIndex].Cells[1].Value.ToString();

                for (int i = 0; i < Data_Interface.Reference_Header.Length; i++)
                {
                    if (CellValue == Data_Interface.Reference_Header[i])
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
                    //    Db_Interface.Get_Selected_Para(Find_DB, Data_Interface.New_Header[i]);

                        CSV_Class.CSV CSV = new CSV_Class.CSV();
                        CSV_Class.CSV.INT CSV_Interface = CSV.Open(Key);

                        CSV_Interface.Write_Open("C:\\temp\\dummy\\" + Data_Interface.Reference_Header[i] + ".csv");
                        CSV_Interface.Write(Data_Interface.Reference_Header[i], Db_Interface.ID, Db_Interface.Value);
                        CSV_Interface.Write_Close();


                      //  JMP_Draw("C:\\temp\\dummy\\" + Data_Interface.Reference_Header[i] + ".csv", "", Data_Interface.Reference_Header[i], advanced[index].Rows[e.RowIndex].Cells[e.ColumnIndex + 5].Value.ToString(), advanced[index].Rows[e.RowIndex].Cells[e.ColumnIndex + 6].Value.ToString(),"","", null, "" , false);



                    }
                }
            }
        }
        private void advancedDataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            if (e.ColumnIndex == advanced[index].Columns["Min_Selector"].Index)
            {
                SetComboBoxCellType objChangeCellType = new SetComboBoxCellType(ChangeCellToComboBox_Min);
                advanced[index].BeginInvoke(objChangeCellType, e.RowIndex);

                bIsComboBox = false;

            }
            else if (e.ColumnIndex == advanced[index].Columns["Max_Selector"].Index)
            {
                SetComboBoxCellType objChangeCellType = new SetComboBoxCellType(ChangeCellToComboBox_Max);
                advanced[index].BeginInvoke(objChangeCellType, e.RowIndex);

                bIsComboBox = false;

            }

        }
        private void ChangeCellToComboBox_Min(int iRowIndex)
        {
            int index = tabControl1.SelectedIndex;
            if (bIsComboBox == false)

            {

                DataGridViewComboBoxCell dgComboCell = new DataGridViewComboBoxCell();


                dgComboCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;

                DataTable dt = new DataTable();

                dt.Columns.Add("Min_Selector", typeof(string));

                DataRow dr = dt.NewRow();

                dr["Min_Selector"] = "CPK";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Min_Selector"] = "MANUAL";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Min_Selector"] = "RANGE";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Min_Selector"] = "CUSTOMER";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Min_Selector"] = "FIXEDPOUT";
                dt.Rows.Add(dr);


                dgComboCell.DataSource = dt;

                dgComboCell.ValueMember = "Min_Selector";
                dgComboCell.DisplayMember = "Min_Selector";

                advanced[index].Rows[iRowIndex].Cells[advanced[index].CurrentCell.ColumnIndex] = dgComboCell;

                bIsComboBox = true;

            }

        }
        private void ChangeCellToComboBox_Max(int iRowIndex)
        {
            int index = tabControl1.SelectedIndex;
            if (bIsComboBox == false)

            {


                DataGridViewComboBoxCell dgComboCell = new DataGridViewComboBoxCell();

                dgComboCell.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;

                DataTable dt = new DataTable();

                dt.Columns.Add("Max_Selector", typeof(string));


                DataRow dr = dt.NewRow();

                dr["Max_Selector"] = "CPK";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Max_Selector"] = "MANUAL";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Max_Selector"] = "RANGE";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Max_Selector"] = "CUSTOMER";
                dt.Rows.Add(dr);

                dr = dt.NewRow();

                dr["Max_Selector"] = "FIXEDPOUT";
                dt.Rows.Add(dr);

                dgComboCell.DataSource = dt;

                dgComboCell.ValueMember = "Max_Selector";
                dgComboCell.DisplayMember = "Max_Selector";

                advanced[index].Rows[iRowIndex].Cells[advanced[index].CurrentCell.ColumnIndex] = dgComboCell;


                // advanced[index].Update();
                bIsComboBox = true;

            }

        }


        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            string item = cb.Text;
            currentCell = datagrid[tabControl2.SelectedIndex].CurrentCell;


            datagrid[tabControl2.SelectedIndex][currentCell.ColumnIndex, currentCell.RowIndex].Value = null;

        }
        private void ctl_Enter(object sender, EventArgs e)
        {
            int tabControl1_index = tabControl1.SelectedIndex;

            (sender as ComboBox).DroppedDown = false;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
            {
                DataGridView grid = (DataGridView)sender;

                grid.BeginEdit(true);
                ((ComboBox)grid.EditingControl).DroppedDown = true;
            }
        }
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            if (e.KeyCode == Keys.F1)
            {
                int Tab = tabControl3.SelectedIndex;
                int Tab1 = tabControl3.SelectedIndex;

                currentCell = advanced[Tab].CurrentCell;

                if (Tab == _dataTable.Length - 1)
                {
                    Tab = -1;
                }
                tabControl1.SelectedIndex = Tab + 1;
                //  advanced[Tab1].CurrentCell = currentCell;
                advanced[tabControl1.SelectedIndex].Visible = true;

                // advanced[Tab].Rows[currentCell.RowIndex].Cells[currentCell.ColumnIndex].Selected = true;

            }
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            ComboBox cmbBx = e.Control as ComboBox;

            if (cmbBx != null)
            {
                cmbBx.Enter -= new EventHandler(ctl_Enter);
                cmbBx.Enter += new EventHandler(ctl_Enter);

                cmbBx.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
                cmbBx.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);

                e.CellStyle.BackColor = datagrid[tabControl2.SelectedIndex].DefaultCellStyle.BackColor;
            }
        }
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                _dataTable[s].DefaultView.Sort = "[No] ASC";
                bindingSource[s].Filter = "";

                advanced[s].CleanFilter();
                advanced[s].CleanSort();
            }

            int index = tabControl1.SelectedIndex;
            ForFilter = new StringBuilder();
            string Choose_para = datagrid[0][e.ColumnIndex, e.RowIndex].Value.ToString();
            Row = e.RowIndex;
            Col = e.ColumnIndex;
            if (e.ColumnIndex == 0)
            {

                string[] Select_Para = new string[40000];

                int Count = 0;

                for (int k = 1; k < Data_Interface.Reference_Header.Length; k++)
                {
                    string[] Split = Data_Interface.Reference_Header[k].Split('_');

                    if (Split[1].ToUpper() == Choose_para)
                    {
                        Select_Para[Count] = Data_Interface.Reference_Header[k];
                        Count++;

                    }
                }
                Array.Resize(ref Select_Para, Count);
                for (int k = 0; k < Select_Para.Length; k++)
                {
                    if (k == 0)
                    {
                        ForFilter.Append("Parameter IN (");

                        if (Select_Para.Length == 1)
                        {
                            ForFilter.Append("'" + Select_Para[k] + "')");
                        }
                        else
                        {
                            ForFilter.Append("'" + Select_Para[k] + "',");
                        }

                    }
                    else if (k == Select_Para.Length - 1)
                    {
                        ForFilter.Append("'" + Select_Para[k] + "')");
                    }
                    else
                    {
                        ForFilter.Append("'" + Select_Para[k] + "',");
                    }
                }

            }

            double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;

            ThreadFlags = new ManualResetEvent[Data_Interface.Clotho_List[0].Max.Length];
            bool[] Wait = new bool[Data_Interface.Clotho_List[0].Max.Length];

            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                //int a = s;
                // Wait[s] = false;
                //  ThreadFlags[s] = new ManualResetEvent(false);
                //  ThreadPool.QueueUserWorkItem(new WaitCallback(dd), s);
                //advanced[TabIndex].Visible = false;

                string Test = datagrid[0][s, e.RowIndex].Value.ToString();

                if (s == 0)
                {
                    bindingSource[s].Filter = ForFilter.ToString();
                }

                else if (Test != "9999")
                {
                    bindingSource[s].Filter = ForFilter.ToString();
                }

                //advanced[TabIndex].Visible = true;

            }

            double Testtime42 = TestTime1.Elapsed.TotalMilliseconds;

            ForeColor();


            for (int i = 0; i < Data_Interface.Clotho_List[0].Max.Length; i++)
            {

                //   Wait[i] = ThreadFlags[i].WaitOne();
            }

            //   advanced[0].Refresh();
            //   advanced[1].Refresh();
            //   advanced[2].Refresh();
            //    advanced[3].Refresh();
            //   advanced[4].Refresh();

            tabControl1.SelectedIndex = 0;


            double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;
            //  bindingSource[index].Filter = "Parameter LIKE '%" + Choose_para + "%'";
        }
        private void dd(Object i)
        {

            bindingSource[(int)i].Filter = ForFilter.ToString();

            //    ForeColor((int)i);
            ThreadFlags[(int)i].Set();
        }
        private void advancedDataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                clickedCell = (sender as DataGridView).Rows[e.RowIndex].Cells[e.ColumnIndex];
                this.advanced[index].CurrentCell = clickedCell;
                var relativeMousePosition = advanced[index].PointToClient(Cursor.Position);

                m.Items.Add("Delete Units");
                m.Items.Add("Close_AllWindows");
                m.Items.Add(new ToolStripSeparator());

                m.Items.Add("Drawn by Selected Parameter");


                m.ItemClicked += new ToolStripItemClickedEventHandler(m_ItemClicked);

                m.Show(advanced[index], relativeMousePosition);

            }
        }
        public void m_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            int Tabindex = tabControl1.SelectedIndex;

            string Query = "";



            Dictionary<string, CSV_Class.For_Box> SaveData = new Dictionary<string, CSV_Class.For_Box>();
            string[] id = new string[1];

            switch (e.ClickedItem.Text)
            {
                #region Drawn by Selected Parameter

                case "Drawn by Selected Parameter":


                    string[] Para = new string[advanced[Tabindex].SelectedCells.Count];


                    for (int k = 0; k < advanced[Tabindex].SelectedCells.Count; k++)
                    {
                        object Index = Data_Interface.New_Header[Convert.ToInt16(advanced[Tabindex].Rows[advanced[Tabindex].SelectedCells[k].RowIndex].Cells[0].Value) + 1];


                        int Find_DB = 0;
                        int ColumnLimit = Data_Interface.DB_Column_Limit;
                        if (Convert.ToInt16(advanced[Tabindex].Rows[advanced[Tabindex].SelectedCells[k].RowIndex].Cells[0].Value) + 1 >= ColumnLimit)
                        {
                            int Db_count = Data_Interface.DB_Count;
                            for (int q = 0; q < Db_count; q++)
                            {
                                int Column_end = Data_Interface.Per_DB_Column_Count_End[q];
                                if (Convert.ToInt16(advanced[Tabindex].Rows[advanced[Tabindex].SelectedCells[k].RowIndex].Cells[0].Value) + 1 <= Column_end)
                                {
                                    Find_DB = q;

                                    break;
                                }
                            }
                        }

                        int FindIndex = Convert.ToInt16(advanced[Tabindex].Rows[advanced[Tabindex].SelectedCells[k].RowIndex].Cells[0].Value);

                        int Db_Index = (FindIndex + 1) / (Data_Interface.DB_Column_Limit);

                        int Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
                        int Test_Index2 = FindIndex + 1 - Test_Index;

                        string[] TotalSample = new string[0];

                        for (int loop = 0; loop < Db_Interface.Table_Count; loop++)
                        {
                            string QueryTest = "Select " + Index + " from data" + loop + " where fail not like '1'";

                            //     string[] TotalSample = Db_Interface.Get_Data_By_Query(QueryTest, Find_DB);

                            string[] datas = Db_Interface.Get_Data_By_Query(QueryTest, Find_DB);

                            TotalSample = TotalSample.Concat(datas).ToArray();

                        }


                        string STD = Convert.ToString(Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[FindIndex + 1])].Std);
                        string Meadian = Convert.ToString(Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[FindIndex + 1])].Median_Data);
                        string Yeild = "";

                        long Test1 = Db_Interface.For_Any_Yield_For_New_Spec[Find_DB][Tabindex][Test_Index2];
                        long Test2 = TotalSample.Length;

                        long Test3 = Test1 + Test2;
                        long Test4 = Test2 - Test1;

                        if (Test3 == Test4)
                        {
                            Yeild = Convert.ToString(Convert.ToDouble(Test3) / Convert.ToDouble(Test4) * 100);

                        }
                        else
                        {
                            Yeild = Convert.ToString(Convert.ToDouble(Test4) / Convert.ToDouble(Test3) * 100);
                        }

                        string A_Min = Convert.ToString(Data_Interface.Customor_Clotho_List[FindIndex + 1].Min[Tabindex]);
                        string A_Max = Convert.ToString(Data_Interface.Customor_Clotho_List[FindIndex + 1].Max[Tabindex]);
                        string B_Min = Convert.ToString(Data_Interface.Clotho_List[FindIndex + 1].Min[Tabindex]);
                        string B_Max = Convert.ToString(Data_Interface.Clotho_List[FindIndex + 1].Max[Tabindex]);

                     //   CSV_Class.For_Box Set_Data = new CSV_Class.For_Box("", TotalSample, 0, 0, STD, Meadian, Yeild, A_Min, A_Max, B_Min, B_Max);

                    //    SaveData.Add(Data_Interface.Reference_Header[Convert.ToInt16(advanced[Tabindex].Rows[advanced[Tabindex].SelectedCells[k].RowIndex].Cells[0].Value) + 1], Set_Data);


                    }

                    id = new string[0];

                    for (int loop = 0; loop < Db_Interface.Table_Count; loop++)
                    {
                        string QueryTest1 = "Select id from data" + loop + " where fail not like '1'";
                        string[] datas = Db_Interface.Get_Data_By_Query(QueryTest1);

                        id = id.Concat(datas).ToArray();


                    }




                    CSV_Class.CSV CSV = new CSV_Class.CSV();
                    CSV_Class.CSV.INT CSV_Interface = CSV.Open(Key);
                    CSV_Interface.Write_Open("C:\\temp\\dummy\\BoxPlot.csv");


               //     CSV_Interface.ForBoxplotWrite("", id, SaveData,"");

                    CSV_Interface.Write_Close();

                    JMP_Draw_For_Boxplot("C:\\temp\\dummy\\BoxPlot.csv", SaveData, "");
                    break;

                #endregion

                #region Delete Units

                case "Delete Units":


                    bool falg = JMP_Interface.CheckDoc();
                    JMP_Interface.GetSelect_DataTable("dummy");

                    object Units = JMP_Interface.GetSelected_Gross_Row();


                    if (Units != null && Units != "")
                    {

                        //   DB_Interface.trans(Data_Interface);

                        if (Fail_Units == null)
                        {
                            Fail_Units = new string[0];
                        }
                        List<string> Sb = new List<string>();

                        foreach (object nb in (Array)Units)
                        {
                            Sb.Add(Convert.ToString(nb));
                        }
                        for (int k = 0; k < Db_Interface.Table_Count; k++)
                        {
                            Db_Interface.Gross_Update_Datas(Sb);
                        }


                        //  DB_Interface.Commit(Data_Interface);
                        long None_Sample_Count = 0;

                        for (int k = 0; k < Db_Interface.Table_Count; k++)
                        {
                            Query = "Select count(id) from data" + k + " where Fail like '1'";
                            None_Sample_Count += Db_Interface.Get_Sample_Count(0, Query);

                        }


                        Hidden_Sample_Count = None_Sample_Count;

                        JMP_Interface.Close_Dt("");

                        int Len = ((Array)Units).Length;

                        for (int k = 0; k < Db_Interface.Table_Count; k++)
                        {
                            Query = "Select id from data" + k + " where Fail like '1'";
                            string[] TotalSample = Db_Interface.Get_Data_By_Query(Query);

                            Fail_Units = Fail_Units.Concat(TotalSample).ToArray();

                        }

                        Cal_No_Thread_For_Delete_Unit(this.Total - Hidden_Sample_Count);
                    }
                    break;

                #endregion

                #region Close_AllWindows

                case "Close_AllWindows":

                    JMP_Interface.CloseWindowas();

                    break;

                    #endregion

            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                int index = tabControl1.SelectedIndex;
                _dataTable[s].DefaultView.Sort = "[No] ASC";
                bindingSource[s].Filter = "";

                advanced[s].CleanFilter();
                advanced[s].CleanSort();
            }

        }
        public void View()
        {

            for (int n = 0; n < tabControl1.TabCount; n++)
            {


                advanced[n].Columns[0].Width = 40;
                advanced[n].Columns[1].Width = 500;


                advanced[n].Columns[2].Width = 80;
                advanced[n].Columns[3].Width = 80;
                advanced[n].Columns[4].Width = 50;
                advanced[n].Columns[5].Width = 50;
                advanced[n].Columns[6].Width = 50;
                advanced[n].Columns[7].Width = 50;
                advanced[n].Columns[8].Width = 50;
                advanced[n].Columns[9].Width = 50;
                advanced[n].Columns[10].Width = 50;
                advanced[n].Columns[11].Width = 50;
                advanced[n].Columns[12].Width = 50;
                advanced[n].Columns[13].Width = 50;
                advanced[n].Columns[14].Width = 50;
                //advanced[n].Columns[15].Width = 50;
                //advanced[n].Columns[16].Width = 50;
                advanced[n].Columns[15].Width = 50;

                advanced[n].Columns[7].Frozen = true;
            }

            ForeColor();

        }
        public void EditonDatagrid(int No, int RefTab, int Tabindex, int ColumnIndex, int RowIndex)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            DataColumn[] dtkey = new DataColumn[1];

            int Spec_Index_dummy = Convert.ToInt16(advanced[RefTab].Rows[RowIndex].Cells[0].Value);

            dtkey[0] = _dataTable[RefTab].Columns["No"];
            _dataTable[RefTab].PrimaryKey = dtkey;

            DataRow dr = _dataTable[RefTab].Rows.Find(Spec_Index_dummy);
            int SelRow = _dataTable[RefTab].Rows.IndexOf(dr);

            int Index = Convert.ToInt16(_dataTable[RefTab].Rows[SelRow][0]);
            string Parameter = Convert.ToString(_dataTable[RefTab].Rows[SelRow][1]);

            double ChangedData = Convert.ToDouble(_dataTable[RefTab].Rows[SelRow][ColumnIndex]);
            int MinOrMax = 0; if (ColumnIndex == 2) MinOrMax = 1; else MinOrMax = 2;

            double Testtime24325 = TestTime1.Elapsed.TotalMilliseconds;

            int Header_Length = Data_Interface.Reference_Header.Length;


            int Spec_Index = Convert.ToInt16(advanced[RefTab].Rows[RowIndex].Cells[0].Value);
            int Db_Index = (Spec_Index + 1) / (Data_Interface.DB_Column_Limit);
            int Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
            int Test_Index2 = Spec_Index + 1 - Test_Index;

            Dictionary<string, double[]> Dic_For_Changed_Spec = Db_Interface.Chnaged_Spec_Anl_Yield(Db_Index, Test_Index2, Data_Interface.New_Header[Spec_Index + 1]);

            double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;

            double[] N_Spec = new double[2];


            N_Spec[0] = Convert.ToDouble(_dataTable[RefTab].Rows[SelRow][6]);
            N_Spec[1] = Convert.ToDouble(_dataTable[RefTab].Rows[SelRow][7]);
            //N_Spec[0] = Convert.ToDouble(advanced[Tabindext].Rows[RowIndex].Cells[6].Value);
            //N_Spec[1] = Convert.ToDouble(advanced[Tabindext].Rows[RowIndex].Cells[7].Value);


            double[] N_Data = Dic_For_Changed_Spec["DATA"];

            //  Dic[Data_Interface.New_Header[i]] = new List<int>();

            List<int> Count = new List<int>();
            int List_Count = 0;

            int Data_Length = N_Data.Length;
            int Fail_Count = 0;

            int getNb = Spec_Index - ((Data_Interface.DB_Column_Limit) * Db_Index) + 1;

            double Testtime2234 = TestTime1.Elapsed.TotalMilliseconds;

            Db_Interface.For_Any_Yield_For_New_Spec[Db_Index][RefTab][getNb] = N_Data.Length;


            double Testtime34344 = TestTime1.Elapsed.TotalMilliseconds;

            for (i = 0; i < Data_Length; i++)
            {

                if (N_Spec[1] < N_Data[i] || N_Spec[0] > N_Data[i])
                {
                    Fail_Count++;
                    List_Count = 1;
                }



                if (List_Count == 0)
                {
                    //  Db_Interface.For_New_Spec_ForCampare_Yield[Db_Index][i][RefTab][getNb] = 0;
                    Db_Interface.For_Any_Yield_For_New_Spec[Db_Index][RefTab][getNb]--;


                    var itemToRemove = Db_Interface.Yield_Test_New_Spec[Db_Index][i][RefTab].Find(r => r.Row == getNb);
                    if (itemToRemove != null)
                    {
                        Db_Interface.Yield_Test_New_Spec[Db_Index][i][RefTab].Remove(itemToRemove);
                    }

                }
                else
                {
                    //   Db_Interface.For_New_Spec_ForCampare_Yield[Db_Index][i][RefTab][getNb] = 1;
                    // Db_Interface.For_Any_Yield_For_New_Spec[Db_Index][RefTab][getNb]++;


                    var itemToRemove = Db_Interface.Yield_Test_New_Spec[Db_Index][i][RefTab].Find(r => r.Row == getNb);
                    if (itemToRemove == null)
                    {
                        DB_Class.DB_Editing.RowAndPass ss = new DB_Class.DB_Editing.RowAndPass(i, getNb, 1);
                        Db_Interface.Yield_Test_New_Spec[Db_Index][i][RefTab].Add(ss);
                    }

                }

                List_Count = 0;

            }
            double Testtime22222 = TestTime1.Elapsed.TotalMilliseconds;
            Count.Add(Fail_Count);


            double Min = 0f;
            double Max = 0f;
            double Avg = 0f;
            double L_CPK = 0f;
            double H_CPK = 0f;
            double Worst_CPK = 0f;
            double Median = 0f;
            double Stdev = 0f;

            STD(N_Data, N_Spec, out Min, out Max, out Avg, out L_CPK, out H_CPK, out Median, out Stdev, Spec_Index);

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

            if (L_CPK > H_CPK) Worst_CPK = H_CPK;
            else Worst_CPK = L_CPK;



            _dataTable[RefTab].Rows[SelRow][8] = Min;
            _dataTable[RefTab].Rows[SelRow][9] = Median;
            _dataTable[RefTab].Rows[SelRow][10] = Max;
            _dataTable[RefTab].Rows[SelRow][11] = Worst_CPK;
            _dataTable[RefTab].Rows[SelRow][12] = Stdev;

            double Testtime9999 = TestTime1.Elapsed.TotalMilliseconds;

            Sample_Verify = new int[Bin_Length];

            Cal_No_Thread(this.Total - Hidden_Total);

            double Testtime88 = TestTime1.Elapsed.TotalMilliseconds;

            //  For_New_Spec_Cal_Yield3(N_Data.Length - Hidden_Total);

            double Testtime11 = TestTime1.Elapsed.TotalMilliseconds;

            //     For_New_Spec_Cal_Yield2(N_Data.Length - Hidden_Total);

            double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

            double Yiled = ((Convert.ToDouble(Data_Length) - Convert.ToDouble(Fail_Count)) / Convert.ToDouble(Data_Length)) * 100;

            double Testtime5656 = TestTime1.Elapsed.TotalMilliseconds;

            _dataTable[RefTab].Rows[SelRow][13] = Math.Round(Yiled, 2);
            _dataTable[RefTab].Rows[SelRow][14] = Fail_Count;

            //  _dataTable[RefTab].Rows[No].EndEdit();



            //if (ColumnIndex == 4)
            //{
            //    Value2 = ChangedData;

            //    //    A_ForCtrlz_Min[tabControl1_index][Index] = Convert.ToDouble(advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            //    //    Cz = new forctrlz(Convert.ToInt16(advanced[tabControl1_index].Rows[e.RowIndex].Cells[0].Value), e.ColumnIndex, e.RowIndex, Data_Interface.Clotho_Spcc_List[Index].Min[tabControl1_index]);
            //    Data_Interface.New_Clotho_List[Index + 1].Min[Tabindext] = Value2;
            //    //    A_ForCtrlz_List[tabControl1_index].Add(Cz);
            //}
            //else
            //{
            //    Value2 = ChangedData;
            //    //    A_ForCtrlz_Max[tabControl1_index][Index] = Convert.ToDouble(advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            //    //    Cz = new forctrlz(Convert.ToInt16(advanced[tabControl1_index].Rows[e.RowIndex].Cells[0].Value), e.ColumnIndex, e.RowIndex, Data_Interface.Clotho_Spcc_List[Index].Max[tabControl1_index]);
            //    Data_Interface.New_Clotho_List[Index + 1].Max[Tabindext] = Value2;
            //    //    A_ForCtrlz_List[tabControl1_index].Add(Cz);
            //}


            //_dataTable[Tabindext].Rows[SelRow][ColumnIndex] = Convert.ToDouble(advanced[Tabindext].Rows[RowIndex].Cells[ColumnIndex].Value);
            double Testtime99 = TestTime1.Elapsed.TotalMilliseconds;
            bindingSource[RefTab].DataSource = _dataTable[RefTab];
            //  advanced[Tabindext].Update();


            double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;


        }
        public void EditonDatagrid_Sublot(int No, int RefTab, int Tabindex, int ColumnIndex, int RowIndex)
        {

            DataColumn[] dtkey = new DataColumn[1];

            int Spec_Index = Convert.ToInt16(advanced[RefTab].Rows[RowIndex].Cells[0].Value);

            dtkey[0] = _dataTable[Tabindex].Columns["No"];
            _dataTable[Tabindex].PrimaryKey = dtkey;

            DataRow dr = _dataTable[Tabindex].Rows.Find(Spec_Index);
            int SelRow = _dataTable[Tabindex].Rows.IndexOf(dr);

            int Index = Convert.ToInt16(_dataTable[Tabindex].Rows[SelRow][0]);
            string Parameter = Convert.ToString(_dataTable[Tabindex].Rows[SelRow][1]);

            double ChangedData = Convert.ToDouble(_dataTable[RefTab].Rows[SelRow][ColumnIndex]);
            int MinOrMax = 0; if (ColumnIndex == 2) MinOrMax = 1; else MinOrMax = 2;

            int Header_Length = Data_Interface.Reference_Header.Length;



            int Db_Index = (Spec_Index + 1) / (Data_Interface.DB_Column_Limit);
            int Test_Index = Data_Interface.Per_DB_Column_Count_Start[Db_Index];
            int Test_Index2 = Spec_Index + 1 - Test_Index;

            Dictionary<string, double[]> Dic_For_Changed_Spec = Db_Interface.Chnaged_Spec_Anl_Yield(Db_Index, Test_Index2, Data_Interface.New_Header[Spec_Index + 1]);

            double[] N_Spec = new double[2];


            N_Spec[0] = Convert.ToDouble(_dataTable[Tabindex].Rows[SelRow][6]);
            N_Spec[1] = Convert.ToDouble(_dataTable[Tabindex].Rows[SelRow][7]);
            //N_Spec[0] = Convert.ToDouble(advanced[Tabindext].Rows[RowIndex].Cells[6].Value);
            //N_Spec[1] = Convert.ToDouble(advanced[Tabindext].Rows[RowIndex].Cells[7].Value);


            double[] N_Data = Dic_For_Changed_Spec["DATA"];

            //  Dic[Data_Interface.New_Header[i]] = new List<int>();

            List<int> Count = new List<int>();
            int List_Count = 0;

            int Data_Length = N_Data.Length;
            int Fail_Count = 0;

            int getNb = Spec_Index - ((Data_Interface.DB_Column_Limit) * Db_Index) + 1;

            Db_Interface.For_Any_Yield_For_New_Spec[Db_Index][Tabindex][getNb] = N_Data.Length;

            for (i = 0; i < Data_Length; i++)
            {

                if (N_Spec[1] < N_Data[i] || N_Spec[0] > N_Data[i])
                {
                    Fail_Count++;
                    List_Count = 1;
                }


                if (List_Count == 0)
                {
                    //    Db_Interface.For_New_Spec_ForCampare_Yield[Db_Index][i][RefTab][getNb] = 0;
                    //    Db_Interface.For_Any_Yield_Percent_For_New_Spec[Db_Index][i][RefTab][0] = 0;
                    Db_Interface.For_Any_Yield_For_New_Spec[Db_Index][Tabindex][getNb]--;

                    var itemToRemove = Db_Interface.Yield_Test_New_Spec[Db_Index][i][Tabindex].Find(r => r.Row == getNb);
                    if (itemToRemove != null)
                    {
                        Db_Interface.Yield_Test_New_Spec[Db_Index][i][Tabindex].Remove(itemToRemove);
                    }

                }
                else
                {
                    //    Db_Interface.For_New_Spec_ForCampare_Yield[Db_Index][i][RefTab][getNb] = 1;
                    //    Db_Interface.For_Any_Yield_Percent_For_New_Spec[Db_Index][i][RefTab][0] = 1;

                    //  Db_Interface.For_Any_Yield_For_New_Spec[Db_Index][Tabindex][getNb] = 1;
                    var itemToRemove = Db_Interface.Yield_Test_New_Spec[Db_Index][i][Tabindex].Find(r => r.Row == getNb);
                    if (itemToRemove == null)
                    {
                        DB_Class.DB_Editing.RowAndPass ss = new DB_Class.DB_Editing.RowAndPass(i, getNb, 1);
                        Db_Interface.Yield_Test_New_Spec[Db_Index][i][Tabindex].Add(ss);
                    }

                }
                List_Count = 0;

            }

            Count.Add(Fail_Count);


            double Min = 0f;
            double Max = 0f;
            double Avg = 0f;
            double L_CPK = 0f;
            double H_CPK = 0f;
            double Worst_CPK = 0f;
            double Median = 0f;
            double Stdev = 0f;

            STD(N_Data, N_Spec, out Min, out Max, out Avg, out L_CPK, out H_CPK, out Median, out Stdev, Spec_Index);



            if (L_CPK > H_CPK) Worst_CPK = H_CPK;
            else Worst_CPK = L_CPK;


            _dataTable[Tabindex].Rows[SelRow][8] = Min;
            _dataTable[Tabindex].Rows[SelRow][9] = Median;
            _dataTable[Tabindex].Rows[SelRow][10] = Max;
            _dataTable[Tabindex].Rows[SelRow][11] = Worst_CPK;
            _dataTable[Tabindex].Rows[SelRow][12] = Stdev;


            Sample_Verify = new int[Bin_Length];

            Cal_No_Thread(this.Total - Hidden_Total);

            double Yiled = ((Convert.ToDouble(Data_Length) - Convert.ToDouble(Fail_Count)) / Convert.ToDouble(Data_Length)) * 100;

            _dataTable[Tabindex].Rows[SelRow][13] = Math.Round(Yiled, 2);
            _dataTable[Tabindex].Rows[SelRow][14] = Fail_Count;




            //if (ColumnIndex == 4)
            //{
            //    Value2 = ChangedData;

            //    //    A_ForCtrlz_Min[tabControl1_index][Index] = Convert.ToDouble(advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            //    //    Cz = new forctrlz(Convert.ToInt16(advanced[tabControl1_index].Rows[e.RowIndex].Cells[0].Value), e.ColumnIndex, e.RowIndex, Data_Interface.Clotho_Spcc_List[Index].Min[tabControl1_index]);
            //    Data_Interface.New_Clotho_List[Index + 1].Min[Tabindext] = Value2;
            //    //    A_ForCtrlz_List[tabControl1_index].Add(Cz);
            //}
            //else
            //{
            //    Value2 = ChangedData;
            //    //    A_ForCtrlz_Max[tabControl1_index][Index] = Convert.ToDouble(advanced[tabControl1_index].Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            //    //    Cz = new forctrlz(Convert.ToInt16(advanced[tabControl1_index].Rows[e.RowIndex].Cells[0].Value), e.ColumnIndex, e.RowIndex, Data_Interface.Clotho_Spcc_List[Index].Max[tabControl1_index]);
            //    Data_Interface.New_Clotho_List[Index + 1].Max[Tabindext] = Value2;
            //    //    A_ForCtrlz_List[tabControl1_index].Add(Cz);
            //}


            //_dataTable[Tabindext].Rows[SelRow][ColumnIndex] = Convert.ToDouble(advanced[Tabindext].Rows[RowIndex].Cells[ColumnIndex].Value);

            bindingSource[Tabindex].DataSource = _dataTable[Tabindex];


        }
        public void EditColumnIndex2and3(int TabIndex, int ColumnIndex, int RowIndex, string Key)
        {
            string Para_Text = advanced[TabIndex].Rows[RowIndex].Cells[ColumnIndex].Value.ToString();
            int No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);

            if (Para_Text == "CUSTOMER")
            {
                #region
                for (int length = 0; length < 1; length++)
                {
                    No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);

                    string Value = datagrid[0][Col + length, Row].Value.ToString();
                    string Test = datagrid[0][TabIndex + 1, Row].Value.ToString();
                    object data = null;

                    DataColumn[] dtkey = new DataColumn[1];

                    dtkey[0] = _dataTable[TabIndex].Columns["No"];
                    _dataTable[TabIndex].PrimaryKey = dtkey;

                    DataRow dr = _dataTable[TabIndex].Rows.Find(No);
                    int SelRow = _dataTable[TabIndex].Rows.IndexOf(dr);

                    data = advanced[TabIndex].Rows[RowIndex].Cells[ColumnIndex].Value;
                    _dataTable[TabIndex].Rows[SelRow][ColumnIndex] = data;

                    if (ColumnIndex == 2)
                    {
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 2] = Data_Interface.Clotho_List[No + 1].Min[TabIndex];
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 4] = Data_Interface.Clotho_List[No + 1].Min[TabIndex];
                        bindingSource[TabIndex].DataSource = _dataTable[TabIndex];
                    }
                    else if (ColumnIndex == 3)
                    {

                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 2] = Data_Interface.Clotho_List[No + 1].Max[TabIndex];
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 4] = Data_Interface.Clotho_List[No + 1].Max[TabIndex];
                        bindingSource[TabIndex].DataSource = _dataTable[TabIndex];
                    }

                    Calculate_By_Customer_Spec(TabIndex, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);

                    Flag = true;
                    for (int w = tabControl1.SelectedIndex; w < Bin_Length - 1; w++)
                    {
                        Test = datagrid[0][w + 1, Row].Value.ToString();
                        if (Test != "9999")
                        {
                            No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);
                            dtkey = new DataColumn[1];

                            dtkey[0] = _dataTable[Convert.ToInt16(Test)].Columns["No"];
                            _dataTable[Convert.ToInt16(Test)].PrimaryKey = dtkey;

                            dr = _dataTable[Convert.ToInt16(Test)].Rows.Find(No);
                            SelRow = _dataTable[Convert.ToInt16(Test)].Rows.IndexOf(dr);

                            data = advanced[w].Rows[RowIndex].Cells[ColumnIndex].Value;
                            _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex] = data;



                            // bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];

                            if (ColumnIndex == 2)
                            {
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 2] = Data_Interface.Clotho_List[No + 1].Min[Convert.ToInt16(Test)];
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 4] = Data_Interface.Clotho_List[No + 1].Min[Convert.ToInt16(Test)];
                                bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];
                            }
                            else if (ColumnIndex == 3)
                            {
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 2] = Data_Interface.Clotho_List[No + 1].Max[Convert.ToInt16(Test)];
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 4] = Data_Interface.Clotho_List[No + 1].Max[Convert.ToInt16(Test)];
                                bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];
                            }
                            Calculate_By_Customer_Spec_Sublot(TabIndex, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);

                            // _dataTable[Convert.ToInt16(Test)].AcceptChanges();
                        }
                    }

                }
                Flag = false;
                #endregion
            }
            else if (Para_Text == "FIXEDPOUT")
            {
                for (int length = 0; length < 1; length++)
                {
                    No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);

                    string Value = datagrid[0][Col + length, Row].Value.ToString();
                    string Test = datagrid[0][TabIndex + 1, Row].Value.ToString();
                    object data = null;

                    DataColumn[] dtkey = new DataColumn[1];

                    dtkey[0] = _dataTable[TabIndex].Columns["No"];
                    _dataTable[TabIndex].PrimaryKey = dtkey;

                    DataRow dr = _dataTable[TabIndex].Rows.Find(No);
                    int SelRow = _dataTable[TabIndex].Rows.IndexOf(dr);

                    data = advanced[TabIndex].Rows[RowIndex].Cells[ColumnIndex].Value;
                    _dataTable[TabIndex].Rows[SelRow][ColumnIndex] = data;

                    if (ColumnIndex == 2)
                    {
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 2] = -9999;
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 4] = -9999;
                        bindingSource[TabIndex].DataSource = _dataTable[TabIndex];

                        Data_Interface.Customor_Clotho_List[No + 1].Min[TabIndex] = -9999;


                    }
                    else if (ColumnIndex == 3)
                    {
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 2] = 9999;
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 4] = 9999;
                        bindingSource[TabIndex].DataSource = _dataTable[TabIndex];

                        Data_Interface.Customor_Clotho_List[No + 1].Max[TabIndex] = 9999;
                    }

                    //  Calculate_By_Customer_Spec(TabIndex, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);


                    EditonDatagrid(No, TabIndex, 0, ColumnIndex + 2, RowIndex);

                    Flag = true;
                    for (int w = tabControl1.SelectedIndex; w < Bin_Length - 1; w++)
                    {

                        Test = datagrid[0][w + 1, Row].Value.ToString();

                        if (Test != "9999")
                        {
                            No = Convert.ToInt16(advanced[Convert.ToInt16(Test)].Rows[RowIndex].Cells[0].Value);

                            dtkey = new DataColumn[1];

                            dtkey[0] = _dataTable[Convert.ToInt16(Test)].Columns["No"];
                            _dataTable[Convert.ToInt16(Test)].PrimaryKey = dtkey;

                            dr = _dataTable[Convert.ToInt16(Test)].Rows.Find(No);
                            SelRow = _dataTable[Convert.ToInt16(Test)].Rows.IndexOf(dr);

                            data = advanced[w].Rows[RowIndex].Cells[ColumnIndex].Value;
                            _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex] = data;




                            if (ColumnIndex == 2)
                            {
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 2] = -9999;
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 4] = -9999;
                                bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];

                                Data_Interface.Clotho_List[No + 1].Min[Convert.ToInt16(Test)] = -9999;
                            }
                            else if (ColumnIndex == 3)
                            {

                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 2] = 9999;
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 4] = 9999;
                                bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];

                                Data_Interface.Clotho_List[No + 1].Max[Convert.ToInt16(Test)] = 9999;
                            }

                            EditonDatagrid_Sublot(No, TabIndex, Convert.ToInt16(Test), ColumnIndex + 2, RowIndex);
                        }
                    }




                }
                Flag = false;
            }
            else
            {
                #region
                for (int length = 0; length < 1; length++)
                {
                    No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);

                    string Value = datagrid[0][Col + length, Row].Value.ToString();
                    string Test = datagrid[0][TabIndex + 1, Row].Value.ToString();
                    object data = null;

                    DataColumn[] dtkey = new DataColumn[1];

                    dtkey[0] = _dataTable[TabIndex].Columns["No"];
                    _dataTable[TabIndex].PrimaryKey = dtkey;

                    DataRow dr = _dataTable[TabIndex].Rows.Find(No);
                    int SelRow = _dataTable[TabIndex].Rows.IndexOf(dr);

                    data = advanced[TabIndex].Rows[RowIndex].Cells[ColumnIndex].Value;
                    _dataTable[TabIndex].Rows[SelRow][ColumnIndex] = data;

                    if (ColumnIndex == 2)
                    {
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 2] = -9999;
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 4] = -9999;
                        bindingSource[TabIndex].DataSource = _dataTable[TabIndex];

                        Data_Interface.Customor_Clotho_List[No + 1].Min[TabIndex] = -9999;


                    }
                    else if (ColumnIndex == 3)
                    {
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 2] = 9999;
                        _dataTable[TabIndex].Rows[SelRow][ColumnIndex + 4] = 9999;
                        bindingSource[TabIndex].DataSource = _dataTable[TabIndex];

                        Data_Interface.Customor_Clotho_List[No + 1].Max[TabIndex] = 9999;
                    }

                    //  Calculate_By_Customer_Spec(TabIndex, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);


                    EditonDatagrid(No, TabIndex, 0, ColumnIndex + 2, RowIndex);

                    Flag = true;
                    for (int w = tabControl1.SelectedIndex; w < Bin_Length - 1; w++)
                    {

                        Test = datagrid[0][w + 1, Row].Value.ToString();

                        if (Test != "9999")
                        {
                            No = Convert.ToInt16(advanced[Convert.ToInt16(Test)].Rows[RowIndex].Cells[0].Value);

                            dtkey = new DataColumn[1];

                            dtkey[0] = _dataTable[Convert.ToInt16(Test)].Columns["No"];
                            _dataTable[Convert.ToInt16(Test)].PrimaryKey = dtkey;

                            dr = _dataTable[Convert.ToInt16(Test)].Rows.Find(No);
                            SelRow = _dataTable[Convert.ToInt16(Test)].Rows.IndexOf(dr);

                            data = advanced[w].Rows[RowIndex].Cells[ColumnIndex].Value;
                            _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex] = data;




                            if (ColumnIndex == 2)
                            {
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 2] = -9999;
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 4] = -9999;
                                bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];

                                Data_Interface.Customor_Clotho_List[No + 1].Min[Convert.ToInt16(Test)] = -9999;
                            }
                            else if (ColumnIndex == 3)
                            {

                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 2] = 9999;
                                _dataTable[Convert.ToInt16(Test)].Rows[SelRow][ColumnIndex + 4] = 9999;
                                bindingSource[Convert.ToInt16(Test)].DataSource = _dataTable[Convert.ToInt16(Test)];

                                Data_Interface.Customor_Clotho_List[No + 1].Max[Convert.ToInt16(Test)] = 9999;
                            }

                            EditonDatagrid_Sublot(No, TabIndex, Convert.ToInt16(Test), ColumnIndex + 2, RowIndex);
                        }
                    }




                }
                Flag = false;
                #endregion
            }

        }
        public void EditColumnIndex4and5(int TabIndex, int ColumnIndex, int RowIndex, string Key)
        {

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            string Para_Text = advanced[TabIndex].Rows[RowIndex].Cells[1].Value.ToString();
            int No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);


            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

            for (int length = 0; length < 1; length++)
            {
                No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);

                string Value = datagrid[0][Col + length, Row].Value.ToString();
                string Test = datagrid[0][TabIndex + 1, Row].Value.ToString();


                DataColumn[] dtkey = new DataColumn[1];

                dtkey[0] = _dataTable[TabIndex].Columns["No"];
                _dataTable[TabIndex].PrimaryKey = dtkey;

                DataRow dr = _dataTable[TabIndex].Rows.Find(No);
                int SelRow = _dataTable[TabIndex].Rows.IndexOf(dr);

                _dataTable[TabIndex].Rows[SelRow].BeginEdit();

                if (Test != "9999")
                {
                    Calculate_By_Key_Column4_5(TabIndex, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);
                }
                else
                {
                    Calculate_By_Key_Column4_5(TabIndex, TabIndex, SelRow, ColumnIndex, RowIndex, Key);
                }
                double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

                Flag = true;
                for (int w = tabControl1.SelectedIndex; w < Bin_Length - 1; w++)
                {
                    Test = datagrid[0][w + 1, Row].Value.ToString();

                    if (Test != "9999")
                    {
                        No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);
                        dtkey = new DataColumn[1];

                        dtkey[0] = _dataTable[Convert.ToInt16(Test)].Columns["No"];
                        _dataTable[Convert.ToInt16(Test)].PrimaryKey = dtkey;

                        dr = _dataTable[Convert.ToInt16(Test)].Rows.Find(No);
                        SelRow = _dataTable[Convert.ToInt16(Test)].Rows.IndexOf(dr);

                        _dataTable[Convert.ToInt16(Test)].Rows[SelRow].BeginEdit();

                        Calculate_By_Key_Column4_5_SubLot(Convert.ToInt16(Test) - 1, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);

                    }
                }
                double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                Flag = false;

            }

            double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void EditColumnIndex15and16(int TabIndex, int ColumnIndex, int RowIndex, string Key)
        {

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            string Para_Text = advanced[TabIndex].Rows[RowIndex].Cells[1].Value.ToString();
            int No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);


            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;

            for (int length = 0; length < 1; length++)
            {
                No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);

                string Value = datagrid[0][Col + length, Row].Value.ToString();
                string Test = datagrid[0][TabIndex + 1, Row].Value.ToString();


                DataColumn[] dtkey = new DataColumn[1];

                dtkey[0] = _dataTable[TabIndex].Columns["No"];
                _dataTable[TabIndex].PrimaryKey = dtkey;

                DataRow dr = _dataTable[TabIndex].Rows.Find(No);
                int SelRow = _dataTable[TabIndex].Rows.IndexOf(dr);

                _dataTable[TabIndex].Rows[SelRow].BeginEdit();

                if (Test != "9999")
                {
                    Calculate_By_Key_Column15_16(TabIndex, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);
                }
                else
                {
                    Calculate_By_Key_Column15_16(TabIndex, TabIndex, SelRow, ColumnIndex, RowIndex, Key);
                }
                double Testtime2 = TestTime1.Elapsed.TotalMilliseconds;

                Flag = true;
                for (int w = tabControl1.SelectedIndex; w < Bin_Length - 1; w++)
                {
                    Test = datagrid[0][w + 1, Row].Value.ToString();

                    if (Test != "9999")
                    {
                        //No = Convert.ToInt16(advanced[TabIndex].Rows[RowIndex].Cells[0].Value);
                        //dtkey = new DataColumn[1];

                        //dtkey[0] = _dataTable[Convert.ToInt16(Test)].Columns["No"];
                        //_dataTable[Convert.ToInt16(Test)].PrimaryKey = dtkey;

                        //dr = _dataTable[Convert.ToInt16(Test)].Rows.Find(No);
                        //SelRow = _dataTable[Convert.ToInt16(Test)].Rows.IndexOf(dr);

                        //_dataTable[Convert.ToInt16(Test)].Rows[SelRow].BeginEdit();

                        //Calculate_By_Key_Column15_16_SubLot(Convert.ToInt16(Test) - 1, Convert.ToInt16(Test), SelRow, ColumnIndex, RowIndex, Key);

                    }
                }
                double Testtime3 = TestTime1.Elapsed.TotalMilliseconds;
                Flag = false;

            }



            double Testtime4 = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void Calculate_By_Key_Column4_5(int refTab, int Forndex, int No, int ColumnIndex, int RowIndex, string Key)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            int Spec_Index = Convert.ToInt16(advanced[refTab].Rows[RowIndex].Cells[0].Value);

            object data = null;
            if (Key == "CPK")
            {
                if (ColumnIndex == 4)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Avg[refTab] - (Convert.ToDouble(_dataTable[refTab].Rows[No][ColumnIndex]) * 3 * Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Std[refTab]);

                }
                else if (ColumnIndex == 5)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Avg[refTab] + (Convert.ToDouble(_dataTable[refTab].Rows[No][ColumnIndex]) * 3 * Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Std[refTab]);

                }


                _dataTable[refTab].Rows[No][ColumnIndex] = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[refTab].Rows[No][ColumnIndex + 2] = data;

            }
            else if (Key == "MANUAL")
            {
                data = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[refTab].Rows[No][ColumnIndex] = data;
                _dataTable[refTab].Rows[No][ColumnIndex + 2] = data;

            }
            else if (Key == "RANGE")
            {
                if (ColumnIndex == 4)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Median_Data[refTab] - Convert.ToDouble(_dataTable[refTab].Rows[No][ColumnIndex]);
                }
                else if (ColumnIndex == 5)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Median_Data[refTab] + Convert.ToDouble(_dataTable[refTab].Rows[No][ColumnIndex]);
                }


                _dataTable[refTab].Rows[No][ColumnIndex] = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[refTab].Rows[No][ColumnIndex + 2] = data;
            }
            else if (Key == "CUSTOMER")
            {
                data = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[refTab].Rows[No][ColumnIndex] = data;
                _dataTable[refTab].Rows[No][ColumnIndex + 2] = data;
            }
            else if (Key == "FIXEDPOUT")
            {
                string[] Paraname = advanced[refTab].Rows[RowIndex].Cells[1].Value.ToString().Split('_');

                double Pout = Convert.ToDouble(Paraname[7].Replace("dBm", "").Trim());

                double PoutValue = Convert.ToDouble(advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value);

                if (ColumnIndex == 4)
                {
                    data = Pout - PoutValue;
                }
                else if (ColumnIndex == 5)
                {
                    data = Pout + PoutValue;
                }

                _dataTable[refTab].Rows[No][ColumnIndex] = PoutValue;
                _dataTable[refTab].Rows[No][ColumnIndex + 2] = data;
            }



            if (ColumnIndex == 4)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Min[refTab] = Convert.ToDouble(data);
            }
            else if (ColumnIndex == 5)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Max[refTab] = Convert.ToDouble(data);
            }

            bindingSource[refTab].DataSource = _dataTable[refTab];


            EditonDatagrid(No, refTab, Forndex, ColumnIndex, RowIndex);
            //  advanced[refTab].Update();
            double Testtime = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void Calculate_By_Key_Column4_5_SubLot(int refTab, int Forndex, int No, int ColumnIndex, int RowIndex, string Key)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            int Spec_Index = Convert.ToInt16(advanced[refTab].Rows[RowIndex].Cells[0].Value);


            DataRow dr = _dataTable[refTab].Rows.Find(Spec_Index);
            int Ref_SelRow = _dataTable[refTab].Rows.IndexOf(dr);


            object data = null;
            if (Key == "CPK")
            {
                if (ColumnIndex == 4)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Avg[refTab] - (Convert.ToDouble(advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value) * 3 * Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Std[refTab]);
                }
                else if (ColumnIndex == 5)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Avg[refTab] + (Convert.ToDouble(advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value) * 3 * Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Std[refTab]);
                }


                _dataTable[Forndex].Rows[No][ColumnIndex] = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[Forndex].Rows[No][ColumnIndex + 2] = data;

            }
            else if (Key == "MANUAL")
            {
                data = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[Forndex].Rows[No][ColumnIndex] = data;
                _dataTable[Forndex].Rows[No][ColumnIndex + 2] = data;
            }
            else if (Key == "RANGE")
            {
                if (ColumnIndex == 4)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Median_Data[refTab] - Convert.ToDouble(_dataTable[refTab].Rows[Ref_SelRow][ColumnIndex]);
                }
                else if (ColumnIndex == 5)
                {
                    data = Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Data_Interface.Reference_Header[Spec_Index + 1]].Median_Data[refTab] + Convert.ToDouble(_dataTable[refTab].Rows[Ref_SelRow][ColumnIndex]);
                }


                _dataTable[Forndex].Rows[No][ColumnIndex] = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[Forndex].Rows[No][ColumnIndex + 2] = data;
            }
            else if (Key == "CUSTOMER")
            {
                data = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value;
                _dataTable[Forndex].Rows[No][ColumnIndex] = data;
                _dataTable[Forndex].Rows[No][ColumnIndex + 2] = data;
            }
            else if (Key == "FIXEDPOUT")
            {
                string[] Paraname = advanced[refTab].Rows[RowIndex].Cells[1].Value.ToString().Split('_');

                double Pout = Convert.ToDouble(Paraname[7].Replace("dBm", "").Trim());

                double PoutValue = Convert.ToDouble(advanced[refTab].Rows[RowIndex].Cells[ColumnIndex].Value);

                if (ColumnIndex == 4)
                {
                    data = Pout - PoutValue;
                }
                else if (ColumnIndex == 5)
                {
                    data = Pout + PoutValue;
                }

                _dataTable[Forndex].Rows[No][ColumnIndex] = PoutValue;
                _dataTable[Forndex].Rows[No][ColumnIndex + 2] = data;
            }


            if (ColumnIndex == 4)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Min[Forndex] = Convert.ToDouble(data);
            }
            else if (ColumnIndex == 5)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Max[Forndex] = Convert.ToDouble(data);
            }

            bindingSource[Forndex].DataSource = _dataTable[Forndex];

            EditonDatagrid_Sublot(No, refTab, Forndex, ColumnIndex, RowIndex);
            //  advanced[Forndex].Update();
            double Testtime = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void Calculate_By_Key_Column15_16(int refTab, int Forndex, int No, int ColumnIndex, int RowIndex, string Key)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            int Spec_Index = Convert.ToInt16(advanced[refTab].Rows[RowIndex].Cells[0].Value);



            _dataTable[refTab].Rows[No][17] = 0;

            bindingSource[refTab].DataSource = _dataTable[refTab];

            Outlier_List = new List<string>();

            Db_Interface.DIC_IQR[Data_Interface.Reference_Header[Spec_Index + 1]].SN = new string[0];

            for (int outlier = 0; outlier < Db_Interface.DIC_IQR.Count; outlier++)
            {
                string[] dummy = Db_Interface.DIC_IQR[Data_Interface.Reference_Header[outlier]].SN;

                for (int dummycount = 0; dummycount < dummy.Length; dummycount++)
                {
                    if (!Outlier_List.Contains(dummy[dummycount]))
                    {
                        Outlier_List.Add(dummy[dummycount]);
                    }

                }

            }



            datagrid2[0].Rows[6].Cells[1].Value = Outlier_List.Count;

            //  advanced[refTab].Update();
            double Testtime = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void Calculate_By_Key_Column15_16_SubLot(int refTab, int Forndex, int No, int ColumnIndex, int RowIndex, string Key)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            int Spec_Index = Convert.ToInt16(advanced[refTab].Rows[RowIndex].Cells[0].Value);


            DataRow dr = _dataTable[refTab].Rows.Find(Spec_Index);
            int Ref_SelRow = _dataTable[refTab].Rows.IndexOf(dr);





            _dataTable[Forndex].Rows[No][17] = 0;

            bindingSource[Forndex].DataSource = _dataTable[Forndex];

            Db_Interface.DIC_IQR[Data_Interface.Reference_Header[Spec_Index + 1]].SN = new string[0];

            //for (int length = 0; length < 1; length++)
            //{
            //    datagrid2[0].Rows[6].Cells[1].Value = 0;
            //}




            //  EditonDatagrid_Sublot(No, refTab, Forndex, ColumnIndex, RowIndex);
            //  advanced[Forndex].Update();
            double Testtime = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void Calculate_By_Customer_Spec(int refTab, int Forndex, int No, int ColumnIndex, int RowIndex, string Key)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            int Spec_Index = Convert.ToInt16(advanced[refTab].Rows[RowIndex].Cells[0].Value);

            object data = null;

            if (ColumnIndex == 2)
            {
                data = Data_Interface.Clotho_List[Spec_Index + 1].Min[0];
                _dataTable[refTab].Rows[No][ColumnIndex + 2] = data;
            }
            else
            {
                data = Data_Interface.Clotho_List[Spec_Index + 1].Max[0];
                _dataTable[refTab].Rows[No][ColumnIndex + 2] = data;

            }




            if (ColumnIndex == 2)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Min[refTab] = Convert.ToDouble(data);
            }
            else if (ColumnIndex == 3)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Max[refTab] = Convert.ToDouble(data);
            }




            EditonDatagrid(No, refTab, Forndex, ColumnIndex + 2, RowIndex);
            bindingSource[refTab].DataSource = _dataTable[refTab];
            //  advanced[refTab].Update();
            double Testtime = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void Calculate_By_Customer_Spec_Sublot(int refTab, int Forndex, int No, int ColumnIndex, int RowIndex, string Key)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();


            int Spec_Index = Convert.ToInt16(advanced[refTab].Rows[RowIndex].Cells[0].Value);

            object data = null;

            if (ColumnIndex == 2)
            {
                data = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex + 2].Value;
                //    data = Data_Interface.Customor_Clotho_List[Spec_Index + 1].Min[0];
                _dataTable[Forndex].Rows[No][ColumnIndex + 2] = data;
            }
            else
            {
                data = advanced[refTab].Rows[RowIndex].Cells[ColumnIndex + 2].Value;
                //  data = Data_Interface.Customor_Clotho_List[Spec_Index + 1].Max[0];
                _dataTable[Forndex].Rows[No][ColumnIndex + 2] = data;

            }


            if (ColumnIndex == 2)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Min[Forndex] = Convert.ToDouble(data);
            }
            else if (ColumnIndex == 3)
            {
                Data_Interface.Customor_Clotho_List[Spec_Index + 1].Max[Forndex] = Convert.ToDouble(data);
            }

            bindingSource[Forndex].DataSource = _dataTable[Forndex];

            EditonDatagrid_Sublot(No, refTab, Forndex, ColumnIndex + 2, RowIndex);
            //  advanced[Forndex].Update();
            double Testtime = TestTime1.Elapsed.TotalMilliseconds;
        }
        private void JMP_Draw(string FilePaht, Dictionary<string, CSV_Class.For_Box> Data, string By, bool Save_Falg)
        {
            JMP_Interface.Open_Session(true);

            JMP_File = FilePaht;

            JMP_Interface.Open_Document(FilePaht);
            JMP_Interface.GetDataTable();

            JMP_Class.Script Distribution_Script;

            Distribution_Script = null;
            List<string>[] Para_Test = new List<string>[0];
            Dictionary<int, Dictionary<int, string>> dummy = new Dictionary<int, Dictionary<int, string>>();
            switch (By)
            {
                case "LOT":
                    Db_Interface.Variation = DupCheck<object>(Db_Interface.Variation);
                 //   Distribution_Script = JMP_Interface.Make_Script("DISTRIBUTION", Data);
                    break;
                case "SITE":
               //     Db_Interface.Variation = DupCheck<object>(Db_Interface.Variation);
                  //  Distribution_Script = JMP_Interface.Make_Script("DISTRIBUTION", Parameter, Parameter2, Db_Interface.Value, Db_Interface.Variation, null, "SITE", false, false, false, 0f);
                    break;
                case "BIN":
               //     Db_Interface.Variation = DupCheck<object>(Db_Interface.Variation);
               //     Distribution_Script = JMP_Interface.Make_Script("DISTRIBUTION", Parameter, Parameter2, Db_Interface.Value, Db_Interface.Variation, Spec, "BIN", false, false, false, 0f);
                    break;

                default:

                 //   Distribution_Script = JMP_Interface.Make_Script("DISTRIBUTION", Parameter, Db_Interface.Value, dummy, Save_Falg, ref Para_Test, false, false, false, 0f);

                    break;

            }


            //if (ForGross_Fail_Unit.Count != 0)
            //{
            //    JMP_Class.Script Distribution_HideAndExclude = JMP_Interface.Distribution_HideAndExclude_1("DISTRIBUTION_HideAndExclude_1", ForGross_Fail_Unit);

            //    Csv_Interface.Write_Open("C:\\temp\\dummy\\dummy3.jsl");
            //    Csv_Interface.WriteScript(Distribution_HideAndExclude.Scrip_Data);
            //    Csv_Interface.Write_Close();

            //    JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy3.jsl");

            //}

            Csv_Interface.Write_Open("C:\\temp\\dummy\\dummy.jsl");
            Csv_Interface.WriteScript(Distribution_Script.Scrip_Data);
            Csv_Interface.Write_Close();

            JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy.jsl");


        }
        private void JMP_Draw_For_Boxplot(string FilePaht, Dictionary<string, CSV_Class.For_Box> Data, string FilePath)
        {
            JMP_Interface.Open_Session(true);

            JMP_File = FilePaht;

            JMP_Interface.Open_Document(FilePaht);
            JMP_Interface.GetDataTable();

            JMP_Class.Script Distribution_Script;

            Distribution_Script = null;

            List<string>[] Para_Test = new List<string>[0];
            Dictionary<int, Dictionary<int, string>> dummy = new Dictionary<int, Dictionary<int, string>>();
            Distribution_Script = JMP_Interface.Make_Script("VARIABLILITY", Data,null, FilePaht, dummy, false , ref Para_Test,false,false,false, 0f);

            Csv_Interface.Write_Open("C:\\temp\\dummy\\BOXPLOT.jsl");
            Csv_Interface.WriteScript(Distribution_Script.Scrip_Data);
            Csv_Interface.Write_Close();

            JMP_Interface.Run_Script("C:\\temp\\dummy\\BOXPLOT.jsl");


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
        public void For_New_Spec_TestResult_Cal()
        {

            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            Db_Interface.For_New_Spec_Cal_Value_by_rowsdata = new Dictionary<string, DB_Class.DB_Editing.Data_Calculation>();
            double[] dummy_Test = new double[11];

            for (int j = 0; j < Data_Interface.New_Header.Length; j++)
            {
                Db_Interface.For_New_Spec_Cal_Value_by_rowsdata.Add(Data_Interface.Ref_New_Header[j], new DB_Class.DB_Editing.Data_Calculation(Data_Interface.Clotho_List[0].Max.Length));
            }
            double Testtime5 = TestTime1.Elapsed.TotalMilliseconds;

        }
        public void For_New_Spec_Cal_Yield2(int TotalCount)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool while_Flag = true;
            int[] Yield = new int[advanced.Length];
            int DB = 0;

            for (int n = 0; n < Db_Interface.For_Any_Yield_Percent_For_New_Spec[0].Count; n++)
            {
                DB = 0;

                while (while_Flag)
                {
                    for (int j = 0; j < advanced.Length; j++)
                    {
                        for (int m = 0; m < Db_Interface.For_Any_Yield_Percent_For_New_Spec[DB][n][j].Count; m++)
                        {
                            if (Db_Interface.For_Any_Yield_Percent_For_New_Spec[DB][n][j][m] != 0)
                            {
                                Yield[j]++;
                                while_Flag = false;
                                break;
                            }
                        }

                    }
                    DB++;
                    if (DB == Db_Interface.For_Any_Yield_Percent_For_New_Spec.Length)
                    {
                        while_Flag = true;
                        break;
                    }

                }
                while_Flag = true;
            }

            for (int j = 0; j < advanced.Length; j++)
            {
                datagrid2[j].Rows[0].Cells[1].Value = Total;
                datagrid2[j].Rows[1].Cells[1].Value = Total - Hidden_Total;

                long Pass = 0;
                if ((Total - Hidden_Total) == Yield[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (Total - Hidden_Total) - Yield[j];
                }


                datagrid2[j].Rows[2].Cells[1].Value = Pass;
                datagrid2[j].Rows[3].Cells[1].Value = Yield[j];

                double Dummy = (Convert.ToDouble((Total - Hidden_Total) - Yield[j]) / (Total - Hidden_Total)) * 100;
                datagrid2[j].Rows[4].Cells[1].Value = Dummy;
                datagrid2[j].Rows[5].Cells[1].Value = Hidden_Total;

            }

            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void For_New_Spec_Cal_Yield3(int Sample)
        {


            for (int u = 0; u < Data_Interface.DB_Count; u++)
            {
                List_Sample_Verify[u] = new List<int[]>();
            }

            For_Cal = new ManualResetEvent[Data_Interface.DB_Count];
            Wait = new bool[Data_Interface.DB_Count];

            Stopwatch TestTime3 = new Stopwatch();
            TestTime3.Restart();
            TestTime3.Start();


            for (int u = 0; u < Data_Interface.DB_Count; u++)
            {
                For_Cal[u] = new ManualResetEvent(false);
                ThreadPool.QueueUserWorkItem(new WaitCallback(Cal_Thread), u);
            }

            for (int u = 0; u < Data_Interface.DB_Count; u++)
            {
                Wait[u] = For_Cal[u].WaitOne();
            }

            double Testtime2 = TestTime3.Elapsed.TotalMilliseconds;

            // for (int j = 0; j < advanced.Length; j++)
            // {
            //int Fail = 0;

            //for (int u = 0; u < Data_Interface.DB_Count; u++)
            //{
            //    Fail += List_Sample_Verify[u][0][j];
            //}

            //datagrid2[j].Rows[0].Cells[1].Value = Sample;
            //datagrid2[j].Rows[1].Cells[1].Value = Sample - Hidden_Total;

            //int Pass = 0;
            //if ((Sample - Hidden_Total) == Fail)
            //{
            //    Pass = 0;
            //}
            //else
            //{
            //    Pass = (Sample - Hidden_Total) - Fail;
            //}

            //datagrid2[j].Rows[2].Cells[1].Value = Pass;
            //datagrid2[j].Rows[3].Cells[1].Value = Fail;

            //double Dummy = (Convert.ToDouble((Sample - Hidden_Total) - Fail) / (Sample - Hidden_Total)) * 100;
            //datagrid2[j].Rows[4].Cells[1].Value = Dummy;
            //datagrid2[j].Rows[5].Cells[1].Value = Hidden_Total;

            //  }
        }
        private void Cal_Thread(Object index)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            int[] Bin_Arry = new int[Db_Interface.Yield_Test_New_Spec[0][0].Length];

            int i = (int)index;



            //for (long n = Calculate_thread_Strat[i]; n < Calculate_thread_End[i]; n++)
            //{
            //    for (int j = 0; j < Db_Interface.Yield_Test_New_Spec[Db][n].Length; j++)
            //    {
            //        for (Db = 0; Db < Data_Interface.DB_Count; Db++)
            //        {
            //            for (int m = 0; m < Db_Interface.Yield_Test_New_Spec[Db][n][j].Count; m++)
            //            {
            //                Bin_Arry[j]++;
            //                Db = 0;
            //                flag = true;
            //                break;

            //            }
            //            if (flag) break;
            //        }
            //        flag = false;

            //    }
            //}
            List_Sample_Verify[i].Add(Bin_Arry);
            For_Cal[i].Set();
            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }
        private void Cal_No_Thread(long Total)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            bool flag = false;
            int Db = 0;

            for (int n = 0; n < Db_Interface.Yield_Test_New_Spec[0].Count; n++)
            {
                for (int j = 0; j < Db_Interface.Yield_Test_New_Spec[Db][n].Length; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {
                        for (int m = 0; m < Db_Interface.Yield_Test_New_Spec[Db][n][j].Count; m++)
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


            for (int j = 0; j < advanced.Length; j++)
            {

                datagrid2[j].Rows[0].Cells[1].Value = this.Total;
                datagrid2[j].Rows[1].Cells[1].Value = Any_Total;

                long Pass = 0;
                if ((Any_Total) == Sample_Verify[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (Any_Total) - Sample_Verify[j];
                }

                datagrid2[j].Rows[2].Cells[1].Value = Pass;
                datagrid2[j].Rows[3].Cells[1].Value = Sample_Verify[j];

                double Dummy = (Convert.ToDouble((Any_Total) - Sample_Verify[j]) / (Any_Total)) * 100;
                datagrid2[j].Rows[4].Cells[1].Value = Dummy;
                datagrid2[j].Rows[5].Cells[1].Value = Hidden_Total;

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

            Sample_Verify = new int[Bin_Length];

            for (int n = 0; n < Db_Interface.Yield_Test_New_Spec[0].Count; n++)
            {
                for (int j = 0; j < Db_Interface.Yield_Test_New_Spec[0][0].Length; j++)
                {
                    for (Db = 0; Db < Data_Interface.DB_Count; Db++)
                    {
                        for (int m = 0; m < Db_Interface.Yield_Test_New_Spec[Db][n][j].Count; m++)
                        {

                            if (!Fail_Units.Contains(Convert.ToString(Db_Interface.Yield_Test_New_Spec[Db][n][j][m].SN)))
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

            for (int j = 0; j < datagrid2.Length; j++)
            {
                datagrid2[j].Rows[0].Cells[1].Value = this.Total;
                datagrid2[j].Rows[1].Cells[1].Value = this.Total - Hidden_Sample_Count;

                long Pass = 0;
                if ((this.Total - Hidden_Sample_Count) == Sample_Verify[j])
                {
                    Pass = 0;
                }
                else
                {
                    Pass = (this.Total - Hidden_Sample_Count) - Sample_Verify[j];
                }

                datagrid2[j].Rows[2].Cells[1].Value = Pass;
                datagrid2[j].Rows[3].Cells[1].Value = Sample_Verify[j];

                double Dummy = (Convert.ToDouble((this.Total - Hidden_Sample_Count) - Sample_Verify[j]) / (this.Total - Hidden_Sample_Count)) * 100;
                datagrid2[j].Rows[4].Cells[1].Value = Dummy;
                datagrid2[j].Rows[5].Cells[1].Value = Hidden_Sample_Count;
                datagrid2[j].Rows[6].Cells[1].Value = Outlier_List.Count;

            }



            double Testtime1 = TestTime1.Elapsed.TotalMilliseconds;
        }
        public void STD(double[] Data, double[] Spec, out double Min, out double Max, out double Ave, out double L_CPK, out double H_CPK, out double Median, out double Stdev, int Spec_Index)
        {
            Min = Data.Min();
            Max = Data.Max();
            Ave = Data.Average();
            L_CPK = 0f;
            H_CPK = 0f;
            Median = 0f;
            Stdev = 0f;

            double[] Data_dummy = new double[Data.Length - Db_Interface.DIC_IQR[Data_Interface.Reference_Header[Spec_Index + 1]].SN.Count()];
            int dummy_Count = 0;

            for (int w = 0; w < Data.Length; w++)
            {
                if (Db_Interface.DIC_IQR[Data_Interface.Reference_Header[Spec_Index + 1]].SN.Length == 0)
                {
                    Data_dummy[w] = Data[w];
                    dummy_Count++;
                }
                else
                {

                    if (!Db_Interface.DIC_IQR[Data_Interface.Reference_Header[Spec_Index + 1]].SN.Contains(Convert.ToString(w + 1)))
                    {
                        Data_dummy[dummy_Count] = Data[w];
                        dummy_Count++;
                    }

                }

            }

            if (Data_dummy.Length % 2 == 0)
            {
                Array.Sort(Data_dummy);
                double GetMedian_i = Data_dummy[(Data_dummy.Length / 2) - 1];
                double GetMedian_j = Data_dummy[(Data_dummy.Length / 2)];

                Median = (GetMedian_i + GetMedian_j) / 2;
            }
            else
            {
                Array.Sort(Data_dummy);
                int GetMedian_i = (Data_dummy.Length) / 2;
                Median = Data_dummy[GetMedian_i];
            }

            double minusSquareSummary = 0.0;

            foreach (double source in Data_dummy)
            {
                minusSquareSummary += (source - Ave) * (source - Ave);
            }

            Stdev = Math.Sqrt(minusSquareSummary / (Data_dummy.Length - 1));

            L_CPK = (Ave - Spec[0]) / (3 * Stdev);
            H_CPK = (Spec[1] - Ave) / (3 * Stdev);

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = tabControl1.SelectedIndex;
            tabControl3.SelectedIndex = index;

        }
        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = tabControl2.SelectedIndex;

        }
        private void tabControl3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = tabControl3.SelectedIndex;
            tabControl1.SelectedIndex = index;

        }
        public void ForeColor()
        {


            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                tabControl1.SelectedIndex = s;
                //  advanced[s].Visible = false;


                string[] No_Index = new string[advanced[s].RowCount];

                int advancedDataGridView1_RowCount = advanced[s].RowCount;
                for (int k = 0; k < advancedDataGridView1_RowCount; k++)
                {
                    No_Index[k] = advanced[s].Rows[k].Cells[1].Value.ToString();

                    string[] ParanameSplit = No_Index[k].Split('_');


                    if (ParanameSplit[ParanameSplit.Length - 1].Contains('-'))
                    {
                        //DataGridViewRow rowStyle = advanced[s].Rows[k];

                        //rowStyle.DefaultCellStyle.BackColor = Color.Red;
                        advanced[s].Rows[k].DefaultCellStyle.BackColor = Color.Red;

                        //   advancedDataGridView1.Rows[k].DefaultCellStyle.BackColor = Color.Red;

                    }
                }
                //   advanced[s].Visible = true;

            }
            tabControl1.SelectedIndex = 0;
        }
        public void ForeColor(int i)
        {


            string[] No_Index = new string[advanced[i].RowCount];

            int advancedDataGridView1_RowCount = advanced[i].RowCount;
            for (int k = 0; k < advancedDataGridView1_RowCount; k++)
            {
                No_Index[k] = advanced[i].Rows[k].Cells[1].Value.ToString();

                string[] ParanameSplit = No_Index[k].Split('_');


                if (ParanameSplit[ParanameSplit.Length - 1].Contains('-'))
                {
                    //DataGridViewRow rowStyle = advanced[s].Rows[k];

                    //rowStyle.DefaultCellStyle.BackColor = Color.Red;
                    advanced[i].Rows[k].DefaultCellStyle.BackColor = Color.Red;

                    //   advancedDataGridView1.Rows[k].DefaultCellStyle.BackColor = Color.Red;

                }
            }


            tabControl1.SelectedIndex = 0;
        }

        private void MakeSpec_Form_FormClosing(object sender, FormClosingEventArgs e)
        {

            e.Cancel = true;
            this.Hide();
            ///this.Show();
        }

        private void button2_Click(object sender, EventArgs e) // save spec
        {

            for (int s = 0; s < Data_Interface.Clotho_List[0].Max.Length; s++)
            {
                advanced[s].CleanFilter();
                advanced[s].CleanSort();

                int index = tabControl1.SelectedIndex;
                _dataTable[s].DefaultView.Sort = "[No] ASC";
                bindingSource[s].Filter = "";

                advanced[s].Update();
            }


            for (int i = 0; i < Data_Interface.Customor_Clotho_List[0].Max.Length; i++)
            {
                Db_Interface.DropTable(Data_Interface, "drop table Table" + i);
            }


            for (int i = 0; i < Data_Interface.Customor_Clotho_List[0].Max.Length; i++)
            {
                Data_Interface.Data_Table = "Table" + i;
                Db_Interface.Insert_Spec_Header(Data_Interface);
            }

            for (int g = 0; g < advanced.Length; g++)
            {
                Db_Interface.trans(Data_Interface);
                for (int h = 0; h < Data_Interface.Reference_Header.Length - 1; h++)
                {
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].No[g] = Convert.ToInt16(advanced[g].Rows[h].Cells[0].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Parameter[g] = Convert.ToString(advanced[g].Rows[h].Cells[1].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Min_Selector[g] = Convert.ToString(advanced[g].Rows[h].Cells[2].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Max_Selector[g] = Convert.ToString(advanced[g].Rows[h].Cells[3].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Min_Spec_Control[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[4].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Max_Spec_Control[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[5].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Min_Spec[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[6].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Max_Spec[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[7].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Min_Data[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[8].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Median_Data[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[9].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Max_Data[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[10].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].CPK[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[11].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Std[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[12].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Persent[g] = Convert.ToDouble(advanced[g].Rows[h].Cells[13].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Fail_Count[g] = Convert.ToInt64(advanced[g].Rows[h].Cells[14].Value.ToString());
                    Db_Interface.For_New_Spec_Cal_Value_by_rowsdata[Convert.ToString(Data_Interface.Ref_New_Header[h + 1])].Outlier[g] = Convert.ToInt64(advanced[g].Rows[h].Cells[15].Value.ToString());


                }


                Data_Interface.Data_Table = "Table" + g;
                Db_Interface.Save_table(Data_Interface);

                Db_Interface.Commit(Data_Interface);
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
    }

    public static class ExtensionMethod2
    {
        public static void DoubleBuffered2(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.SetProperty);
            pi.SetValue(dgv, setting, null);
        }
    }
}
