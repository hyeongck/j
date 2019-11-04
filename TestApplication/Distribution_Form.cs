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
using System.IO;

namespace TestApplication
{
    public delegate void Delete(object sender, ToolStripItemClickedEventArgs e);

    public partial class Distribution_Form : Form
    {

        DataTable _dataTable;
        BindingSource bindingSource;

        DataTable _dataTable_Fit;
        BindingSource bindingSource_Fit;

        List<string> Item_Y_Dist;
        List<string> Item_By_Dist;

        List<string> Item_X_Fit;
        List<string> Item_Y_Fit;
        List<string> Item_By_Fit;



        DataGridViewCell clickedCell;
        object[] Valuse;
        int i;

        bool Delete_Flag;

        bool Customer_enable;
        bool NPI_enable;
        bool CPK_enable;
        double CPK_Value;

        string Key = "";


        Dictionary<int, Dictionary<int, string>> OrderbySequence = new Dictionary<int, Dictionary<int, string>>();
        Dictionary<int, Dictionary<int, string>> Box_Enum = new Dictionary<int, Dictionary<int, string>>();
        Dictionary<string, CSV_Class.For_Box>[] Dic_Test;
        Dictionary<string, CSV_Class.For_Box> Dic_X_Test = new Dictionary<string, CSV_Class.For_Box>();
        Dictionary<string, CSV_Class.For_Box> Dic_By_Test = new Dictionary<string, CSV_Class.For_Box>();
        string[] Header;
        string[] New_Header;

        DB_Class.DB_Editing DB;
        DB_Class.DB_Editing.INT DB_Interface;

        Data_Class.Data_Editing.INT Data_Interface;

        CSV_Class.CSV CSV = new CSV_Class.CSV();
        CSV_Class.CSV.INT CSV_Interface;


        JMP_Class.JMP_Editing.INT JMP_Interface;


        public  static event Delete Send;


        public Distribution_Form()
        {
            InitializeComponent();

            Data_Grid();
        }

        public Distribution_Form(Data_Class.Data_Editing.INT Data_Interface, DB_Class.DB_Editing.INT DB_Interface, JMP_Class.JMP_Editing.INT JMP_Interface, Dictionary<int, Dictionary<int, string>> Box_Enum, string[] Header, string[] New_Header, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, ref bool Delete_Flag, string Key)
        {
            this.Header = Header;
            this.New_Header = New_Header;
          
            this.Item_Y_Dist = new List<string>();
            this.Item_By_Dist = new List<string>();

            this.Item_X_Fit = new List<string>();
            this.Item_Y_Fit = new List<string>();
            this.Item_By_Fit = new List<string>();

            this.Customer_enable = Customer_enable;
            this.NPI_enable = NPI_enable;
            this.CPK_enable = CPK_enable;
            this.CPK_Value = CPK_Value;
            this.Data_Interface = Data_Interface;
            this.DB_Interface = DB_Interface;
            this.Box_Enum = Box_Enum;
            this.JMP_Interface = JMP_Interface;
            this.Key = Key;
            InitializeComponent();

            tabControl1.TabPages[0].Text = "Distribution";
            tabControl1.TabPages[1].Text = "Fit Y By X";

            listBox1.SelectionMode = SelectionMode.MultiExtended;
            listBox2.SelectionMode = SelectionMode.MultiExtended;
            listBox3.SelectionMode = SelectionMode.MultiExtended;
            listBox4.SelectionMode = SelectionMode.MultiExtended;
            listBox5.SelectionMode = SelectionMode.MultiExtended;


            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);

            Data_Grid();
            Data_Grid2();

            Data_Grid_Fit();
            Data_Grid2_Fit();


        }

        public void Data_Grid()
        {

            _dataTable = new DataTable();
            bindingSource = new BindingSource();

            dataGridView1.Anchor = (AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom);
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.VirtualMode = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            //     dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            //     dataGridView1.Location = new System.Drawing.Point(10, 10);
            dataGridView1.Name = "advancedDataGridView1";
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.RowTemplate.Height = 40;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            dataGridView1.TabIndex = 19;


            dataGridView1.RowHeadersVisible = false;
            //   dataGridView1.ColumnCount = 1;
            dataGridView1.BackgroundColor = Color.White;


            bindingSource.DataSource = _dataTable;
            dataGridView1.DataSource = bindingSource;

            dataGridView1.DoubleBuffereds(true);

            DataColumn[] dtkey = new DataColumn[1];

            Valuse = new object[2];

            _dataTable.Columns.Add("No", typeof(int));
            dtkey[0] = _dataTable.Columns["No"];
            _dataTable.PrimaryKey = dtkey;

    
            _dataTable.Columns.Add("Parameter" , typeof(string));

            Valuse[0] = 1; Valuse[1] = "SBIN";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 2; Valuse[1] = "HBIN";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 3; Valuse[1] = "DIE_X";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 4; Valuse[1] = "DIE_Y";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 5; Valuse[1] = "SITE";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 6; Valuse[1] = "TIME";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 7; Valuse[1] = "TOTAL_TESTS";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 8; Valuse[1] = "LOT_ID";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = 9; Valuse[1] = "WAFER_ID";
            _dataTable.Rows.Add(Valuse);


            bindingSource.DataMember = _dataTable.TableName;

            for (i = 10; i < Header.Length + 10; i++)
            {
                Valuse[0] = i;
                Valuse[1] = Header[i -10];
                _dataTable.Rows.Add(Valuse);
            }

            Valuse[0] = i++; Valuse[1] = "PassFail";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "TimeStamp";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "IndexTime";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "PartSN";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "SWBinName";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "HWBinName";
            _dataTable.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "SubLot";
            _dataTable.Rows.Add(Valuse);


            this.Show();
            dataGridView1.Visible = true;

            dataGridView1.Columns[0].Width = 40;


            dataGridView1.Update();

          
        }  ///Dist

        public void Data_Grid2()
        {

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            dataGridView2.AllowUserToAddRows = false;
            dataGridView2.AllowUserToDeleteRows = false;
            //     dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            //     dataGridView2.Location = new System.Drawing.Point(10, 10);
            dataGridView2.Name = "advancedDataGridView2";
            dataGridView2.RowHeadersVisible = false;
            dataGridView2.RowTemplate.Height = 40;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //      dataGridView1.Size = new System.Drawing.Size(2854, 1650);
            dataGridView2.TabIndex = 19;
            //  dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;

            dataGridView2.RowHeadersVisible = false;
            dataGridView2.ColumnCount = 2;
            dataGridView2.BackgroundColor = Color.White;

            Valuse = new object[3];

            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();

            dataGridView2.Columns.Add(buttonColumn);

            buttonColumn.HeaderText = "Check";

            List<string> Split = new List<string>();
            for (i = 0; i < Header.Length; i++)
            {
                if (Header[i].Split('_').Length != 1)
                {
                    Split.Add(Header[i].Split('_')[1]);
                }

            }
            Split = Split.Distinct().ToList();
            for (i = 0; i < Split.Count; i++)
            {

                Valuse[0] = Split[i];
                Valuse[1] = "";
                dataGridView2.Rows.Add(Valuse);
            }



            dataGridView2.Columns[0].Name = "Parameter";
            dataGridView2.Columns[1].Name = "Option";

            dataGridView2.Columns[0].Width = 80;
            dataGridView2.Columns[1].Width = 80;

            dataGridView2.Columns[0].ReadOnly = true;
            dataGridView2.Columns[1].ReadOnly = false;





        } ///Dist

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (e.ColumnIndex == 2)
                {



                    // _dataTable.DefaultView.Sort = advancedDataGridView1.SortString;
                    //  _dataTable = _dataTable.DefaultView.ToTable();
                    _dataTable.PrimaryKey = new DataColumn[] { _dataTable.Columns["No"] };
                    bindingSource.DataSource = _dataTable;
                    bindingSource.Filter = "";

                    dataGridView1.DataSource = bindingSource;
                    dataGridView1.Update();

                    StringBuilder ForFilter = new StringBuilder();


                    if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != null && dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != "")
                    {
                        string[] split = dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString().Split(',');

                        for (int i = 0; i < split.Length; i++)
                        {
                            if (i == 0)
                            {
                                ForFilter.Append("([Parameter] LIKE '%_" + dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());
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
                        bindingSource.Filter = ForFilter.ToString();

                        dataGridView1.DataSource = bindingSource;
                        dataGridView1.Update();
                        ForFilter = new StringBuilder();
                    }
                    else
                    {


                        ForFilter.Append("([Parameter] LIKE '%" + dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());
                        ForFilter.Append("%'");


                        ForFilter.Append(")");

                        bindingSource.Filter = ForFilter.ToString(); ;

                        dataGridView1.DataSource = bindingSource;
                        dataGridView1.Update();
                        ForFilter = new StringBuilder();

                    }

                }
                else
                {

                    if (e.ColumnIndex == 1)
                    {

                        //  dataGridView1.CleanFilter();
                        //   dataGridView1.CleanSort();


                        // _dataTable.DefaultView.Sort = advancedDataGridView1.SortString;
                        //  _dataTable = _dataTable.DefaultView.ToTable();
                        _dataTable.PrimaryKey = new DataColumn[] { _dataTable.Columns["No"] };
                        bindingSource.DataSource = _dataTable;
                        bindingSource.Filter = "";

                        dataGridView1.DataSource = bindingSource;
                        dataGridView1.Update();

                        StringBuilder ForFilter = new StringBuilder();


                        if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != null)
                        {

                            ForFilter.Append("([Parameter] LIKE '%" + dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString());
                            ForFilter.Append("%'");


                            ForFilter.Append(")");

                            bindingSource.Filter = ForFilter.ToString(); ;

                            dataGridView1.DataSource = bindingSource;
                            dataGridView1.Update();
                            ForFilter = new StringBuilder();


                        }


                    }

                }
            }

        }   ///Dist

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex != -1)
            {
                if(!Item_Y_Dist.Contains(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()))
                {
                    Item_Y_Dist.Add(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
                }

                listBox1.Items.Clear();
                for (i = 0; i < Item_Y_Dist.Count; i ++)
                {
                    listBox1.Items.Add(Item_Y_Dist[i]);
                }
     
            }
        }    ///Dist

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.ColumnIndex == 1)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                clickedCell = (sender as DataGridView).Rows[e.RowIndex].Cells[e.ColumnIndex];
                dataGridView1.CurrentCell = clickedCell;
                var relativeMousePosition = dataGridView1.PointToClient(Cursor.Position);

                m.Items.Add("Reset Filter");
                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Distribution Setting");
                m.Items.Add(new ToolStripSeparator());

                m.Items.Add("Add Item - Y");
                m.Items.Add("Add Item - By");

                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Delete Units");
                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Close All Windows");
                m.ItemClicked += new ToolStripItemClickedEventHandler(m_ItemClicked);

                m.Show(dataGridView1, relativeMousePosition);

            }
        }   ///Dist

        public void m_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string CellValue = dataGridView1.Rows[clickedCell.RowIndex].Cells[0].Value.ToString();
            DataObject Do = dataGridView1.GetClipboardContent();

            Clipboard.SetDataObject(Do);
            string s = Clipboard.GetText();

            string[] lines = s.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            switch (e.ClickedItem.Text)
            {
                case "Close All Windows":

                    JMP_Interface.CloseWindowas();

                    break;

                case "Add Item - Y":

                    for(i = 0; i < lines.Length; i ++)
                    {
                        if (!Item_Y_Dist.Contains(lines[i]))
                        {
                            Item_Y_Dist.Add(lines[i]);
                        }
                    }
                    listBox1.Items.Clear();
                    for (i = 0; i < Item_Y_Dist.Count; i++)
                    {
                        listBox1.Items.Add(Item_Y_Dist[i]);
                    }

                    break;

                case "Add Item - By":

                    for (i = 0; i < lines.Length; i++)
                    {
                        if (!Item_By_Dist.Contains(lines[i]))
                        {
                            Item_By_Dist.Add(lines[i]);
                        }
                    }
                    listBox2.Items.Clear();
                    for (i = 0; i < Item_By_Dist.Count; i++)
                    {
                        listBox2.Items.Add(Item_By_Dist[i]);
                    }
                    break;
                case "Reset Filter":

                    _dataTable.DefaultView.RowFilter = "";
                    _dataTable.DefaultView.Sort = "[No] ASC";
                    bindingSource.DataSource = _dataTable;
                    bindingSource.Filter = "";

                    break;

                case "Delete Units":

                
                    Send(sender, e);

                  


                    break;
            }
        }   ///Dist

        public void Listbox1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

            switch (e.ClickedItem.Text)
            {

                case "Delete":

                    var dd = listBox1.SelectedItems.Cast<string>().ToList();
                    foreach (string i in dd)
                    {
                        Item_Y_Dist.Remove(i);
                        listBox1.Items.Remove(i);
                    }

                    break;

            }
        }  ///Dist

        public void Listbox2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            switch (e.ClickedItem.Text)
            {

                case "Delete":

                    var dd = listBox2.SelectedItems.Cast<string>().ToList();
                    foreach (string i in dd)
                    {
                        Item_By_Dist.Remove(i);
                        listBox2.Items.Remove(i);
                    }


                    break;

            }
        }  ///Dist

        private void button1_Click(object sender, EventArgs e)
        {
            CSV_Interface = CSV.Open("YIELD");




            DB_Interface.Line = Item_Y_Dist.ToArray();

            if (DB_Interface.Line.Length != 0)
            {
                DB_Interface.Get_Selected_Para(Data_Interface, _dataTable);


                DB_Interface.Dic_Test_For_Spec_Gen = new Dictionary<string, CSV_Class.For_Box>();


                foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                {
                    foreach (KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                    {
                        DB_Interface.Dic_Test_For_Spec_Gen.Add(test2.Key, test2.Value);
                    }

                }

            }

            DB_Interface.Line = Item_By_Dist.ToArray();

            string[] By = new string[DB_Interface.Line.Length];

            if (DB_Interface.Line.Length != 0)
            {
                int i = 0;


                DB_Interface.Get_Selected_Para(Data_Interface, _dataTable);


                Dic_By_Test = new Dictionary<string, CSV_Class.For_Box>();


                foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                {
                    foreach (KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                    {
                        Dic_By_Test.Add(test2.Key, test2.Value);
                        By[i] = ":" + test2.Key;
                    }

                }


            }


            int k = 0;



            Dictionary<string, CSV_Class.For_Box> Concot = new Dictionary<string, CSV_Class.For_Box>();


            Concot = DB_Interface.Dic_Test_For_Spec_Gen.Concat(Dic_By_Test).ToDictionary(x => x.Key, x => x.Value);

            if (Concot.Count != 0)
            {
                CSV_Interface.Write_Open("C:\\temp\\dummy\\Distributions.csv");

                CSV_Interface.Write(Concot);

                CSV_Interface.Write_Close();



                Ordersequence_Method();

                JMP_Draw("C:\\temp\\dummy\\Distributions.csv", DB_Interface.Dic_Test_For_Spec_Gen, null,  Dic_By_Test, OrderbySequence, "Seleted_Distributions", null, By);

                DB_Interface.Dic_Test = new Dictionary<string, CSV_Class.For_Box>[Data_Interface.DB_Count];
            }
        }  ///Dist


        private void listBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                m.Items.Add("Delete");
                m.ItemClicked += new ToolStripItemClickedEventHandler(Listbox1_ItemClicked);
                var index = Control.MousePosition.X;

                m.Show(Control.MousePosition.X, Control.MousePosition.Y);

            }
        }   ///Dist

        private void listBox2_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                m.Items.Add("Delete");
                m.ItemClicked += new ToolStripItemClickedEventHandler(Listbox2_ItemClicked);
                var index = Control.MousePosition.X;

                m.Show(Control.MousePosition.X, Control.MousePosition.Y);

            }
        }   ///Dist



        public void Data_Grid_Fit()
        {

            _dataTable_Fit = new DataTable();
            bindingSource_Fit = new BindingSource();

            dataGridView3.Anchor = (AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom);
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView3.VirtualMode = true;
            dataGridView3.AllowUserToAddRows = false;
            dataGridView3.AllowUserToDeleteRows = false;
            //     dataGridView3.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dataGridView3.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            //     dataGridView3.Location = new System.Drawing.Point(10, 10);
            dataGridView3.Name = "advanceddataGridView3";
            dataGridView3.RowHeadersVisible = false;
            dataGridView3.RowTemplate.Height = 40;
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            dataGridView3.TabIndex = 19;


            dataGridView3.RowHeadersVisible = false;
            //   dataGridView3.ColumnCount = 1;
            dataGridView3.BackgroundColor = Color.White;


            bindingSource_Fit.DataSource = _dataTable_Fit;
            dataGridView3.DataSource = bindingSource_Fit;

            dataGridView3.DoubleBuffereds(true);

            DataColumn[] dtkey = new DataColumn[1];

            Valuse = new object[2];

            _dataTable_Fit.Columns.Add("No", typeof(int));
            dtkey[0] = _dataTable_Fit.Columns["No"];
            _dataTable_Fit.PrimaryKey = dtkey;


            _dataTable_Fit.Columns.Add("Parameter", typeof(string));

            Valuse[0] = 1; Valuse[1] = "SBIN";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 2; Valuse[1] = "HBIN";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 3; Valuse[1] = "DIE_X";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 4; Valuse[1] = "DIE_Y";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 5; Valuse[1] = "SITE";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 6; Valuse[1] = "TIME";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 7; Valuse[1] = "TOTAL_TESTS";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 8; Valuse[1] = "LOT_ID";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = 9; Valuse[1] = "WAFER_ID";
            _dataTable_Fit.Rows.Add(Valuse);


            bindingSource_Fit.DataMember = _dataTable_Fit.TableName;

            for (i = 10; i < Header.Length + 10; i++)
            {
                Valuse[0] = i;
                Valuse[1] = Header[i - 10];
                _dataTable_Fit.Rows.Add(Valuse);
            }

            Valuse[0] = i++; Valuse[1] = "PassFail";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "TimeStamp";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "IndexTime";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "PartSN";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "SWBinName";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "HWBinName";
            _dataTable_Fit.Rows.Add(Valuse);
            Valuse[0] = i++; Valuse[1] = "SubLot";
            _dataTable_Fit.Rows.Add(Valuse);


            this.Show();
            dataGridView3.Visible = true;

          //  dataGridView3.Columns[0].Width = 40;


            dataGridView3.Update();


        }  ///Fit

        public void Data_Grid2_Fit()
        {

            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            dataGridView4.AllowUserToAddRows = false;
            dataGridView4.AllowUserToDeleteRows = false;
            //     dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dataGridView4.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            //     dataGridView4.Location = new System.Drawing.Point(10, 10);
            dataGridView4.Name = "advanceddataGridView4";
            dataGridView4.RowHeadersVisible = false;
            dataGridView4.RowTemplate.Height = 40;
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //      dataGridView1.Size = new System.Drawing.Size(2854, 1650);
            dataGridView4.TabIndex = 19;
            //  dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;

            dataGridView4.RowHeadersVisible = false;
            dataGridView4.ColumnCount = 2;
            dataGridView4.BackgroundColor = Color.White;

            Valuse = new object[3];

            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();

            dataGridView4.Columns.Add(buttonColumn);

            buttonColumn.HeaderText = "Check";

            List<string> Split = new List<string>();
            for (i = 0; i < Header.Length; i++)
            {
                if (Header[i].Split('_').Length != 1)
                {
                    Split.Add(Header[i].Split('_')[1]);
                }

            }
            Split = Split.Distinct().ToList();
            for (i = 0; i < Split.Count; i++)
            {

                Valuse[0] = Split[i];
                Valuse[1] = "";
                dataGridView4.Rows.Add(Valuse);
            }



            dataGridView4.Columns[0].Name = "Parameter";
            dataGridView4.Columns[1].Name = "Option";

            dataGridView4.Columns[0].Width = 80;
            dataGridView4.Columns[1].Width = 80;

            dataGridView4.Columns[0].ReadOnly = true;
            dataGridView4.Columns[1].ReadOnly = false;





        } ///Fit

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (e.ColumnIndex == 2)
                {



                    // _dataTable_Fitt.DefaultView.Sort = advanceddataGridView3.SortString;
                    //  _dataTable_Fitt = _dataTable_Fitt.DefaultView.ToTable();
                    _dataTable_Fit.PrimaryKey = new DataColumn[] { _dataTable_Fit.Columns["No"] };
                    bindingSource_Fit.DataSource = _dataTable_Fit;
                    bindingSource_Fit.Filter = "";

                    dataGridView3.DataSource = bindingSource_Fit;
                    dataGridView3.Update();

                    StringBuilder ForFilter = new StringBuilder();


                    if (dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != null && dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != "")
                    {
                        string[] split = dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString().Split(',');

                        for (int i = 0; i < split.Length; i++)
                        {
                            if (i == 0)
                            {
                                ForFilter.Append("([Parameter] LIKE '%_" + dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());
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
                        bindingSource_Fit.Filter = ForFilter.ToString();

                        dataGridView3.DataSource = bindingSource_Fit;
                        dataGridView3.Update();
                        ForFilter = new StringBuilder();
                    }
                    else
                    {


                        ForFilter.Append("([Parameter] LIKE '%" + dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex - 2].Value.ToString());
                        ForFilter.Append("%'");


                        ForFilter.Append(")");

                        bindingSource_Fit.Filter = ForFilter.ToString(); ;

                        dataGridView3.DataSource = bindingSource_Fit;
                        dataGridView3.Update();
                        ForFilter = new StringBuilder();

                    }

                }
                else
                {

                    if (e.ColumnIndex == 1)
                    {

                        //  dataGridView3.CleanFilter();
                        //   dataGridView3.CleanSort();


                        // _dataTable_Fitt.DefaultView.Sort = advanceddataGridView3.SortString;
                        //  _dataTable_Fitt = _dataTable_Fitt.DefaultView.ToTable();
                        _dataTable_Fit.PrimaryKey = new DataColumn[] { _dataTable_Fit.Columns["No"] };
                        bindingSource_Fit.DataSource = _dataTable_Fit;
                        bindingSource_Fit.Filter = "";

                        dataGridView3.DataSource = bindingSource_Fit;
                        dataGridView3.Update();

                        StringBuilder ForFilter = new StringBuilder();


                        if (dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value != null)
                        {

                            ForFilter.Append("([Parameter] LIKE '%" + dataGridView4.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString());
                            ForFilter.Append("%'");


                            ForFilter.Append(")");

                            bindingSource_Fit.Filter = ForFilter.ToString(); ;

                            dataGridView3.DataSource = bindingSource_Fit;
                            dataGridView3.Update();
                            ForFilter = new StringBuilder();


                        }


                    }

                }
            }

        }   ///Fit

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                if (!Item_Y_Fit.Contains(dataGridView3.Rows[e.RowIndex].Cells[1].Value.ToString()))
                {
                    Item_Y_Fit.Add(dataGridView3.Rows[e.RowIndex].Cells[1].Value.ToString());
                }

                listBox3.Items.Clear();
                for (i = 0; i < Item_Y_Fit.Count; i++)
                {
                    listBox3.Items.Add(Item_Y_Fit[i]);
                }

            }
        }    ///Fit

        private void dataGridView3_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.ColumnIndex == 1)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                clickedCell = (sender as DataGridView).Rows[e.RowIndex].Cells[e.ColumnIndex];
                dataGridView3.CurrentCell = clickedCell;
                var relativeMousePosition = dataGridView3.PointToClient(Cursor.Position);

                m.Items.Add("Reset Filter");
                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Distribution Setting");
                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Add Item - Y");
                m.Items.Add("Add Item - X");
                m.Items.Add("Add Item - By");

                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Delete Units");
                m.Items.Add(new ToolStripSeparator());
                m.Items.Add("Close All Windows");
                m.ItemClicked += new ToolStripItemClickedEventHandler(m_ItemClicked_FIt);

                m.Show(dataGridView3, relativeMousePosition);

            }
        }   ///Fit

        public void m_ItemClicked_FIt(object sender, ToolStripItemClickedEventArgs e)
        {
            string CellValue = dataGridView3.Rows[clickedCell.RowIndex].Cells[0].Value.ToString();
            DataObject Do = dataGridView3.GetClipboardContent();

            Clipboard.SetDataObject(Do);
            string s = Clipboard.GetText();

            string[] lines = s.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            switch (e.ClickedItem.Text)
            {
                case "Close All Windows":

                    JMP_Interface.CloseWindowas();

                    break;

                case "Add Item - X":

                    for (i = 0; i < lines.Length; i++)
                    {
                        if (!Item_X_Fit.Contains(lines[i]))
                        {
                            Item_X_Fit.Add(lines[i]);
                        }
                    }
                    listBox5.Items.Clear();
                    for (i = 0; i < Item_X_Fit.Count; i++)
                    {
                        listBox5.Items.Add(Item_X_Fit[i]);
                    }

                    break;
                case "Add Item - Y":

                    for (i = 0; i < lines.Length; i++)
                    {
                        if (!Item_Y_Fit.Contains(lines[i]))
                        {
                            Item_Y_Fit.Add(lines[i]);
                        }
                    }
                    listBox4.Items.Clear();
                    for (i = 0; i < Item_Y_Fit.Count; i++)
                    {
                        listBox4.Items.Add(Item_Y_Fit[i]);
                    }

                    break;

                case "Add Item - By":

                    for (i = 0; i < lines.Length; i++)
                    {
                        if (!Item_By_Fit.Contains(lines[i]))
                        {
                            Item_By_Fit.Add(lines[i]);
                        }
                    }
                    listBox3.Items.Clear();
                    for (i = 0; i < Item_By_Fit.Count; i++)
                    {
                        listBox3.Items.Add(Item_By_Fit[i]);
                    }
                    break;
                case "Reset Filter":

                    _dataTable_Fit.DefaultView.RowFilter = "";
                    _dataTable_Fit.DefaultView.Sort = "[No] ASC";
                    bindingSource.DataSource = _dataTable_Fit;
                    bindingSource.Filter = "";

                    break;

                case "Delete Units":


                    Send(sender, e);




                    break;
            }
        }   ///Fit

        public void Listbox3_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

            switch (e.ClickedItem.Text)
            {

                case "Delete":

                    var dd = listBox3.SelectedItems.Cast<string>().ToList();
                    foreach (string i in dd)
                    {
                        Item_Y_Fit.Remove(i);
                        listBox3.Items.Remove(i);
                    }

                    break;

            }
        }  ///Fit

        public void Listbox4_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            switch (e.ClickedItem.Text)
            {

                case "Delete":

                    var dd = listBox4.SelectedItems.Cast<string>().ToList();
                    foreach (string i in dd)
                    {
                        Item_By_Fit.Remove(i);
                        listBox4.Items.Remove(i);
                    }


                    break;

            }
        }  ///Fit

        public void Listbox5_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            switch (e.ClickedItem.Text)
            {

                case "Delete":

                    var dd = listBox5.SelectedItems.Cast<string>().ToList();
                    foreach (string i in dd)
                    {
                        Item_X_Fit.Remove(i);
                        listBox5.Items.Remove(i);
                    }


                    break;

            }
        }  ///Fit

        private void button2_Click(object sender, EventArgs e)
        {
            CSV_Interface = CSV.Open("YIELD");




            DB_Interface.Line = Item_Y_Fit.ToArray();
            DB_Interface.Dic_Test_For_Spec_Gen = new Dictionary<string, CSV_Class.For_Box>();
            if (DB_Interface.Line.Length != 0)
            {
                DB_Interface.Get_Selected_Para(Data_Interface, _dataTable_Fit);


       


                foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                {
                    foreach (KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                    {
                        DB_Interface.Dic_Test_For_Spec_Gen.Add(test2.Key, test2.Value);
                    }

                }

            }

            DB_Interface.Line = Item_X_Fit.ToArray();

            Dic_X_Test = new Dictionary<string, CSV_Class.For_Box>();
            string[] X = new string[DB_Interface.Line.Length];

            if (DB_Interface.Line.Length != 0)
            {
                int i = 0;

                DB_Interface.Get_Selected_Para(Data_Interface, _dataTable_Fit);




                foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                {
                    foreach (KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                    {
                       Dic_X_Test.Add(test2.Key, test2.Value);
                       X[i] = ":" + test2.Key;
                        i++;
                    }

                }

            }



            DB_Interface.Line = Item_By_Fit.ToArray();

            string[] By = new string[DB_Interface.Line.Length];
            Dic_By_Test = new Dictionary<string, CSV_Class.For_Box>();
            if (DB_Interface.Line.Length != 0)
            {
                int i = 0;


                DB_Interface.Get_Selected_Para(Data_Interface, _dataTable_Fit);





                foreach (Dictionary<string, CSV_Class.For_Box> test in DB_Interface.Dic_Test)
                {
                    foreach (KeyValuePair<string, CSV_Class.For_Box> test2 in test)
                    {
                        Dic_By_Test.Add(test2.Key, test2.Value);
                        By[i] = ":" + test2.Key;
                        i++;
                    }

                }


            }


            int k = 0;

            if (DB_Interface.Dic_Test_For_Spec_Gen.Count != 0)
            {

                if (Dic_X_Test.Count != 0 || Dic_By_Test.Count != 0)
                {

                    Dictionary<string, CSV_Class.For_Box> Concot = new Dictionary<string, CSV_Class.For_Box>();

                    Concot = DB_Interface.Dic_Test_For_Spec_Gen.Concat(Dic_X_Test).ToDictionary(x => x.Key, x => x.Value);

                    Concot = Concot.Concat(Dic_By_Test).ToDictionary(x => x.Key, x => x.Value);

                    if (Concot.Count != 0)
                    {
                        CSV_Interface.Write_Open("C:\\temp\\dummy\\Fit_Y_By_X.csv");

                        CSV_Interface.Write(Concot);

                        CSV_Interface.Write_Close();



                        Ordersequence_Method();

                        JMP_Draw("C:\\temp\\dummy\\Fit_Y_By_X.csv", DB_Interface.Dic_Test_For_Spec_Gen, Dic_X_Test, Dic_By_Test, OrderbySequence, "Seleted_Fit_Y_By_X", X, By);

                        DB_Interface.Dic_Test = new Dictionary<string, CSV_Class.For_Box>[Data_Interface.DB_Count];

                    }
                }
                else
                {
                    MessageBox.Show("Please Check 'X' or 'By' Factor");
                }
            }

          
        }  ///Fit



        private void listBox3_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                m.Items.Add("Delete");
                m.ItemClicked += new ToolStripItemClickedEventHandler(Listbox3_ItemClicked);
                var index = Control.MousePosition.X;

                m.Show(Control.MousePosition.X, Control.MousePosition.Y);

            }
        }   ///Fit
            
        private void listBox4_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                m.Items.Add("Delete");
                m.ItemClicked += new ToolStripItemClickedEventHandler(Listbox4_ItemClicked);
                var index = Control.MousePosition.X;

                m.Show(Control.MousePosition.X, Control.MousePosition.Y);

            }
        }   ///Fit

        private void listBox5_MouseUp(object sender, MouseEventArgs e)   ///Fit
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip m = new ContextMenuStrip();
                m.Items.Add("Delete");
                m.ItemClicked += new ToolStripItemClickedEventHandler(Listbox5_ItemClicked);
                var index = Control.MousePosition.X;

                m.Show(Control.MousePosition.X, Control.MousePosition.Y);

            }
        }   ///Fit

        private void JMP_Draw(string FilePaht, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<string, CSV_Class.For_Box> By_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, string Key, string[] X, string[] By)
        {
            JMP_Interface.Open_Session(true);


            JMP_Interface.Open_Document(FilePaht);
            JMP_Interface.GetDataTable();

            JMP_Class.Script Distribution_Script;

            Distribution_Script = null;
            List<string>[] Para_Test = new List<string>[OrderbySequence.Count];
            Dictionary<int, Dictionary<int, string>> dummy = new Dictionary<int, Dictionary<int, string>>();
            
            switch (Key)
            {

                case "Seleted_Distributions":

                    Distribution_Script = JMP_Interface.Make_Script(Key, Data, X_Data, By_Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X,By);

                    break;

                case "Seleted_Fit_Y_By_X":

                    Distribution_Script = JMP_Interface.Make_Script(Key, Data, X_Data, By_Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X, By);

                    break;

            }



            string Script = Distribution_Script.Scrip_Data;
            string[] Split = Script.Split('#');

            for (i = 0; i < Split.Length; i++)
            {
                CSV_Interface.Write_Open("C:\\temp\\dummy\\dummy.jsl");
                CSV_Interface.WriteScript(Split[i]);
                CSV_Interface.Write_Close();

                JMP_Interface.Run_Script("C:\\temp\\dummy\\dummy.jsl");
            }



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
                foreach (KeyValuePair<int, Dictionary<int, string>> D in this.Box_Enum)
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

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = tabControl1.SelectedIndex;

            //if(index == 0)
            //{
            //    Data_Grid();
            //    Data_Grid2();
            //}
            //else if(index == 1)
            //{
            //    Data_Grid_Fit();
            //    Data_Grid2_Fit();
            //}
      
        }

  
    }
    public static class ExtensionMethods
    {
        public static void DoubleBuffereds(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.SetProperty);
            pi.SetValue(dgv, setting, null);
        }
    }

}
