using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ATE
{
    public partial class Box_Plot_For_Yield_Form : Form
    {

        DataTable _dataTable;
        DataSet _dataSet = new DataSet();
        BindingSource bindingSource;

        public Box_Plot_For_Yield_Form()
        {
            InitializeComponent();
            Gridview();

        }
        public void Gridview()
        {


            dataGridView1 = new Zuby.ADGV.AdvancedDataGridView();
            _dataTable = new DataTable();
            bindingSource = new BindingSource();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            dataGridView1.VirtualMode = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dataGridView1.Anchor = (AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom);
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;


            dataGridView1.Location = new System.Drawing.Point(10, 10);
            dataGridView1.Name = "advancedDataGridView1";
            dataGridView1.RowHeadersVisible = false;

            dataGridView1.Size = new System.Drawing.Size(2854, 1650);
            dataGridView1.TabIndex = 19;

            dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;

            bindingSource.DataSource = _dataTable;
            dataGridView1.DataSource = bindingSource;



            DataColumn[] dtkey = new DataColumn[1];

            _dataTable.Columns.Add("No", typeof(int));
            dtkey[0] = _dataTable.Columns["No"];
            _dataTable.PrimaryKey = dtkey;

            _dataTable.Columns.Add("Identifier");
            _dataTable.Columns.Add("Parameter", typeof(string));
            _dataTable.Columns.Add("Measure", typeof(string));
            _dataTable.Columns.Add("Band", typeof(string));
            _dataTable.Columns.Add("PowerMode", typeof(string));
            _dataTable.Columns.Add("Modulation", typeof(string));
            _dataTable.Columns.Add("Waveform", typeof(string));
            _dataTable.Columns.Add("Power_Identifier", typeof(string));

            _dataTable.Columns.Add("Pout", typeof(string));
            _dataTable.Columns.Add("Freq", typeof(string));
            _dataTable.Columns.Add("Vcc", typeof(string));
            _dataTable.Columns.Add("Vdd", typeof(string));
            _dataTable.Columns.Add("DAC01", typeof(string));
            _dataTable.Columns.Add("DAC02", typeof(string));
            _dataTable.Columns.Add("Input", typeof(string));
            _dataTable.Columns.Add("Ant", typeof(string));
            _dataTable.Columns.Add("Out", typeof(string));
            _dataTable.Columns.Add("Extra", typeof(string));
            _dataTable.Columns.Add("Note1", typeof(string));
            _dataTable.Columns.Add("Note2", typeof(string));
            bindingSource.DataMember = _dataTable.TableName;


            object Valuse = new object[11];


            _dataTable.Rows.Add(Valuse);


        }
    }
}
