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
using System.Windows.Forms.DataVisualization.Charting;

namespace ATE
{
    public partial class SPARA_Form : Form
    {
        DB_Class.DB_Editing.INT DB_Interface;

        ChartArea[] CA;
        RectangleAnnotation[] RA;
        VerticalLineAnnotation[] VA;
        Series series;
        DataTable[] dt;

        Point? prevPosition;
        ToolTip tooltip = new ToolTip();
        Annotation Ann;

        Dictionary<string, Dictionary<string, Ann_Class>> Ann_Dic;
        Dictionary<string, List<double>> Data = new Dictionary<string, List<double>>();
        int i = 0;
        int F = 1;
        int j;
        int dt_Count = 0;

        int Selected_X;
        int Selected_Y;

        int[] Trace;

        double PositionX;

        List<int> Units;
        List<string> Column;
        List<double> Freq;

        MouseEventArgs e_value;


        System.Windows.Forms.DataVisualization.Charting.Chart[] mYChart;
        TextBox[] _T;
        TextBox[] _T_1;
        TextBox[] _T_2;
        TextBox[] _T_3;

        Label[] _Label;
        Label[] _Label_1;

        ChartArea chartArea1;

        Legend[] legend1;
        Series[] series1;

        //  double Min_X = 10e6;
        //  double Max_X = 1000e6;

        double Min_X;
        double Max_X;

        Marker_Form M;
        Marker_Setting_Form Ms;

        ContextMenuStrip m;
        Ann_Class TestAn;

        public SPARA_Form(DB_Class.DB_Editing.INT DB_Interface, int[] Trace)
        {
            InitializeComponent();
            Marker_Setting_Form.Marker_Set_Send += new Marker_Set(Marker_Set_Send);

            checkBox1.Checked = true;

            this.DB_Interface = DB_Interface;
            this.Trace = Trace;


            for (i = 0; i < this.Trace.Length; i++)
            {
                listBox1.Items.Add("CHAN" + this.Trace[i]);
            }

            tabControl1.TabPages[0].Text = "Spara";
            tabControl1.TabPages[1].Text = "Phase";
            tabControl1.TabPages[2].Text = "GDEL";

            this.MouseWheel += new MouseEventHandler(BaseForm_MouseWheel);
            testc();

            Ann_Dic = new Dictionary<string, Dictionary<string, Ann_Class>>();

            for (i = 0; i < this.Trace.Length; i++)
            {
                Ann_Dic.Add("CHAN" + this.Trace[i], new Dictionary<string, Ann_Class>());
            }

            dataGridView1.Visible = false;
            dataGridView1.Enabled = false;
        }

        public void testc()
        {

            this.Show();
        }

        bool Marker_Flag = false;
        private void chart1_AnnotationPositionChanging(object sender, AnnotationPositionChangingEventArgs e)
        {


            int pt1 = (int)e.NewLocationX;
            //if (pt1 <= series.Points.Count - 1 && pt1 > 0)
            //{

            //   PositionX = chart1.ChartAreas[0].AxisX.PixelPositionToValue(Selected_X);

            Ann = SelectedAnnotation(e.NewLocationX, Selected_Y);

            if (Ann != null)
            {
                Marker_Flag = true;
                string R_Name = Ann.Name.Remove(0, 2);

                //    Ann_Dic[R_Name].RA.X = e.NewLocationX;

                //if (Ann != null)
                //{
                //    RA.X = Ann.X - RA.Width / 2;
                //}


                //double step = (series.Points[pt1 + 1].YValues[0] - series.Points[pt1].YValues[0]);
                //double deltaX = e.NewLocationX - series.Points[pt1].XValue;
                //double val = series.Points[pt1].YValues[0] + step * deltaX;

                //for(int i = 0; i < 3; i ++ )
                //{

                //    chart1.Titles[i].Text =  chart1.Series[i].Points[pt1].YValues[0].ToString();

                //}



                //     RA.Text = String.Format("{0:0.00}", val);
             //   chart1.Update();
            }
            else
            {
                Marker_Flag = false;
            }
            //   }
        }

        bool Delete_Marker_Flag = false;
        private void button1_Click(object sender, EventArgs e)
        {
            int index = tabControl2.TabCount;

            for(int total = 0; total < index; total ++)
            {
                Annotation[] A = mYChart[total].Annotations.ToArray();

                for (i = 0; i < A.Length; i++)
                {
                    mYChart[total].Annotations.Remove(A[i]);
                }


            }


            for (i = 0; i < 20; i ++)
            {
                Delete_Marker_Flag = true;
                dataGridView1.Rows[i].Cells[1].Value = "";
            }
            Delete_Marker_Flag = false;

            if (Marker_Form != null)
            {

                for (i = 0; i < tabControl2.TabPages.Count; i++)
                {
                    Marker_Form.listView1[i].Clear();
                }

                Marker_Form.Close();
            }
        
                Marker_Form = null;
                Snp_Data = new Dictionary<string, Dictionary<string, Dictionary<string, double[]>>>();
                _Marker_Data = new Dictionary<string, Dictionary<string, string>>();


                Dictionary<string, string> T = new Dictionary<string, string>();
                string Chan = listBox1.SelectedItem.ToString();

                dataGridView1.Rows.Clear();

                for (i = 0; i < 20; i++)
                {
                    object[] Valuse = new object[2];

                    Valuse[0] = "Marker" + (i + 1);
                    Valuse[1] = "";


                    dataGridView1.Rows.Add(Valuse);
                    if (!T.ContainsKey(Valuse[0].ToString()))
                        T.Add(Valuse[0].ToString(), Valuse[1].ToString());

                }
                _Marker_Data.Add(Chan, T);

                int width = TextRenderer.MeasureText(dataGridView1.Rows[0].Cells[0].Value.ToString(), dataGridView1.DefaultCellStyle.Font).Width;
                int heigh = TextRenderer.MeasureText(dataGridView1.Rows[0].Cells[0].Value.ToString(), dataGridView1.DefaultCellStyle.Font).Height;

                for (i = 0; i < 20; i++)
                {
                    dataGridView1.Rows[i].Height = Convert.ToInt16(heigh * 1.8);
                }


                dataGridView1.Columns[0].Width = Convert.ToInt16(width * 1.2);
                dataGridView1.Columns[1].Width = Convert.ToInt16(dataGridView1.Size.Width - dataGridView1.Columns[0].Width);

                dataGridView1.Visible = true;
                dataGridView1.Enabled = true;
            
        }

        private void BaseForm_MouseWheel(object sender, MouseEventArgs e)
        {
            try
            {
                int index = tabControl2.SelectedIndex;

                if (mYChart[index].ChartAreas[0].AxisX.ScaleView.ViewMaximum > 1000000)
                    return;

                if (e.Delta > 0)
                {
                    double xMin = mYChart[index].ChartAreas[0].AxisX.ScaleView.ViewMinimum;
                    double xMax = mYChart[index].ChartAreas[0].AxisX.ScaleView.ViewMaximum;
                    double yMin = mYChart[index].ChartAreas[0].AxisY.ScaleView.ViewMinimum;
                    double yMax = mYChart[index].ChartAreas[0].AxisY.ScaleView.ViewMaximum;

                    double posXStart = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) / 2;
                    double posXFinish = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) / 4;
                    double posYStart = mYChart[index].ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / 2;
                    double posYFinish = mYChart[index].ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / 4;

                    mYChart[index].ChartAreas[0].AxisX.ScaleView.Zoom(posXStart, posXFinish);

                    mYChart[index].ChartAreas[0].AxisX.Minimum = posXStart;
                    mYChart[index].ChartAreas[0].AxisX.Maximum = posXFinish;


                    mYChart[index].ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);

                    mYChart[index].ChartAreas[0].AxisY.Minimum = posYStart;
                    mYChart[index].ChartAreas[0].AxisY.Maximum = posYFinish;


                    Min_X = posXStart;
                    Max_X = posXFinish;
                    //     chart1.ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);
                }
                else if (e.Delta < 0)
                {
                    double xMin = mYChart[index].ChartAreas[0].AxisX.ScaleView.ViewMinimum;
                    double xMax = mYChart[index].ChartAreas[0].AxisX.ScaleView.ViewMaximum;
                    double yMin = mYChart[index].ChartAreas[0].AxisY.ScaleView.ViewMinimum;
                    double yMax = mYChart[index].ChartAreas[0].AxisY.ScaleView.ViewMaximum;

                    double posXStart = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) * 2;
                    double posXFinish = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) * 2;

                    posXStart = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) * 1.05;
                    posXFinish = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) * 1.05;

                    double posYStart = mYChart[index].ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) * 2;
                    double posYFinish = mYChart[index].ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) * 2;

                    posYStart = mYChart[index].ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) * 0.95;
                    posYFinish = mYChart[index].ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) * 0.95;

                    mYChart[index].ChartAreas[0].AxisX.ScaleView.Zoom(posXStart, posXFinish);

                    mYChart[index].ChartAreas[0].AxisX.Minimum = posXStart;
                    mYChart[index].ChartAreas[0].AxisX.Maximum = posXFinish;

                    mYChart[index].ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);

                    mYChart[index].ChartAreas[0].AxisY.Minimum = posYStart;
                    mYChart[index].ChartAreas[0].AxisY.Maximum = posYFinish;



                    Min_X = posXStart;
                    Max_X = posXFinish;
                    //    chart1.ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);

                    //chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset();
                    //chart1.ChartAreas[0].AxisY.ScaleView.ZoomReset();
                }
            }
            catch
            {

            }
        }

        private void chart1_MouseUp(object sender, MouseEventArgs e)
        {
            int index = tabControl2.SelectedIndex;


            if (e.Button == MouseButtons.Right)
            {
              //  mYChart[index].Update();
             //   mYChart[index].Focus();

                Axis ax = mYChart[index].ChartAreas[0].AxisX;
                double xv = ax.PixelPositionToValue(e.Location.X);

                Axis ay = mYChart[index].ChartAreas[0].AxisY;
                double xy = ay.PixelPositionToValue(e.Location.Y);

                if (xv < ax.Maximum && xv > ax.Minimum && xy < ay.Maximum && xy > ay.Minimum)
                {

                    m = new ContextMenuStrip();


                    m.Items.Add("Add Marker");
                    m.Items.Add(new ToolStripSeparator());
                    m.Items.Add("Delete Marker");
                    m.Items.Add(new ToolStripSeparator());
                    m.Items.Add("Delete Unit");
                    m.Items.Add(new ToolStripSeparator());
                    mYChart[index].ContextMenuStrip = m;

                    Selected_X = e.X;
                    Selected_Y = e.Y;

                    e_value = e;

                    m.ItemClicked += new ToolStripItemClickedEventHandler(ItemClicked);

                 //   mYChart[index].Update();

                }
                else
                {

                    //m = new ContextMenuStrip();
                    //m.Refresh();
                    //m.Update();
                    //m.Invalidate();
                    //m.Items.Add("Delete Unit");
                    //m.Items.Add(new ToolStripSeparator());


                    //mYChart[index].ContextMenuStrip = m;

                    //Selected_X = e.X;
                    //Selected_Y = e.Y;

                    //e_value = e;

                    //m.ItemClicked += new ToolStripItemClickedEventHandler(ItemClicked);
                    //mYChart[index].Update();
                }




                //var sourceChart = sender as Chart;
                //HitTestResult result = sourceChart.HitTest(e.X, e.Y);
                //ChartArea chartAreas = sourceChart.ChartAreas[0];

                //if (result.ChartElementType == ChartElementType.DataPoint)
                //{
                //    chartAreas.CursorX.Position = chartAreas.AxisX.PixelPositionToValue(e.X);
                //    chartAreas.CursorY.Position = chartAreas.AxisY.PixelPositionToValue(e.Y);
                //}


                //   var pos = LocationInChart(e.X, e.Y);
                //   m.ItemClicked += new ToolStripItemClickedEventHandler(ItemClicked);
                //int X = Convert.ToInt16(this.chart1.ChartAreas[0].CursorX.Position);
                //int Y = Convert.ToInt16(this.chart1.ChartAreas[0].CursorY.Position);
                //    m.Show(e.X,e.Y);

            }

        }

        bool isMouseDown = false;
        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                int index = tabControl2.SelectedIndex;


                if (!e.Button.HasFlag(MouseButtons.Left) && isMouseDown == true)
                {

                }
                else if (e.Button.HasFlag(MouseButtons.Left) && !Marker_Flag)
                {
                    Axis ax = mYChart[index].ChartAreas[0].AxisX;
                    double xv = ax.PixelPositionToValue(e.Location.X);

                    Axis ay = mYChart[index].ChartAreas[0].AxisY;
                    double xy = ay.PixelPositionToValue(e.Location.Y);

   
                    if (xv < ax.Maximum && xv > ax.Minimum && xy < ay.Maximum && xy > ay.Minimum)
                    {

                  //      mYChart[index].Update();
                   //     mYChart[index].Focus();

                        double range = ax.Maximum - ax.Minimum;

                        ax.Minimum -= Math.Round((xv - XDown),2);
                        ax.Maximum = Math.Round((ax.Minimum + range),2);


                        range = ay.Maximum - ay.Minimum;

                        ay.Minimum -= Math.Round((xy - YDown),1);
                        ay.Maximum = Math.Round((ay.Minimum + range),1);


                        mYChart[index].ChartAreas[0].AxisX.ScaleView.ZoomReset();

                        mYChart[index].ChartAreas[0].AxisY.ScaleView.ZoomReset();

                    //    mYChart[index].Update();
                    }
       
          
                }
                else
                {
                    var pos = e.Location;

                    if (pos != prevPosition)
                    {
                        prevPosition = pos;
                        var results = chart1.HitTest(pos.X, pos.Y, false,
                                                     ChartElementType.DataPoint);
                        foreach (var result in results)
                        {
                            if (result.ChartElementType == ChartElementType.DataPoint)
                            {
                                e_value = e;
                                var xVal = result.ChartArea.AxisX.PixelPositionToValue(pos.X);
                                var yVal = result.ChartArea.AxisY.PixelPositionToValue(pos.Y);




                                tooltip.Show("X=" + xVal + ", Y=" + yVal,
                                             this.chart1, e.Location.X, e.Location.Y - 15);
                            }
                        }
                    }
                    else
                    {
                        //      tooltip.RemoveAll();
                    }
                }



                //HitTestResult result = chart1.HitTest(e.X, e.Y);
                //System.Drawing.Point p = new System.Drawing.Point(e.X, e.Y);

                //chart1.ChartAreas[0].CursorX.Interval = 0;
                //chart1.ChartAreas[0].CursorX.SetCursorPixelPosition(p, true);
                //chart1.ChartAreas[0].CursorY.SetCursorPixelPosition(p, true);

            }
            catch
            {

            }
        }

        double XDown = double.NaN;
        double YDown = double.NaN;
        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            bool isMouseDown = true;

            Axis ax = mYChart[index].ChartAreas[0].AxisX;

            XDown = ax.PixelPositionToValue(e.Location.X);

            Axis ay = mYChart[index].ChartAreas[0].AxisY;

            YDown = ay.PixelPositionToValue(e.Location.Y);

        }

        public void ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            switch (e.ClickedItem.Text)
            {
                case "Add Marker":

                    string[] ItemList = new string[listBox2.Items.Count];

                    for (int k = 0; k < listBox2.Items.Count; k++)
                    {
                        ItemList[k] = listBox2.Items[k].ToString();

                    }

                    Ms = new Marker_Setting_Form(ItemList, index, mYChart[index], listBox1.Text, tabControl2.TabPages[index].Text);

                    break;
                case "Delete Marker":

                    #region

                    PositionX = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(Selected_X);

                    Ann = SelectedAnnotation(PositionX, Selected_Y);

                    mYChart[index].Annotations.Remove(Ann);

                    #endregion

                    break;

                case "Delete Unit":

                    #region

                    var pos = e_value;

                    var results = mYChart[index].HitTest(pos.X, pos.Y, false,
                                                 ChartElementType.DataPoint);
                    foreach (var result in results)
                    {
                        if (result.ChartElementType == ChartElementType.DataPoint)
                        {

                            try
                            {
                                Series Se = mYChart[index].Series[result.Series.Name];

                                mYChart[index].Series.Remove(Se);
                            }
                            catch
                            {

                            }

                        }
                    }

                    #endregion


                    break;
            }
        }

        private Annotation SelectedAnnotation(double x, double y)
        {
            int index = tabControl2.SelectedIndex;

            foreach (Annotation an in mYChart[index].Annotations)
            {
                if (Convert.ToInt16(an.X) == Convert.ToInt16(x))
                {
                    return an;
                }
            }
            return null;
        }

        bool TabPage_Flag = false;
        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string[] C = DB_Interface.Get_Data_By_Query("Select info from Trace_info where Chan = '" + listBox1.Text + "'");
            Column = new List<string>();
            Column = C[0].Split(',').ToList();



            listBox2.Items.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();


            for (i = 0; i < Column.Count - 1; i++)
            {

                if (Column[i].ToUpper().Contains("GDEL"))
                {
                    if (Column[i].ToUpper().Contains("PHASE"))
                    {
                        listBox3.Items.Add(Column[i]);
                    }
                    else
                    {
                        listBox4.Items.Add(Column[i]);
                    }

                }
                else if (Column[i].ToUpper().Contains("PHASE"))
                {
                    listBox3.Items.Add(Column[i]);
                }
                else
                {
                    listBox2.Items.Add(Column[i]);
                }

            }
            TabPage_Flag = true;
            tabControl2.TabPages.Clear();
            TabPage_Flag = false;


            for (int k = 0; k < listBox2.Items.Count; k++)
            {
                TabPage myTabPage = new TabPage(listBox2.Items[k].ToString());
                tabControl2.TabPages.Add(myTabPage);

            }



            mYChart = new Chart[listBox2.Items.Count];
            _T = new TextBox[listBox2.Items.Count];
            _T_1 = new TextBox[listBox2.Items.Count];
            _Label = new Label[listBox2.Items.Count];

            _T_2 = new TextBox[listBox2.Items.Count];
            _T_3 = new TextBox[listBox2.Items.Count];
            _Label_1 = new Label[listBox2.Items.Count];

            legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend[listBox2.Items.Count];
            series1 = new System.Windows.Forms.DataVisualization.Charting.Series[listBox2.Items.Count];

            CA = new ChartArea[listBox2.Items.Count];
            RA = new RectangleAnnotation[listBox2.Items.Count];
            VA = new VerticalLineAnnotation[listBox2.Items.Count];



            string[] _Data;
            string Text = listBox1.SelectedItem.ToString();
            string Data = "";

            Units = new List<int>();

            List<object[]> _info = DB_Interface.Get_Data_By_Querys("Select DISTINCT id, lotid from " + Text);

            for (i = 0; i < _info.Count; i++)
            {
                int D = Convert.ToInt16(_info[i][0].ToString().Remove(0, 4));
                Units.Add(D);
            }

            Units.Sort();


            chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();


            //if (dt_Count == 0)
            //{
                _Data = DB_Interface.Get_Data_By_Query("Select Freq from " + Text + " where id = 'Unit1'");

                Freq = Array.ConvertAll<string, double>(_Data, Convert.ToDouble).ToList();
          //  }


            F = 1;
            dt = new DataTable[listBox2.Items.Count];
            for (int item = 0; item < listBox2.Items.Count; item++)
            {
                dt[item] = new DataTable();
                dt_Count = 0;

                this.Data = new Dictionary<string, List<double>>();

                for (int j = 0; j < Units.Count; j++)
                {
                    Stopwatch TestTime1 = new Stopwatch();
                    TestTime1.Restart();
                    TestTime1.Start();

                    _Data = DB_Interface.Get_Data_By_Query("Select " + listBox2.Items[item] + " from " + Text + " where id =  'Unit" + Units[j] + "'");

                    double test3 = TestTime1.Elapsed.TotalMilliseconds;

                    List<double> _DummyData = Array.ConvertAll<string, double>(_Data, Convert.ToDouble).ToList();

                    if (!this.Data.ContainsKey(listBox2.Items[item].ToString()))
                    {
                        // this.Data.Add(listBox2.Items[item].ToString() + "_" + Text + "_Unit" + Units[j], _DummyData);
                        string text = listBox2.Items[item].ToString().Remove(listBox2.Items[item].ToString().Length - 3, 3);
                        this.Data.Add(text + "_SN" + Units[j], _DummyData);

                        if (dt_Count == 0)
                        {
                            dt[item].Columns.Add("Freq", typeof(double));
                            foreach (KeyValuePair<string, List<double>> _D in this.Data)
                            {
                                dt[item].Columns.Add(_D.Key, typeof(double));
                            }
                            dt_Count++;
                        }

                        double test4 = TestTime1.Elapsed.TotalMilliseconds;

                        foreach (KeyValuePair<string, List<double>> _D in this.Data)
                        {
                            if (!dt[item].Columns.Contains(_D.Key))
                                dt[item].Columns.Add(_D.Key, typeof(double));
                        }


                        object[] value = new object[this.Data.Count + 1];

                        double test15 = TestTime1.Elapsed.TotalMilliseconds;

                        if (dt[item].Rows.Count == 0)
                        {

                            for (i = 0; i < Freq.Count; i++)
                            {
                                F = 1;
                                value[0] = Freq[i] / 1e6;

                                foreach (KeyValuePair<string, List<double>> _D in this.Data)
                                {
                                    value[F] = _D.Value[i];
                                    F++;
                                }
                                dt[item].Rows.Add(value);
                            }
                        }
                        else
                        {

                            for (i = 0; i < Freq.Count; i++)
                            {
                                int F = 1;
                                value[0] = Freq[i] / 1e6;

                                foreach (KeyValuePair<string, List<double>> _D in this.Data)
                                {
                                    value[F] = _D.Value[i];
                                    F++;
                                }
                                dt[item].Rows[i].ItemArray = value;
                            }
                        }
                        double test6 = TestTime1.Elapsed.TotalMilliseconds;

                    }
                }

                int k = 0;

                mYChart[item] = new Chart();
                _T[item] = new TextBox();
                _T_1[item] = new TextBox();
                _Label[item] = new Label();

                _T_2[item] = new TextBox();
                _T_3[item] = new TextBox();
                _Label_1[item] = new Label();

                for (int unit = 0; unit < 1; unit++)
                {
                    chartArea1 = new ChartArea();

                    chartArea1.Name = "ChartArea" + unit;
                    mYChart[item].ChartAreas.Add(chartArea1);
                }


                legend1[item] = new Legend();
                series1[item] = new Series();


                mYChart[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                              | System.Windows.Forms.AnchorStyles.Left)
                              | System.Windows.Forms.AnchorStyles.Right)));

                mYChart[item].BorderlineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dash;


                legend1[item].Name = "Legend1";
                mYChart[item].Legends.Add(legend1[item]);
                mYChart[item].Location = new System.Drawing.Point(10, 10);
                mYChart[item].Name = "chart1";
                series1[item].BackHatchStyle = System.Windows.Forms.DataVisualization.Charting.ChartHatchStyle.LargeCheckerBoard;
                series1[item].ChartArea = "ChartArea1";
                series1[item].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                series1[item].Legend = "Legend1";
                series1[item].Name = "Series1";

                mYChart[item].Series.Add(series1[item]);

              //  mYChart[item].ChartAreas[0].AxisX.MaximumAutoSize = true;
                mYChart[item].Size = new System.Drawing.Size(tabControl2.TabPages[item].Size.Width - 100, tabControl2.TabPages[item].Size.Height - 20);
             //   mYChart[item].TabStop = false;
                mYChart[item].AnnotationPositionChanging += new System.EventHandler<System.Windows.Forms.DataVisualization.Charting.AnnotationPositionChangingEventArgs>(chart1_AnnotationPositionChanging);
                mYChart[item].MouseDown += new System.Windows.Forms.MouseEventHandler(chart1_MouseDown);
                mYChart[item].MouseMove += new System.Windows.Forms.MouseEventHandler(chart1_MouseMove);
                mYChart[item].MouseUp += new System.Windows.Forms.MouseEventHandler(this.chart1_MouseUp);

                _Label[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                       | System.Windows.Forms.AnchorStyles.Right))));

                _Label[item].TabStop = false;
                _Label[item].Location = new System.Drawing.Point(tabControl2.TabPages[item].Size.Width - 80, 10);
                _Label[item].Text = "Freq Range";
                _Label[item].AutoSize = true;

                _T[item].Location = new System.Drawing.Point(tabControl2.TabPages[item].Size.Width - 80, 30);
                _T[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                              | System.Windows.Forms.AnchorStyles.Right))));

                _T[item].Size = new System.Drawing.Size(80, 40);
                _T[item].KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
                _T[item].PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.textBox1_KeyDown);

                _T[item].TabIndex = 10;


                _T_1[item].Location = new System.Drawing.Point(tabControl2.TabPages[item].Size.Width - 80, 55);
                _T_1[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                              | System.Windows.Forms.AnchorStyles.Right))));
                _T_1[item].Size = new System.Drawing.Size(80, 40);
                _T_1[item].KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox2_KeyPress);
                _T_1[item].PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.textBox2_KeyDown);
                _T_1[item].TabIndex = 11;

                _Label_1[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
              | System.Windows.Forms.AnchorStyles.Right))));

                _Label_1[item].TabStop = false;
                _Label_1[item].Location = new System.Drawing.Point(tabControl2.TabPages[item].Size.Width - 80, 85);
                _Label_1[item].Text = "Scale";
                _Label_1[item].AutoSize = true;


                _T_2[item].Location = new System.Drawing.Point(tabControl2.TabPages[item].Size.Width - 80, 105);
                _T_2[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                              | System.Windows.Forms.AnchorStyles.Right))));
                _T_2[item].Size = new System.Drawing.Size(80, 40);
                _T_2[item].KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox3_KeyPress);
                _T_2[item].PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.textBox3_KeyDown);

                _T_2[item].TabIndex = 10;

                _T_3[item].Location = new System.Drawing.Point(tabControl2.TabPages[item].Size.Width - 80, 130);
                _T_3[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                              | System.Windows.Forms.AnchorStyles.Right))));
                _T_3[item].Size = new System.Drawing.Size(80, 40);
                _T_3[item].KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox4_KeyPress);
                _T_3[item].PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.textBox4_KeyDown);
                _T_3[item].TabIndex = 11;



                for (i = 0; i < 100; i++)
                {
                    CA[item] = mYChart[item].ChartAreas[0];  // 
                    RA[item] = new RectangleAnnotation();
                    VA[item] = new VerticalLineAnnotation();
                    if (!Ann_Dic[listBox1.Text].ContainsKey(Convert.ToString((i + 1))))
                        Ann_Dic[listBox1.Text].Add(Convert.ToString((i + 1)), new Ann_Class(VA[item], RA[item], CA[item]));
                }

                tabControl2.TabPages[item].Controls.Add(_T[item]);
                tabControl2.TabPages[item].Controls.Add(_T_1[item]);
                tabControl2.TabPages[item].Controls.Add(_Label[item]);

                tabControl2.TabPages[item].Controls.Add(_T_2[item]);
                tabControl2.TabPages[item].Controls.Add(_T_3[item]);
                tabControl2.TabPages[item].Controls.Add(_Label_1[item]);
                tabControl2.TabPages[item].Controls.Add(mYChart[item]);


                tabControl2.TabPages[item].TabStop = false;

                //this.tabPage4.Controls.Add(this.chart1);
                //this.tabPage4.Location = new System.Drawing.Point(10, 47);
                //this.tabPage4.Name = "tabPage4";
                //this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
                //this.tabPage4.Size = new System.Drawing.Size(1314, 837);
                //this.tabPage4.TabIndex = 0;
                //this.tabPage4.Text = "tabPage4";
                //this.tabPage4.UseVisualStyleBackColor = true;

                //      tabControl2.TabPages[item].Size = new System.Drawing.Size(1314, 837);
                tabControl2.TabPages[item].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                                                   | System.Windows.Forms.AnchorStyles.Left)
                                                   | System.Windows.Forms.AnchorStyles.Right)));



                mYChart[item].DataSource = dt[item];
                mYChart[item].ChartAreas[0].AxisX.Interval = 0;
                mYChart[item].Series.Clear();

             //   mYChart[item].ChartAreas[0].AxisX.IsMarginVisible = false;

                mYChart[item].ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
                mYChart[item].ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
                mYChart[item].ChartAreas[0].AxisX.ScrollBar.Enabled = false;
                mYChart[item].ChartAreas[0].AxisY.ScrollBar.Enabled = false;

                mYChart[item].ChartAreas[0].AxisX.LabelStyle.Format = "#.##";
                mYChart[item].ChartAreas[0].AxisY.LabelStyle.Format = "#.#";

                mYChart[item].ChartAreas[0].AxisX.Title = "Frequency [MHz]";
                mYChart[item].ChartAreas[0].AxisX.TitleForeColor = Color.Red;
                mYChart[item].ChartAreas[0].AxisX.TitleAlignment = StringAlignment.Center; // Chart X axis Text Alignment 
                mYChart[item].ChartAreas[0].AxisX.TextOrientation = TextOrientation.Horizontal; // Chart X Axis Text Orientation 
                mYChart[item].ChartAreas[0].AxisX.TitleFont = new Font("Arial", 20, FontStyle.Bold); // Chart X axis Title Font
             

                mYChart[item].ChartAreas[0].AxisY.Title = "Magnitude [DB]";
                mYChart[item].ChartAreas[0].AxisY.TitleForeColor = Color.Red;
                mYChart[item].ChartAreas[0].AxisY.TitleAlignment = StringAlignment.Center; // Chart X axis Text Alignment 
                mYChart[item].ChartAreas[0].AxisY.TextOrientation = TextOrientation.Rotated270; // Chart X Axis Text Orientation 
                mYChart[item].ChartAreas[0].AxisY.TitleFont = new Font("Arial", 20, FontStyle.Bold); // Chart X axis Title Font

             //   mYChart[item].ChartAreas[0].AxisY.Minimum = -300;
             //   mYChart[item].ChartAreas[0].AxisY.Maximum = 100;

                double[] Min = new double[this.Data.Count];
                double[] Max = new double[this.Data.Count];
                int Count = 0;
                foreach (KeyValuePair<string, List<double>> _D in this.Data)
                {
                    //for (int C = 0; C < Units.Count; C++)
                    //{



                    //    double test7 = TestTime1.Elapsed.TotalMilliseconds;

                    //  Series series = chart1.Series.Add(Convert.ToString(Units[C]));

                    //      int Index = Units.BinarySearch(Units[C]);

                    Series series = mYChart[item].Series.Add(_D.Key);

                    series.XValueMember = dt[item].Columns[0].ColumnName;
                    /*      series.YValueMembers = dt.Columns[Index + 1].ColumnName*/
                    ;
                    series.YValueMembers = _D.Key;
                    series.ChartType = SeriesChartType.Line;

                    series.IsVisibleInLegend = true;
                    series.IsValueShownAsLabel = false;
                    series.BorderWidth = 3;


                    series.LegendText = _D.Key;
                    // chart1.Series.add(series);


                    //    List<double> s = this.Data[Data + "_" + Text + "_Unit" + Units[C]];

                    if (k == 0)
                    {
                        for (int d = 0; d < _D.Value.Count; d++)
                        {
                            series.Points.AddXY(Freq[d] / 1e6, _D.Value[d]);
                        }
                    }
                    else
                    {
                        for (int d = 0; d < _D.Value.Count; d++)
                        {
                            series.Points.AddY(_D.Value[d]);
                        }

                    }

                    k++;

                    Min[Count] = _D.Value.Min();
                    Max[Count] = _D.Value.Max();
                    Count++;
                }

                mYChart[item].ChartAreas[0].AxisY.Maximum = Max.Max();
                mYChart[item].ChartAreas[0].AxisY.Minimum = Min.Min();
    
                F++;


            }
            _T[0].Focus();

            if (_Marker_Data == null)
            {
                _Marker_Data = new Dictionary<string, Dictionary<string, string>>();
            }

            dataGridView1.ColumnHeadersVisible = false;
            dataGridView1.ScrollBars = ScrollBars.None;
            dataGridView1.BackgroundColor = Color.White;

            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].Name = "Marker";
            dataGridView1.Columns[1].Name = "Freq (Mhz)";

            Dictionary<string, string> T = new Dictionary<string, string>();
            string Chan = listBox1.SelectedItem.ToString();

            for (i = 0; i < 20; i++)
            {
                object[] Valuse = new object[2];

                Valuse[0] = "Marker" + (i + 1);
                Valuse[1] = "";


                dataGridView1.Rows.Add(Valuse);
                if (!T.ContainsKey(Valuse[0].ToString()))
                    T.Add(Valuse[0].ToString(), Valuse[1].ToString());

            }
            _Marker_Data.Add(Chan, T);

            int width = TextRenderer.MeasureText(dataGridView1.Rows[0].Cells[0].Value.ToString(), dataGridView1.DefaultCellStyle.Font).Width;
            int heigh = TextRenderer.MeasureText(dataGridView1.Rows[0].Cells[0].Value.ToString(), dataGridView1.DefaultCellStyle.Font).Height;

            for (i = 0; i < 20; i++)
            {
                dataGridView1.Rows[i].Height = Convert.ToInt16(heigh * 1.8);
            }


            dataGridView1.Columns[0].Width = Convert.ToInt16(width * 1.2);
            dataGridView1.Columns[1].Width = Convert.ToInt16(dataGridView1.Size.Width - dataGridView1.Columns[0].Width);

            dataGridView1.Visible = true;
            dataGridView1.Enabled = true;

        }

        private void listBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Stopwatch TestTime1 = new Stopwatch();
            TestTime1.Restart();
            TestTime1.Start();

            int item_Test = listBox2.SelectedIndex;
            string Text = listBox1.SelectedItem.ToString();


            this.Data = new Dictionary<string, List<double>>();

            for (int j = 0; j < Units.Count; j++)
            {


                string[] _Data = DB_Interface.Get_Data_By_Query("Select " + listBox2.Items[item_Test] + " from " + Text + " where id =  'Unit" + Units[j] + "'");

                double test3 = TestTime1.Elapsed.TotalMilliseconds;

                List<double> _DummyData = Array.ConvertAll<string, double>(_Data, Convert.ToDouble).ToList();

                if (!this.Data.ContainsKey(listBox2.Items[item_Test].ToString()))
                {
                    // this.Data.Add(listBox2.Items[item].ToString() + "_" + Text + "_Unit" + Units[j], _DummyData);
                    string text = listBox2.Items[item_Test].ToString().Remove(listBox2.Items[item_Test].ToString().Length - 3, 3);
                    this.Data.Add(text + "_SN" + Units[j], _DummyData);

                    foreach (KeyValuePair<string, List<double>> _D in this.Data)
                    {
                        if (!dt[tabControl2.SelectedIndex].Columns.Contains(_D.Key))
                            dt[tabControl2.SelectedIndex].Columns.Add(_D.Key, typeof(double));
                    }

                    int Coulumn_Count = dt[tabControl2.SelectedIndex].Columns.Count - 1;
                    object[] value = new object[1];

                    for (i = 0; i < Freq.Count; i++)
                    {
                        foreach (KeyValuePair<string, List<double>> _D in this.Data)
                        {
                            value[0] = _D.Value[i];
                            //   dt[tabControl2.SelectedIndex].Rows[i].BeginEdit();
                            dt[tabControl2.SelectedIndex].Rows[i][Coulumn_Count] = Convert.ToDouble(value[0]);
                            //   dt[tabControl2.SelectedIndex].Rows[i].EndEdit();
                        }
                    }

                }

                double test6 = TestTime1.Elapsed.TotalMilliseconds;

            }


            int k = 0;

            double test7 = TestTime1.Elapsed.TotalMilliseconds;

            mYChart[tabControl2.SelectedIndex].DataSource = dt[tabControl2.SelectedIndex];
            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisX.Interval = 0;
            //mYChart[tabControl2.SelectedIndex].Series.Clear();

            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisX.IsMarginVisible = false;

            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisX.ScrollBar.Enabled = false;
            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisY.ScrollBar.Enabled = false;

            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisX.LabelStyle.Format = "#.##";
            //mYChart[tabControl2.SelectedIndex].ChartAreas[0].AxisY.LabelStyle.Format = "#.#";

            foreach (KeyValuePair<string, List<double>> _D in this.Data)
            {

                Series series = mYChart[tabControl2.SelectedIndex].Series.Add(_D.Key);

                series.XValueMember = dt[tabControl2.SelectedIndex].Columns[0].ColumnName;

                series.YValueMembers = _D.Key;
                series.ChartType = SeriesChartType.Line;

                series.IsVisibleInLegend = true;
                series.IsValueShownAsLabel = false;
                series.BorderWidth = 3;

                series.LegendText = _D.Key;

                if (k == 0)
                {
                    for (int d = 0; d < _D.Value.Count; d++)
                    {
                        series.Points.AddXY(Freq[d] / 1e6, _D.Value[d]);

                    }
                }
                else
                {
                    for (int d = 0; d < _D.Value.Count; d++)
                    {
                        series.Points.AddY(_D.Value[d]);
                    }

                }

                k++;

            }



            F++;
            double test8 = TestTime1.Elapsed.TotalMilliseconds;
        }

        public double[] Marker_Set_Send(int index, string Test, string Freq_Data, int Row)
        {
            //int index = tabControl2.SelectedIndex;

            #region
            TestAn = Ann_Dic[Test][Convert.ToString(Row)];

            // PositionX = mYChart[index].ChartAreas[0].AxisX.PixelPositionToValue(Selected_X);
            PositionX = Convert.ToDouble(Freq_Data);
            TestAn.VA.AxisX = TestAn.CA.AxisX;
       
            TestAn.VA.IsInfinitive = true;
            TestAn.VA.ClipToChartArea = TestAn.CA.Name;
            TestAn.VA.Name = Convert.ToString("A_" + Test + "_" + Row);
            TestAn.VA.LineDashStyle = ChartDashStyle.Dash;
            TestAn.VA.LineColor = Color.Black;
            TestAn.VA.LineWidth = 2;         // use your numbers!
            TestAn.VA.X = PositionX;

            double xFactor = 0.03;         // use your numbers!
            double yFactor = 0.02;


            TestAn.RA.AxisX = TestAn.CA.AxisX;
            //TestAn.RA.AxisY = TestAn.CA.AxisY;

            TestAn.RA.X = PositionX;
            TestAn.RA.Y = -1.5;
         //   TestAn.RA.AllowMoving = true;
         //   TestAn.RA.IsSizeAlwaysRelative = false;
            TestAn.RA.Width = 20 * xFactor;         // use your numbers!
            TestAn.RA.Height = 8 * yFactor;        // use your numbers!
            TestAn.RA.Name = Convert.ToString("R_" + Test + "_" + Row);
            TestAn.RA.LineColor = Color.Red;
            TestAn.RA.BackColor = Color.Red;



            //TestAn.RA.AllowMoving = true;
            //TestAn.RA.AllowAnchorMoving = true;
            //TestAn.RA.AllowPathEditing = true;
            //TestAn.RA.AllowResizing = true;
            TestAn.RA.AnchorAlignment = ContentAlignment.BottomRight;
            TestAn.RA.Alignment = ContentAlignment.TopCenter;

            //  TestAn.RA.AnchorY = mYChart[index].ChartAreas[0].AxisY.Maximum;

            //  TestAn.RA.Y = mYChart[index].ChartAreas[0].AxisY.Maximum;

            //  TestAn.RA.X = TestAn.VA.X - TestAn.RA.Width / 2;

            TestAn.RA.Text = Convert.ToString(Row);
            TestAn.RA.Font = new System.Drawing.Font("Arial", 15);
            TestAn.RA.ForeColor = Color.Red;
            TestAn.RA.BackColor = Color.White;
            //   RA.Font = new System.Drawing.Font("Arial", 8f);

            if (!mYChart[index].Annotations.Contains(TestAn.VA))
            {
                mYChart[index].Annotations.Add(TestAn.VA);
                mYChart[index].Annotations.Add(TestAn.RA);

            }
        
            j++;

            var pos1 = e_value;

            Axis ax = mYChart[index].ChartAreas[0].AxisX;
            Axis ay = mYChart[index].ChartAreas[0].AxisY;
            //  double x = ax.PixelPositionToValue(pos1.X);
            //   double y = ay.PixelPositionToValue(pos1.Y);

            int Freq_Index = Freq.BinarySearch(PositionX * 1e6);

            double[] data = new double[mYChart[index].Series.Count];

            if (Freq_Index == 0)
            {
                for (i = 0; i < data.Length; i++)
                {
                    data[i] = mYChart[index].Series[i].Points[Freq_Index].YValues[0];
                }
            }
            else if (Freq_Index * -1 > Freq.Count)
            {
                Freq_Index = ~Freq_Index;   // index just after target freq
                for (i = 0; i < data.Length; i++)
                {
                    data[i] = InterpolateLinear(Freq[Freq_Index - 1], Freq[Freq_Index - 1], mYChart[index].Series[i].Points[Freq_Index - 1].YValues[0], mYChart[index].Series[i].Points[Freq_Index - 1].YValues[0], PositionX * 1e6);
                }

            }
            else if (Freq_Index < 0)
            {
                Freq_Index = ~Freq_Index;   // index just after target freq
                for (i = 0; i < data.Length; i++)
                {
                    data[i] = InterpolateLinear(Freq[Freq_Index - 1], Freq[Freq_Index], mYChart[index].Series[i].Points[Freq_Index - 1].YValues[0], mYChart[index].Series[i].Points[Freq_Index].YValues[0], PositionX * 1e6);
                }

            }
       

            else
            {
                for (i = 0; i < data.Length; i++)
                {
             
                    data[i] = InterpolateLinear(Freq[Freq_Index - 1], Freq[Freq_Index], mYChart[index].Series[i].Points[Freq_Index - 1].YValues[0], mYChart[index].Series[i].Points[Freq_Index].YValues[0], PositionX * 1e6);
                }
            }



            // M = new Marker_Form(Test, Freq_Data, tabControl2.TabPages[index].Text.ToString(), mYChart[index], data);
            return data;
            #endregion
        }

        public class Ann_Class
        {
            public VerticalLineAnnotation VA;
            public RectangleAnnotation RA;
            public ChartArea CA;

            public Ann_Class(VerticalLineAnnotation VA, RectangleAnnotation RA, ChartArea CA)
            {
                this.VA = VA;
                this.RA = RA;
                this.CA = CA;

            }
        }

        private PointD LocationInChart(double xMouse, double yMouse)
        {
            var ca = chart1.ChartAreas[0];

            //Position inside the control, from 0 to 100
            var relPosInControl = new PointD
            (
              ((double)xMouse / (double)chart1.Width) * 100,
              ((double)yMouse / (double)chart1.Height) * 100
            );

            //Verify we are inside the Chart Area
            if (relPosInControl.X < ca.Position.X || relPosInControl.X > ca.Position.Right
            || relPosInControl.Y < ca.Position.Y || relPosInControl.Y > ca.Position.Bottom) return new PointD(double.NaN, double.NaN);

            //Position inside the Chart Area, from 0 to 100
            var relPosInChartArea = new PointD
            (
              ((relPosInControl.X - ca.Position.X) / ca.Position.Width) * 100,
              ((relPosInControl.Y - ca.Position.Y) / ca.Position.Height) * 100
            );

            //Verify we are inside the Plot Area
            if (relPosInChartArea.X < ca.InnerPlotPosition.X || relPosInChartArea.X > ca.InnerPlotPosition.Right
            || relPosInChartArea.Y < ca.InnerPlotPosition.Y || relPosInChartArea.Y > ca.InnerPlotPosition.Bottom) return new PointD(double.NaN, double.NaN);

            //Position inside the Plot Area, 0 to 1
            var relPosInPlotArea = new PointD
            (
              ((relPosInChartArea.X - ca.InnerPlotPosition.X) / ca.InnerPlotPosition.Width),
              ((relPosInChartArea.Y - ca.InnerPlotPosition.Y) / ca.InnerPlotPosition.Height)
            );

            var X = relPosInPlotArea.X * (ca.AxisX.Maximum - ca.AxisX.Minimum) + ca.AxisX.Minimum;
            var Y = (1 - relPosInPlotArea.Y) * (ca.AxisY.Maximum - ca.AxisY.Minimum) + ca.AxisY.Minimum;

            return new PointD(X, Y);
        }
        private struct PointD
        {
            public double X;
            public double Y;
            public PointD(double X, double Y)
            {
                this.X = X;
                this.Y = Y;
            }
        }

        public static double InterpLinear(double[] array, double xVal)
        {
            if (xVal <= 0) return array[0];
            if (xVal >= array.Length - 1) return array.Last();

            double lowerX = Math.Floor(xVal);
            double upperX = lowerX + 1;
            double lowerY = array[(int)lowerX];
            double upperY = array[(int)upperX];
            double yInterp = (lowerY + (xVal - lowerX) * (upperY - lowerY) / (upperX - lowerX));

            return yInterp;
        }

        private double InterpolateLinear(double lowerX, double upperX, double lowerY, double upperY, double xTarget)
        {
            try
            {
                return (((upperY - lowerY) * (xTarget - lowerX)) / (upperX - lowerX)) + lowerY;
            }
            catch (Exception e)
            {
                return -99999;
            }
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!TabPage_Flag)
            {
                int total = tabControl2.TabCount;
                int index = tabControl2.SelectedIndex;

                if (tabControl2.TabCount - 1 >= index)
                {
                    _T[index].Focus();
                    _T[index].SelectAll();
                }


                Marker_Setting_Form.Snp_Index = index;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {

            int index = tabControl2.SelectedIndex;

            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {

                    if (_T_1[index].Text == "" || Convert.ToDouble(_T[index].Text) < Convert.ToDouble(_T_1[index].Text))
                    {
                        if (checkBox1.Checked)
                        {
                            for (i = 0; i < mYChart.Length; i++)
                            {
                                mYChart[i].ChartAreas[0].AxisX.Minimum = Convert.ToDouble(_T[index].Text);
                                _T[i].Text = _T[index].Text;
                            }
                        }
                        else
                        {
                            mYChart[index].ChartAreas[0].AxisX.Minimum = Convert.ToDouble(_T[index].Text);
                        }


                        _T_1[index].Focus();
                        _T_1[index].SelectAll();
                    }
                }
                else if (e.KeyChar == (char)Keys.Tab)
                {
                    _T_1[index].Focus();
                    _T_1[index].SelectAll();

                }
            }
            catch
            {

            }
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (_T[index].Text == "" || Convert.ToDouble(_T[index].Text) < Convert.ToDouble(_T_1[index].Text))
                    {
                        if (checkBox1.Checked)
                        {
                            for (i = 0; i < mYChart.Length; i++)
                            {
                                mYChart[i].ChartAreas[0].AxisX.Maximum = Convert.ToDouble(_T_1[index].Text);
                                _T_1[i].Text = _T_1[index].Text;
                            }
                        }
                        else
                        {
                            mYChart[index].ChartAreas[0].AxisX.Maximum = Convert.ToDouble(_T_1[index].Text);
                        }

                        _T[index].Focus();
                        _T[index].SelectAll();
                    }
                }
                else if (e.KeyChar == (char)Keys.Tab)
                {
                    _T[index].Focus();
                }
            }
            catch
            {

            }
        }

        private void textBox1_KeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            try
            {

                if (e.KeyCode == Keys.Tab)
                {
                    //       _T[index].Focus();
                    //       _T[index].SelectAll();

                    //  _T_1[index].Focus();
                    _T_1[index].SelectAll();
                }
            }
            catch
            {

            }
        }
        private void textBox2_KeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            try
            {
                if (e.KeyCode == Keys.Tab)
                {
                    //       _T_1[index].Focus();
                    //       _T_1[index].SelectAll();

                    //   _T[index].Focus();
                    _T[index].SelectAll();
                }
            }
            catch
            {

            }
        }

                
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {

            int index = tabControl2.SelectedIndex;

            try
            {

                if (e.KeyChar == (char)Keys.Enter)
                {

                    if (mYChart[index].ChartAreas[0].AxisY.Maximum > Convert.ToDouble(_T_2[index].Text))
                    {
                        mYChart[index].ChartAreas[0].AxisY.Minimum = Convert.ToDouble(_T_2[index].Text);

                        _T_3[index].Focus();
                        _T_3[index].SelectAll();
                    }


                }
                else if (e.KeyChar == (char)Keys.Tab)
                {
                    _T_3[index].Focus();
                    _T_3[index].SelectAll();

                }
            }
            catch
            {

            }
        }
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            try
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (mYChart[index].ChartAreas[0].AxisY.Minimum < Convert.ToDouble(_T_3[index].Text))
                    {
                        mYChart[index].ChartAreas[0].AxisY.Maximum = Convert.ToDouble(_T_3[index].Text);


                        _T_2[index].Focus();
                        _T_2[index].SelectAll();
                    }
                }
                else if (e.KeyChar == (char)Keys.Tab)
                {
                    _T_2[index].Focus();
                }
            }
            catch
            {

            }
        }

        private void textBox3_KeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            try
            {

                if (e.KeyCode == Keys.Tab)
                {
                    //       _T[index].Focus();
                    //       _T[index].SelectAll();

                    //  _T_1[index].Focus();
                    _T_3[index].SelectAll();
                }
            }
            catch
            {

            }
        }
        private void textBox4_KeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            int index = tabControl2.SelectedIndex;

            try
            {
                if (e.KeyCode == Keys.Tab)
                {
                    //       _T_1[index].Focus();
                    //       _T_1[index].SelectAll();

                    //   _T[index].Focus();
                    _T_2[index].SelectAll();
                }
            }
            catch
            {

            }
        }

        Dictionary<string, Dictionary<string, Dictionary<string, double[]>>> Snp_Data;
        Dictionary<string, Dictionary<string, string>> _Marker_Data;

        Marker_Setting_Form Marker_Form;

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!Delete_Marker_Flag)
            {
                if (e.ColumnIndex == 1 && e.RowIndex > -1)
                {
                    string Chan = listBox1.SelectedItem.ToString();

                    string[] ItemList = new string[listBox2.Items.Count];

                    for (int k = 0; k < listBox2.Items.Count; k++)
                    {
                        ItemList[k] = listBox2.Items[k].ToString();

                    }

                    if (Snp_Data == null)
                        Snp_Data = new Dictionary<string, Dictionary<string, Dictionary<string, double[]>>>();


                    int Row = e.RowIndex;

                    _Marker_Data[Chan]["Marker" + (Row + 1)] = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                    Dictionary<string, Dictionary<string, double[]>> Tag = new Dictionary<string, Dictionary<string, double[]>>();


                    for (int d = 0; d < ItemList.Length; d++)
                    {
                        double[] data = Marker_Set_Send(d, Chan, dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), Row + 1);

                        Dictionary<string, double[]> Row_Data = new Dictionary<string, double[]>();
                        if (!Row_Data.ContainsKey(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()))
                            Row_Data.Add(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), data);

                        if (!Tag.ContainsKey(ItemList[d]))
                            Tag.Add(ItemList[d], Row_Data);

                    }

                    if (!Snp_Data.ContainsKey((e.RowIndex + 1).ToString()))
                        Snp_Data.Add((e.RowIndex + 1).ToString(), Tag);

                    int index = tabControl2.SelectedIndex;



                    if (Marker_Form == null)
                    {
                        Marker_Form = new Marker_Setting_Form(ItemList, Chan, tabControl2.TabPages[index].Text);
                        Marker_Form.Show();

                    }
                    else if (!Marker_Form.Flag_Form)
                    {
                        Marker_Form = new Marker_Setting_Form(ItemList, Chan, tabControl2.TabPages[index].Text);
                        Marker_Form.Show();
                    }

                    Listview_Run(index, e.RowIndex, Marker_Form);

                }
            }
            Delete_Marker_Flag = false;
        }
        public void Listview_Run(int index, int marker, Marker_Setting_Form Marker_Form)
        {

         
            if (Snp_Data.Count == 1)
            {

                string[] ItemList = new string[listBox2.Items.Count];

                for (int k = 0; k < listBox2.Items.Count; k++)
                {
                    ItemList[k] = listBox2.Items[k].ToString();

                }

                int kk = 0;
                for (int k = 0; k < ItemList.Length; k++)
                {

                    Marker_Form.listView1[k].Clear();
                    Marker_Form.listView1[k].BeginUpdate();

                    Marker_Form.listView1[k].View = View.Details;


                    foreach (Dictionary<string, Dictionary<string, double[]>> T in Snp_Data.Values)
                    {

                        int j = 0;

                        foreach (Dictionary<string, double[]> _T in T.Values)
                        {
                            if (k == kk)
                            {

                                //if (j == index)
                                //{
                                    foreach (double[] _T_Data in _T.Values)
                                    {
                                        i = 0;
                                        foreach (double Row in _T_Data)
                                        {
                                            ListViewItem Ivi = new ListViewItem(mYChart[k].Series[i].Name.ToString());

                                            Ivi.SubItems.Add(Convert.ToString(_T_Data[i]));
                                            Marker_Form.listView1[k].Items.Add(Ivi);
                                            i++;
                                        }
                                        Marker_Form.listView1[k].Columns.Add("SN");
                                        Marker_Form.listView1[k].Columns.Add("Marker" + (marker + 1));

                                    }

                               // }
                                j++;
                            }
                            kk++;
                        }
                        kk = 0;
                    }
                    Marker_Form.listView1[k].EndUpdate();
                }

            }

            else 
            {

                string[] ItemList = new string[listBox2.Items.Count];

                for (int k = 0; k < listBox2.Items.Count; k++)
                {
                    Marker_Form.listView1[k].Clear();
                    ItemList[k] = listBox2.Items[k].ToString();

                }

                int SN_Count = 0;
                foreach (Dictionary<string, Dictionary<string, double[]>> T in Snp_Data.Values)
                {
                    foreach (Dictionary<string, double[]> _T in T.Values)
                    {
                        foreach(double[] d in _T.Values)
                        {
                            SN_Count = d.Length;
                        }
                    }
                }

                Dictionary<string, string> Test = _Marker_Data[listBox1.Text];
                List<string> MarkerName = new List<string>();

                foreach(KeyValuePair<string,string> s in Test)
                {
                    if (s.Value != "")
                    {
                        MarkerName.Add(s.Key);
                    }
                }

         

                for (i = 0; i < tabControl2.TabPages.Count; i ++)
                {
                    Marker_Form.listView1[i].Columns.Add("SN");

                    for (int f = 0; f < MarkerName.Count; f++)
                    {
                        Marker_Form.listView1[i].Columns.Add(MarkerName[f]);
                    }
               
                }

                i = 0;

     


                bool flag = false;
                ListViewItem Ivi = new ListViewItem();

                //for (int q = 0; q < SN_Count; q++)
                //{

                int kk = 0;
                int Num = 0;

                for (int k = 0; k < ItemList.Length; k++)
                {
                    Marker_Form.listView1[k].BeginUpdate(); Marker_Form.listView1[k].View = View.Details;

                    while (true)
                    {
                        foreach (Dictionary<string, Dictionary<string, double[]>> T in Snp_Data.Values)
                        {
                            int j = 0; int S = 0;
                            foreach (Dictionary<string, double[]> _T in T.Values)
                            {
                                if (S == kk)
                                {
                                    //if (j == index)
                                    //{
                                        foreach (double[] _T_Data in _T.Values)
                                        {
                                            foreach (double Row in _T_Data)
                                            {
                                                if (!flag)
                                                {
                                                    Ivi = new ListViewItem(mYChart[k].Series[Num].Name.ToString());
                                                    flag = true;
                                                }

                                                Ivi.SubItems.Add(Convert.ToString(_T_Data[Num]));
                                                if (flag)
                                                {
                                                    break;
                                                }
                                            }
                                        }
                                   // }
                                    if (flag)
                                    {
                                        break;
                                    }
                                    j++;
                                }
                                else
                                {
                                    S++;
                                }
                            }
                            if (Snp_Data.Count == Ivi.SubItems.Count - 1)
                            {
                                flag = false;
                                Marker_Form.listView1[k].Items.Add(Ivi);
                            }
                        }
            
                        Num++;

                        if(SN_Count == Num)
                        {
                            Num = 0;
                            break;
                        }
                    }
                    kk++;
                    Marker_Form.listView1[k].EndUpdate();
                }
                //  }

            }
        }

        void chartMonitor_AnnotationPositionChanging(object sender, AnnotationPositionChangingEventArgs e)
        {
            if (sender == TestAn.VA) TestAn.RA.X = TestAn.VA.X - TestAn.RA.Width / 2;
        }

        void chartMonitor_AnnotationPositionChanged(object sender, EventArgs e)
        {
            TestAn.VA.X = (int)(TestAn.VA.X + 0.5);
            TestAn.RA.X = TestAn.VA.X - TestAn.RA.Width / 2;
        }
    }
}
