using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestApplication
{
    public partial class Box_Plot_Form : Form
    {
        string Key = "MERGE";
        string[] Files_path;
        string[] Iden;

        int i = 0;
        int j = 0;

        CSV_Class.CSV CSV = new CSV_Class.CSV();
        CSV_Class.CSV.INT Csv_Interface;

        Data_Class.Data_Editing Data_Edit = new Data_Class.Data_Editing();
        Data_Class.Data_Editing.INT Data_Interface;

        DB_Class.DB_Editing DB = new DB_Class.DB_Editing();
        DB_Class.DB_Editing.INT DB_Interface;

        Dir.Dir_Directory Dir;

        List<Dictionary<string, string>> List_Doc = new List<Dictionary<string, string>>();
        Dictionary<string, string> Dic_Doc;
        public Box_Plot_Form()
        {
            InitializeComponent();

            Dir = new Dir.Dir_Directory("C:\\Automation\\Box_Plot");
            Dir = new Dir.Dir_Directory("C:\\Automation\\Box_Plot\\Add_option");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog Dialog = new OpenFileDialog();

            Dialog.Filter = "DB Files (*.db)| *.db";
            Dialog.InitialDirectory = "C:\\Automation\\DB\\Yield";
            Dialog.Multiselect = true;
            Dialog.ShowDialog();

            if (Dialog.FileNames.Length > 0)
            {
                Csv_Interface = CSV.Open(Key);
                Data_Interface = Data_Edit.Open(Key);
                DB_Interface = DB.Open(Key);

                Csv_Interface.Read_Open(this.Files_path[0]);

                while (!Csv_Interface.StreamReader.EndOfStream)
                {
                    Csv_Interface.Read();
                    Dic_Doc = new Dictionary<string, string>();
                    if (i == 0)
                    {
                        Iden = new string[Csv_Interface.Get_String.Length];

                        for (j = 0; j < Csv_Interface.Get_String.Length; j++)
                        {
                            Iden[j] = Csv_Interface.Get_String[j];
                        }

                    }
                    else
                    {
                        for (j = 0; j < Csv_Interface.Get_String.Length; j++)
                        {
                            Dic_Doc.Add(Iden[j], Csv_Interface.Get_String[j]);
                        }
                        List_Doc.Add(Dic_Doc);
                    }
                    i++;
                }

            }
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                Files_path = (string[])e.Data.GetData(DataFormats.FileDrop);

                this.textBox1.Text = Files_path[0];


            }
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
            }
        }
    }
}
