using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JMP;
using System.Threading;
using System.Web.UI.DataVisualization.Charting;

namespace JMP_Class
{
    public class JMP_Editing
    {
        public class YIELD : INT
        {
            public JMP.Application myJMP { get; set; }
            public JMP.Document myJMPDoc { get; set; }
            public JMP.Document myJMPDoc2 { get; set; }
            public JMP.Bivariate myJMPBiv { get; set; }
            public JMP.Fit myJMPFit { get; set; }
            public JMP.DataTable DT { get; set; }
            public JMP.DataTable DT2 { get; set; }
            public JMP.DataTable TranDT { get; set; }
            public JMP.DataTable joinDt { get; set; }
            public JMP.column Col { get; set; }
            public string Path { get; set; }
            public object[] Value { get; set; }
            public double Convert { get; set; }
            public bool Convert_Flag { get; set; }
            public void Open_Session(bool Visible)
            {
                if (myJMP == null) myJMP = new JMP.Application();
                if (Visible == true) myJMP.Visible = true;
                else myJMP.Visible = false;
            }

            public void Open_Document(string Filepath)
            {

                myJMPDoc = myJMP.OpenDocument(Filepath);
                myJMPDoc.Visible = true;
            }
            public void Open_Document2(string Filepath)
            {

                myJMPDoc2 = myJMP.OpenDocument(Filepath);
                //    myJMPDoc.Visible = true;
            }
            public void GetDataTable()
            {
                try
                {

                    DT = myJMPDoc.GetDataTable();
                    DT.Visible = true;
                }
                catch
                {

                }

            }
            public void GetDataTable2()
            {
                try
                {

                    DT2 = myJMPDoc2.GetDataTable();
                    DT2.Visible = false;
                }
                catch
                {

                }

            }

            public void CloseWindowas()
            {
                try
                {
                    myJMP.CloseAllWindows();

                }
                catch
                {

                }

            }


            public void GetSelect_DataTable(string DataTable)
            {
                try
                {
                    DT = myJMP.GetTableHandleFromName(DataTable);
                    if (DT == null) DT = myJMPDoc.GetDataTable();
                }
                catch
                {

                }
            }

            public bool CheckDoc()
            {
                bool flag = false;

                if (myJMPDoc != null)
                {
                    flag = true;
                }
                return flag;
            }

            public object GetSelected_Row()
            {
                Col = DT.GetColumn("Label");
                object Sel = Col.GetRowStateVectorData(JMP.rowStateConstants.rowStateHidden);

                return Sel;

            }
            public object GetSelected_Gross_Row()
            {
                object Sel = "";
                try
                {
                    Col = DT.GetColumn("Label");
                    Sel = Col.GetRowStateVectorData(JMP.rowStateConstants.rowStateHidden);
                }
                catch
                {

                }
                return Sel;

            }

            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script(Key, Data, X_Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X,By);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<string, CSV_Class.For_Box> By_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script(Key, Data, X_Data, By_Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value,X, By);

                return MakeScript;
            }


            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, string FilePath, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg,  ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script(Key, Data, FilePath, OrderbySequence, Save_Falg, ref Para_Test, Customer_enable, NPI_enable, CPK_enable, CPK_Value);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script(string Key, string Parameter,object[] Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            {
                Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script(Key, Parameter, OrderbySequence, Value, Customer_enable, NPI_enable, CPK_enable, CPK_Value);

                return MakeScript;
            }

            //public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            //{
            //    Value = Data;
            //    JMP_Class.Script MakeScript = new JMP_Class.Script("FitYbyX", Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value);

            //    return MakeScript;
            //}


            public JMP_Class.Script Make_Script_Delete(string Key, string Title)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DELETE", Title);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script_Save(string Key, string Title)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("SAVE", Title);

                return MakeScript;
            }

            public JMP_Class.Script Transpose(string Key, string FileName, List<Dictionary<string, double[]>[]> Gross, object[] ID)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("Transpose", FileName, Gross, ID);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_for_Gross(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross, object[] ID)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_FOR_GROSS", Gross, ID);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_HideAndExclude(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_HideAndExclude", Gross);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_HideAndExclude_1(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_HideAndExclude_1", Gross);

                return MakeScript;
            }


            public void Write(Dictionary<string, object> data)
            {
                DT = myJMP.NewDataTable("Row Data");
            }


            public void Run_Script(string Filepath)
            {
                if (myJMP != null)
                {
                    myJMP.RunJSLFile(Filepath);
                }

            }
            public void Join(string Table)
            {

            }
            public void Join2(string Table, int i)
            {

            }
            public void Close_Dt(string Table)
            {
                if (myJMPDoc != null)
                {
                    // myJMP.CloseAllWindows();
                    try
                    {
                        myJMPDoc.Close(false, "");
                    }
                    catch
                    {

                    }
                    //  myJMPDoc = null;
                }


            }

            public void Close_Dt2(string Table)
            {
                if (myJMPDoc2 != null)
                {
                    // myJMP.CloseAllWindows();
                    try
                    {
                        myJMPDoc2.Close(false, "");
                    }
                    catch
                    {

                    }
                    //  myJMPDoc = null;
                }


            }
            public void Close_JoinDT(string Table)
            {
                joinDt.Document.Close(false, Table);

            }
        }
        public class MERGE : INT
        {
            public JMP.Application myJMP { get; set; }
            public JMP.Document myJMPDoc { get; set; }
            public JMP.Document myJMPDoc2 { get; set; }
            public JMP.Bivariate myJMPBiv { get; set; }
            public JMP.Fit myJMPFit { get; set; }
            public JMP.DataTable DT { get; set; }
            public JMP.DataTable DT2 { get; set; }
            public JMP.DataTable TranDT { get; set; }
            public JMP.DataTable joinDt { get; set; }
            public JMP.column Col { get; set; }
            public string Path { get; set; }
            public object[] Value { get; set; }
            public double Convert { get; set; }
            public bool Convert_Flag { get; set; }
            public void Open_Session(bool Visible)
            {
                if (myJMP == null) myJMP = new JMP.Application();
                if (Visible == true) myJMP.Visible = false;
                else myJMP.Visible = false;

                myJMP.ClearLog();
                myJMP.CloseAllWindows();

            }

            public void Open_Document(string Filepath)
            {

                myJMPDoc = myJMP.OpenDocument(Filepath);
                myJMPDoc.Visible = false;
            }
            public void Open_Document2(string Filepath)
            {

                myJMPDoc2 = myJMP.OpenDocument(Filepath);
                myJMPDoc2.Visible = false;
            }
            public void GetDataTable()
            {
                try
                {

                    DT = myJMPDoc.GetDataTable();
                    DT.Visible = false;
                }
                catch
                {

                }

            }

            public void GetDataTable2()
            {

                DT2 = myJMPDoc2.GetDataTable();
                DT2.Visible = false;


            }

            public void CloseWindowas()
            {
                try
                {
                    myJMP.CloseAllWindows();

                }
                catch
                {

                }

            }


            public void GetSelect_DataTable(string DataTable)
            {
                try
                {
                    DT = myJMP.GetTableHandleFromName(DataTable);
                    if (DT == null) DT = myJMPDoc.GetDataTable();
                }
                catch
                {

                }
            }

            public bool CheckDoc()
            {
                bool flag = false;

                if (myJMPDoc != null)
                {
                    flag = true;
                }
                return flag;
            }

            public object GetSelected_Row()
            {
                Col = DT.GetColumn("Label");
                object Sel = Col.GetRowStateVectorData(JMP.rowStateConstants.rowStateHidden);

                return Sel;

            }
            public object GetSelected_Gross_Row()
            {
                object Sel = "";
                try
                {
                    Col = DT.GetColumn("Label");
                    Sel = Col.GetRowStateVectorData(JMP.rowStateConstants.rowStateHidden);
                }
                catch
                {

                }
                return Sel;

            }

            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script("VARIABLILITY", Data, X_Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value,X, By);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<string, CSV_Class.For_Box> By_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script("VARIABLILITY", Data, X_Data, By_Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X, By);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, string FilePath, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg,ref  List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script("VARIABLILITY", Data, FilePath, OrderbySequence, Save_Falg, ref  Para_Test,  Customer_enable,  NPI_enable, CPK_enable, CPK_Value);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script(string Key, string Parameter, object[] Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            {
                Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION", Parameter, OrderbySequence, Value,  Customer_enable,  NPI_enable, CPK_enable, CPK_Value);

                return MakeScript;
            }

            //public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            //{
            //    Value = Data;
            //    JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION", Data, Customer_enable, NPI_enable, CPK_enable, CPK_Value);

            //    return MakeScript;
            //}


            public JMP_Class.Script Make_Script_Delete(string Key, string Title)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DELETE", Title);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script_Save(string Key, string Title)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("SAVE", Title);

                return MakeScript;
            }
            public JMP_Class.Script Transpose(string Key, string FileName, List<Dictionary<string, double[]>[]> Gross, object[] ID)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("Transpose", FileName, Gross, ID);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_for_Gross(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross, object[] ID)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_FOR_GROSS", Gross, ID);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_HideAndExclude(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_HideAndExclude", Gross);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_HideAndExclude_1(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_HideAndExclude_1", Gross);

                return MakeScript;
            }


            public void Write(Dictionary<string, object> data)
            {
                DT = myJMP.NewDataTable("Row Data");
            }


            public void Run_Script(string Filepath)
            {
                if (myJMP != null)
                {
                    myJMP.RunJSLFile(Filepath);
                }

            }

            public void Join(string Table)
            {


                joinDt = DT.Join(DT2, JMP.dtJoinConstants.dtJoinByRow, Table);
            }
            public void Join2(string Table, int i)
            {

                joinDt = myJMP.GetTableHandleFromName(Table + (i - 1));
                joinDt = joinDt.Join(DT2, JMP.dtJoinConstants.dtJoinByRow, Table + i);
            }

            public void Close_Dt(string Table)
            {
                if (myJMPDoc != null)
                {
                    // myJMP.CloseAllWindows();
                    //try
                    //{
                    DT = myJMP.GetTableHandleFromName(Table);
                    DT.Document.Close(false, Table);

                    //   myJMPDoc.Close(false, "");
                    //}
                    //catch
                    //{

                    //}
                    //  myJMPDoc = null;
                }


            }
            public void Close_Dt2(string Table)
            {
                if (myJMPDoc2 != null)
                {
                    // myJMP.CloseAllWindows();
                    //try
                    //{

                    //  DT2.Document.Close(false, "");
                    DT2 = myJMP.GetTableHandleFromName(Table);
                    DT2.Document.Close(false, Table);


                    //   myJMPDoc2.Close(false, "");


                    //}
                    //catch
                    //{

                    //}
                    //  myJMPDoc = null;
                }


            }

            public void Close_JoinDT(string Table)
            {
                joinDt = myJMP.GetTableHandleFromName(Table);
                joinDt.Document.Close(false, Table);

            }
        }
        public class FCM : INT
        {
            public JMP.Application myJMP { get; set; }
            public JMP.Document myJMPDoc { get; set; }
            public JMP.Document myJMPDoc2 { get; set; }
            public JMP.Bivariate myJMPBiv { get; set; }
            public JMP.Fit myJMPFit { get; set; }
            public JMP.DataTable DT { get; set; }
            public JMP.DataTable DT2 { get; set; }
            public JMP.DataTable TranDT { get; set; }
            public JMP.DataTable joinDt { get; set; }
            public JMP.column Col { get; set; }
            public string Path { get; set; }
            public object[] Value { get; set; }
            public double Convert { get; set; }
            public bool Convert_Flag { get; set; }

            public void Open_Session(bool Visible)
            {
                if (myJMP == null) myJMP = new JMP.Application();
                if (Visible == true) myJMP.Visible = true;
                else myJMP.Visible = false;
            }

            public void Open_Document(string Filepath)
            {

                myJMPDoc = myJMP.OpenDocument(Filepath);
                //    myJMPDoc.Visible = true;
            }
            public void Open_Document2(string Filepath)
            {

                myJMPDoc2 = myJMP.OpenDocument(Filepath);
                //    myJMPDoc.Visible = true;
            }
            public void GetDataTable()
            {
                try
                {

                    DT = myJMPDoc.GetDataTable();
                    DT.Visible = true;


                }
                catch
                {

                }

            }
            public void GetDataTable2()
            {
                try
                {

                    DT2 = myJMPDoc2.GetDataTable();
                    DT2.Visible = false;
                }
                catch
                {

                }

            }

            public void CloseWindowas()
            {
                try
                {

                    myJMP.CloseAllWindows();

                }
                catch
                {

                }

            }


            public void GetSelect_DataTable(string DataTable)
            {
                try
                {
                    DT = myJMP.GetTableHandleFromName(DataTable);
                    if (DT == null) DT = myJMPDoc.GetDataTable();
                }
                catch
                {

                }
            }

            public bool CheckDoc()
            {
                bool flag = false;

                if (myJMPDoc != null)
                {
                    flag = true;
                }
                return flag;
            }

            public object GetSelected_Row()
            {
                Col = DT.GetColumn("Label");
                object Sel = Col.GetRowStateVectorData(JMP.rowStateConstants.rowStateHidden);

                return Sel;

            }
            public object GetSelected_Gross_Row()
            {
                object Sel = "";
                try
                {
                    Col = DT.GetColumn("Label");
                    Sel = Col.GetRowStateVectorData(JMP.rowStateConstants.rowStateHidden);
                }
                catch
                {

                }
                return Sel;

            }

            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script(Key, Data, X_Data,OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value,X, By);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<string, CSV_Class.For_Box> By_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script(Key, Data, X_Data, By_Data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X , By);

                return MakeScript;
            }


            public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, string FilePath, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            {
                // Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script(Key, Data, FilePath, OrderbySequence, Save_Falg , ref Para_Test, Customer_enable, NPI_enable, CPK_enable, CPK_Value);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script(string Key, string Parameter, object[] Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            {
                Value = Data;
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION", Parameter, OrderbySequence, Value, Customer_enable,  NPI_enable, CPK_enable, CPK_Value);

                return MakeScript;
            }

            //public JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
            //{
            //    Value = Data;
            //    JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION", Data, Customer_enable, NPI_enable, CPK_enable, CPK_Value);

            //    return MakeScript;
            //}


            public JMP_Class.Script Make_Script_Delete(string Key, string Title)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DELETE", Title);

                return MakeScript;
            }

            public JMP_Class.Script Make_Script_Save(string Key, string Title)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("SAVE", Title);

                return MakeScript;
            }

            public JMP_Class.Script Transpose(string Key, string FileName, List<Dictionary<string, double[]>[]> Gross, object[] ID)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("Transpose", FileName, Gross, ID);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_for_Gross(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross, object[] ID)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_FOR_GROSS", Gross, ID);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_HideAndExclude(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_HideAndExclude", Gross);

                return MakeScript;
            }
            public JMP_Class.Script Distribution_HideAndExclude_1(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
            {
                JMP_Class.Script MakeScript = new JMP_Class.Script("DISTRIBUTION_HideAndExclude_1", Gross);

                return MakeScript;
            }


            public void Write(Dictionary<string, object> data)
            {
                DT = myJMP.NewDataTable("Row Data");
            }


            public void Run_Script(string Filepath)
            {
                if (myJMP != null)
                {
                    myJMP.RunJSLFile(Filepath);
                }

            }
            public void Join(string Table)
            {

            }
            public void Join2(string Table, int i)
            {

            }
            public void Close_Dt(string Table)
            {
                if (myJMPDoc != null)
                {
                    // myJMP.CloseAllWindows();
                    //try
                    //{



                    //  DT = myJMP.GetTableHandleFromName(Table);
                    DT.Document.Close(false, "");

                    //   myJMPDoc.Close(false, "");
                    //}
                    //catch
                    //{

                    //}
                    //  myJMPDoc = null;
                }


            }
            public void Close_Dt2(string Table)
            {
                if (myJMPDoc2 != null)
                {
                    // myJMP.CloseAllWindows();
                    //try
                    //{

                    //  DT2.Document.Close(false, "");
                    DT2 = myJMP.GetTableHandleFromName(Table);
                    DT2.Document.Close(false, Table);


                    //   myJMPDoc2.Close(false, "");


                    //}
                    //catch
                    //{

                    //}
                    //  myJMPDoc = null;
                }


            }
            public void Close_JoinDT(string Table)
            {
                joinDt.Document.Close(false, Table);

            }
        }
        public interface INT
        {
            JMP.Application myJMP { get; set; }
            JMP.Document myJMPDoc { get; set; }
            JMP.Document myJMPDoc2 { get; set; }
            JMP.Bivariate myJMPBiv { get; set; }
            JMP.Fit myJMPFit { get; set; }
            JMP.DataTable DT { get; set; }
            JMP.DataTable DT2 { get; set; }
            JMP.DataTable joinDt { get; set; }
            JMP.column Col { get; set; }
            string Path { get; set; }
            object[] Value { get; set; }
            double Convert { get; set; }
            bool Convert_Flag { get; set; }
            void Open_Session(bool Visible);
            void Open_Document(string Filepath);
            void Open_Document2(string Filepath);
            void GetDataTable();
            void GetDataTable2();
            void GetSelect_DataTable(string DataTable);
            bool CheckDoc();
            void CloseWindowas();
            object GetSelected_Row();
            object GetSelected_Gross_Row();

            JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By);

            JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, Dictionary<string, CSV_Class.For_Box> By_Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By);


            JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<string, CSV_Class.For_Box> X_Data, string FilePath, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value);
            JMP_Class.Script Make_Script(string Key, string Parameter, object[] Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg,ref List<string>[] Para_Test,bool Customer_enable,bool NPI_enable, bool CPK_enable, double CPK_Value);

        //    JMP_Class.Script Make_Script(string Key, Dictionary<string, CSV_Class.For_Box> Data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value);

            JMP_Class.Script Make_Script_Delete(string Key, string Title);
            JMP_Class.Script Make_Script_Save(string Key, string Title);
            JMP_Class.Script Transpose(string Key, string FileNName, List<Dictionary<string, double[]>[]> Gross, object[] ID);
            JMP_Class.Script Distribution_for_Gross(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross, object[] ID);
            JMP_Class.Script Distribution_HideAndExclude(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross);

            JMP_Class.Script Distribution_HideAndExclude_1(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross);
            void Write(Dictionary<string, object> data);

            //      void Transpose(string Filepath, int Count);
            void Run_Script(string Filepath);
            void Join(string Table);
            void Join2(string Table, int i);

            void Close_Dt(string Table);
            void Close_Dt2(string Table);
            void Close_JoinDT(string Table);



        }

        public INT Open(string Key)
        {
            INT Int = null;
            switch (Key)
            {
                case "YIELD":
                    Int = new YIELD();
                    break;
                case "MERGE":
                    Int = new MERGE();
                    break;
                case "FCM":
                    Int = new FCM();
                    break;
            }
            return Int;
        }
    }

    public class Script
    {

        public string Scrip_Data;

        //  public static double Conver;
        //  public static bool Convert_Flag;
        public Script()
        {

        }

        public Script(string Key, Dictionary<string, CSV_Class.For_Box> data, Dictionary<string, CSV_Class.For_Box> X_data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
        {
            if (Key == "FCM_VARIABLILITY")
            {
                //  this.Scrip_Data = Distribution(data, Key, FilePath, OrderbySequence, Save_Falg, ref Para_Test, Customer_enable, NPI_enable, CPK_enable, CPK_Value);
            }
            else if (Key == "FCM_VARIABLILITY")
            {
                //    this.Scrip_Data = Distribution(data, Key, FilePath, OrderbySequence, Save_Falg, ref Para_Test, Customer_enable, NPI_enable, CPK_enable, CPK_Value);
            }
            else if (Key == "Fit Y X Lot" || Key == "Fit Y X Site")
            {
                this.Scrip_Data = FitYbyX(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X,By);
            }
            else if (Key == "Distributions" || Key == "Distribution")
            {
                this.Scrip_Data = Distribution(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X,By);
            }
            else if (Key == "Seleted_Distributions")
            {
                this.Scrip_Data = Distribution_By_X(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X,By);
            }
        }

        public Script(string Key, Dictionary<string, CSV_Class.For_Box> data, Dictionary<string, CSV_Class.For_Box> X_data, Dictionary<string, CSV_Class.For_Box> By_data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
        {
            if (Key == "FCM_VARIABLILITY")
            {
                //  this.Scrip_Data = Distribution(data, Key, FilePath, OrderbySequence, Save_Falg, ref Para_Test, Customer_enable, NPI_enable, CPK_enable, CPK_Value);
            }
            else if (Key == "FCM_VARIABLILITY")
            {
                //    this.Scrip_Data = Distribution(data, Key, FilePath, OrderbySequence, Save_Falg, ref Para_Test, Customer_enable, NPI_enable, CPK_enable, CPK_Value);
            }




            else if (Key == "Fit Y X:Lot" || Key == "Fit Y X:Site")
            {
                this.Scrip_Data = FitYbyX(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value,X, By);
            }
            else if (Key == "Distributions" || Key == "Distribution")
            {
                this.Scrip_Data = Distribution(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value,X, By);
            }
            else if (Key == "Seleted_Distributions")
            {
                if (By_data.Count == 0)
                {
                    this.Scrip_Data = Distribution(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X,By);
                }
                else
                {
                    this.Scrip_Data = Distribution_By_X(Key, data, By_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X,By);
                }

            }
            else if (Key == "Seleted_Fit_Y_By_X")
            {
                if (X.Length != 0)
                {
                    if(X.Length > 1)
                    {
                        this.Scrip_Data = FitYbyXs(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X, By);
                    }
                    else
                    {
                        this.Scrip_Data = FitYbyX(Key, data, X_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X, By);
                    }
                  
                }
                else if (X.Length == 0 && By.Length != 0)
                {
                    this.Scrip_Data = Distribution_By_X(Key, data, By_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, X, By);
                }
                else
                {
                    //   this.Scrip_Data = FitYbyX(Key, data, By_data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, By);
                }

            }
        }

        public Script(string Key, Dictionary<string, CSV_Class.For_Box> data, string FilePath, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg,  ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
        {
            if (Key == "BoxPlot")
            {
               this.Scrip_Data = Boxplot(data, Key, FilePath, OrderbySequence, Save_Falg, ref Para_Test, Customer_enable,  NPI_enable, CPK_enable, CPK_Value);
            }







            //else if (Key == "FCM_VARIABLILITY")
            //{
            //    this.Scrip_Data = Distribution(data, Key, FilePath, OrderbySequence, Save_Falg, ref Para_Test,  Customer_enable,  NPI_enable, CPK_enable, CPK_Value);
            //}
            //else if (Key == "BOX")
            //{
            //    this.Scrip_Data = Distribution(data, Key, FilePath, OrderbySequence, Save_Falg, ref Para_Test,  Customer_enable,  NPI_enable, CPK_enable, CPK_Value);
            //}
        }

        public Script(string Key, string Parameter, Dictionary<int, Dictionary<int, string>> OrderbySequence, object[] data, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
        {
            if (Key == "DISTRIBUTION")
            {

             //   this.Scrip_Data = Distribution(Parameter, OrderbySequence, data, Customer_enable, NPI_enable, CPK_enable, CPK_Value);
            }

        }

        public Script(string Key, string Parameter, string Parameter2, object[] data, object[] Lot, Data_Class.Data_Editing.Clotho_Spec Spec, string Variation)
        {
            if (Key == "DISTRIBUTION")
            {
            //    this.Scrip_Data = Distribution(Parameter, data, Lot, Spec, Variation);
            }
            else if (Key == "FitYbyX")
            {
              //  this.Scrip_Data = FitYbyX(Parameter, Parameter2, data, Lot, Spec, Variation);
            }

        }

        public Script(string Key, string Title)
        {
            if (Key == "DELETE")
            {
                this.Scrip_Data = DELETE();
            }
            else if (Key == "SAVE")
            {
                this.Scrip_Data = SAVE(Title);
            }

        }

        public Script(string Key, string FileName, List<Dictionary<string, double[]>[]> Gross, object[] ID)
        {
            if (Key == "Transpose")
            {
                this.Scrip_Data = Transpose(Key, FileName, Gross, ID);
            }

        }

        public Script(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross, object[] ID)
        {
            if (Key == "DISTRIBUTION_FOR_GROSS")
            {
                this.Scrip_Data = Distribution_for_Gross(Key, Gross, ID);
            }
        }

        public Script(string Key, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
        {
            if (Key == "DISTRIBUTION_HideAndExclude")
            {
                this.Scrip_Data = Distribution_HideAndExclude(Key, Gross);
            }
            else if (Key == "DISTRIBUTION_HideAndExclude_1")
            {
                this.Scrip_Data = Distribution_HideAndExclude_1(Key, Gross);
            }

        }

        public string Distribution(string Key, Dictionary<string, CSV_Class.For_Box> data, Dictionary<string, CSV_Class.For_Box> X_data, string FilePath, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
        {
            string String_data = "";
            string test = FilePath;

            string paht = FilePath.Substring(FilePath.LastIndexOf("\\") + 1);
            string[] FileSplit = paht.Split('.');
            string FIles = paht.Substring(0, paht.Length - 4);

            string[] split = FilePath.Split('\\');

            FilePath = "";
            for (int ii = 0; ii < split.Length - 1; ii++)
            {
                FilePath += split[ii] + "\\";
            }


            string[] Split = new string[1];
            foreach (string item in data.Keys)
            {
                Split = item.Split('_');
                break;
            }

            var Solt = data.Keys.ToList();
            Solt.Sort();


            int Count_Script = 0;

            List<string>[] For_Special = new List<string>[OrderbySequence.Count];

            foreach (KeyValuePair<int, Dictionary<int, string>> Data in OrderbySequence)
            {
                if (Data.Value.Keys.Contains(888))
                {
                    BoxPlots_By Box_Plot = new BoxPlots_By();

                    For_Special[Count_Script] = new List<string>();
                    #region

                    Para_Test[Count_Script] = new List<string>();

                    int k = 0;

                    string byNum = Data.Value[888];

                    string[] Bynum_Split = byNum.Split('>');
                    string[] ByValue_Split = new string[Bynum_Split.Length];
                    for (k = 0; k < Bynum_Split.Length; k++)
                    {
                        var t = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[k]);
                        ByValue_Split[k] = t.ToString();
                    }


                    Dictionary<int, List<string>> Para_Dic = new Dictionary<int, List<string>>();
                    List<string>[] Para_List = new List<string>[Bynum_Split.Length];
                    List<string>[] Test = new List<string>[Bynum_Split.Length];
                    List<string> Test1 = new List<string>();


                    for (k = 0; k < Bynum_Split.Length; k++)
                    {
                        Para_List[k] = new List<string>();
                        Test[k] = new List<string>();

                    }


                    int index = 0;
                    foreach (KeyValuePair<string, CSV_Class.For_Box> test_data in data)
                    {
                        if (index == 0)
                        {
                            for (int j = 0; j < test_data.Value.WAFER_ID.Length; j++)
                            {
                                Test1.Add("");
                            }

                        }

                        string[] dummy = test_data.Key.Split('_');


                        for (k = 0; k < Bynum_Split.Length; k++)
                        {

                            var content = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[k]);

                            int t = (int)Enum.Parse(typeof(BoxPlot), Bynum_Split[k]);

                            if (t >= 20)
                            {

                                string[] dummy_Test = new string[0];
                                if (t == 20)
                                {
                                    dummy_Test = test_data.Value.SITE_ID.ToArray();
                                }
                                else if (t == 21)
                                {
                                    dummy_Test = test_data.Value.LOT_ID.ToArray();

                                }
                                else if (t == 22)
                                {
                                    dummy_Test = test_data.Value.WAFER_ID.ToArray();
                                }

                                if (index == 0)
                                {
                                    for (int dummy_len = 0; dummy_len < dummy_Test.Length; dummy_len++)
                                    {

                                        if (k == Bynum_Split.Length - 1)
                                            Test1[dummy_len] += content + "," + dummy_Test[dummy_len];
                                        else
                                            Test1[dummy_len] += content + "," + dummy_Test[dummy_len] + ",";
                                    }
                                }
                                //for (int dummy_len = 0; dummy_len < dummy_Test.Length; dummy_len++)
                                //{
                                //    Test[k].Add(content + "," + dummy_Test[dummy_len]);
                                //}


                            }
                            else
                            {

                                string dummy_Test = dummy[Convert.ToInt16(Bynum_Split[k])];

                                if (index == 0)
                                {
                                    for (int dummy_len = 0; dummy_len < test_data.Value.SITE_ID.Length; dummy_len++)
                                    {

                                        if (k == Bynum_Split.Length - 1)
                                            Test1[dummy_len] += content + "," + dummy_Test;
                                        else
                                            Test1[dummy_len] += content + "," + dummy_Test + ",";
                                    }
                                }


                                //   Test[k].Add(content + "," + dummy_Test);

                            }

                        }
                        index++;
                    }

                    Test1 = Test1.Distinct().ToList();

                    for (int m = 0; m < Test1.Count; m++)
                    {
                        Para_Test[Count_Script].Add(Test1[m]);
                        //    For_Special[Count_Script].Add(Test1[0][m]);
                    }


                    Box_Plot.DT_Open("");
                    Box_Plot.SendToByGroup(Para_Test, Bynum_Split, Data.Value[999], Count_Script);
                    Box_Plot.X(Data);

                    Box_Plot.Setting(Bynum_Split);

                    Dictionary<string, CSV_Class.For_Box>[] Data_Test1 = new Dictionary<string, CSV_Class.For_Box>[Para_Test[Count_Script].Count];
                    Dictionary<string, CSV_Class.For_Box> Data_Test = new Dictionary<string, CSV_Class.For_Box>();

                    Find_Sequence_Parameter(data, Data.Value, ref Data_Test);

                    for (int p = 0; p < Para_Test[Count_Script].Count; p++)
                    {
                        Data_Test1[p] = new Dictionary<string, CSV_Class.For_Box>();

                        foreach (KeyValuePair<string, CSV_Class.For_Box> _S in Data_Test)
                        {
                            string[] _D = _S.Value.Parameter.Split('_');

                            string[] _Split = Para_Test[Count_Script][p].Split(',');

                            bool _Falg = true;
                            int bulk = 0;

                            for (k = 0; k < _Split.Length; k++)
                            {
                                if (_Split[k] == "Lot")
                                {
                                    if (!_S.Value.LOT_ID.Contains(_Split[k + 1])) _Falg = false;

                                }
                                else if (_Split[k] == "Site")
                                {
                                    if (!_S.Value.SITE_ID.Contains(_Split[k + 1])) _Falg = false;


                                }
                                else if (_Split[k] == "Wafer")
                                {
                                    if (!_S.Value.WAFER_ID.Contains(_Split[k + 1])) _Falg = false;
                                }

                                else if (_D[Convert.ToInt16(Bynum_Split[k / 2])] == _Split[k + 1])
                                {

                                    bulk++;
                                }
                                else
                                {
                                    _Falg = false;
                                }
                                k++;

                            }

                            if (_Falg)
                            {
                                Data_Test1[p].Add(_S.Key, _S.Value);
                            }

                        }

                    }


                    Box_Plot.Dispatch_Set_MinMax(Data.Value[999], Bynum_Split, Para_Test, Data_Test1, Data, Count_Script);
                    Box_Plot.End();

                    String_data += Box_Plot.Jmp_Script;
                    Count_Script++;
                    #endregion

                }

                else
                {
                    #region

                    BoxPlots Box_Plot = new BoxPlots();


                    Box_Plot.DT_Open("");
                    Box_Plot.Y(Split[1]);
                    Box_Plot.X(Data);

                    Box_Plot.Setting("");

                    Box_Plot.SendReport();

                    Dictionary<string, CSV_Class.For_Box> Data_Test = new Dictionary<string, CSV_Class.For_Box>();

                    Find_Sequence_Parameter(data, Data.Value, ref Data_Test);

                    Box_Plot.Dispatch_Set_MinMax(Split[1], Data_Test, Data);

                    var varList = Data_Test.Keys.ToList();
                    varList.Sort();

                    string[] List = varList.ToArray();

                    Box_Plot.Dispatch_Set_SpecLine(Split[1], List, Data_Test);

                    Box_Plot.End();

                    Count_Script++;

                    String_data = Box_Plot.Jmp_Script;
                    #endregion
                }


            }



            return String_data;
        }

        public string Boxplot(Dictionary<string, CSV_Class.For_Box> data, string Key, string FilePath, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Save_Falg, ref List<string>[] Para_Test, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
        {
            string String_data = "";
            string test = FilePath;

            string paht = FilePath.Substring(FilePath.LastIndexOf("\\") + 1);
            string[] FileSplit = paht.Split('.');
            string FIles = paht.Substring(0, paht.Length - 4);

            string[] split = FilePath.Split('\\');

            FilePath = "";
            for (int ii = 0; ii < split.Length - 1; ii++)
            {
                FilePath += split[ii] + "\\";
            }


            string[] Split = new string[1];
            foreach (string item in data.Keys)
            {
                Split = item.Split('_');
                break;
            }

            var Solt = data.Keys.ToList();
            Solt.Sort();


            int Count_Script = 0;

            List<string>[] For_Special = new List<string>[OrderbySequence.Count];

            foreach (KeyValuePair<int, Dictionary<int, string>> Data in OrderbySequence)
            {
                if (Data.Value.Keys.Contains(888))
                {
                    BoxPlots_By Box_Plot = new BoxPlots_By();

                    For_Special[Count_Script] = new List<string>();
                    #region

                    Para_Test[Count_Script] = new List<string>();

                    int k = 0;

                    string byNum = Data.Value[888];

                    string[] Bynum_Split = byNum.Split('>');
                    string[] ByValue_Split = new string[Bynum_Split.Length];
                    for (k = 0; k < Bynum_Split.Length; k++)
                    {
                        var t = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[k]);
                        ByValue_Split[k] = t.ToString();
                    }


                    Dictionary<int, List<string>> Para_Dic = new Dictionary<int, List<string>>();
                    List<string>[] Para_List = new List<string>[Bynum_Split.Length];
                    List<string>[] Test = new List<string>[Bynum_Split.Length];
                    List<string> Test1 = new List<string>();


                    for (k = 0; k < Bynum_Split.Length; k++)
                    {
                        Para_List[k] = new List<string>();
                        Test[k] = new List<string>();
            
                    }


                    int index = 0;
                    foreach (KeyValuePair<string, CSV_Class.For_Box> test_data in data)
                    {
                        if(index == 0)
                        {
                            for(int j = 0; j < test_data.Value.WAFER_ID.Length; j ++)
                            {
                                Test1.Add("");
                            }
                      
                        }

                        string[] dummy = test_data.Key.Split('_');
                

                        for (k = 0; k < Bynum_Split.Length; k++)
                        {

                            var content = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[k]);

                            int t = (int)Enum.Parse(typeof(BoxPlot), Bynum_Split[k]);

                            if(t >= 20)
                            {

                                string[] dummy_Test = new string[0];
                                if (t == 20)
                                {
                                   dummy_Test = test_data.Value.SITE_ID.ToArray();
                                }
                                else if(t == 21)
                                {
                                    dummy_Test = test_data.Value.LOT_ID.ToArray();

                                }
                                else if(t == 22)
                                {
                                    dummy_Test = test_data.Value.WAFER_ID.ToArray();
                                }

                                if (index == 0)
                                {
                                    for (int dummy_len = 0; dummy_len < dummy_Test.Length; dummy_len++)
                                    {

                                        if (k == Bynum_Split.Length - 1)
                                            Test1[dummy_len] += content + "," + dummy_Test[dummy_len];
                                        else
                                            Test1[dummy_len] += content + "," + dummy_Test[dummy_len] + ",";
                                    }
                                }
                                //for (int dummy_len = 0; dummy_len < dummy_Test.Length; dummy_len++)
                                //{
                                //    Test[k].Add(content + "," + dummy_Test[dummy_len]);
                                //}

                          
                            }
                            else
                            {

                                string dummy_Test = dummy[Convert.ToInt16(Bynum_Split[k])];

                                if (index == 0)
                                {
                                    for (int dummy_len = 0; dummy_len < test_data.Value.SITE_ID.Length; dummy_len++)
                                    {
                                
                                        if (k == Bynum_Split.Length - 1)
                                            Test1[dummy_len] += content + "," + dummy_Test;
                                        else
                                            Test1[dummy_len] += content + "," + dummy_Test + ",";
                                    }
                                }


                             //   Test[k].Add(content + "," + dummy_Test);

                            }
             
                        }
                        index++;
                    }

                    Test1 = Test1.Distinct().ToList();

                    for (int m = 0; m < Test1.Count; m ++)
                    {
                        Para_Test[Count_Script].Add(Test1[m]);
                    //    For_Special[Count_Script].Add(Test1[0][m]);
                    }
              

                    Box_Plot.DT_Open("");
                    Box_Plot.SendToByGroup(Para_Test, Bynum_Split, Data.Value[999], Count_Script);
                    Box_Plot.X(Data);

                    Box_Plot.Setting(Bynum_Split);

                    Dictionary<string, CSV_Class.For_Box>[] Data_Test1 = new Dictionary<string, CSV_Class.For_Box>[Para_Test[Count_Script].Count];
                    Dictionary<string, CSV_Class.For_Box> Data_Test = new Dictionary<string, CSV_Class.For_Box>();

                    Find_Sequence_Parameter(data, Data.Value, ref Data_Test);

                    for (int p = 0; p < Para_Test[Count_Script].Count; p++)
                    {
                        Data_Test1[p] = new Dictionary<string, CSV_Class.For_Box>();

                        foreach (KeyValuePair<string, CSV_Class.For_Box> _S in Data_Test)
                        {
                            string[] _D = _S.Value.Parameter.Split('_');

                            string[] _Split = Para_Test[Count_Script][p].Split(',');

                            bool _Falg = true;
                            int bulk = 0;

                            for (k = 0; k < _Split.Length; k++)
                            {
                                if(_Split[k] == "Lot")
                                {
                                    if(!_S.Value.LOT_ID.Contains(_Split[k+1])) _Falg = false;
    
                                }
                                else if (_Split[k] == "Site")
                                {
                                    if (!_S.Value.SITE_ID.Contains(_Split[k + 1])) _Falg = false;


                                }
                                else if (_Split[k] == "Wafer")
                                {
                                    if (!_S.Value.WAFER_ID.Contains(_Split[k + 1])) _Falg = false;
                                }

                                else if (_D[Convert.ToInt16(Bynum_Split[k / 2])] == _Split[k + 1])
                                {

                                    bulk++;
                                }
                                else
                                {
                                    _Falg = false;
                                }
                                k++;
                                
                            }

                            if (_Falg)
                            {
                                Data_Test1[p].Add(_S.Key, _S.Value);
                            }

                        }

                    }


                    Box_Plot.Dispatch_Set_MinMax(Data.Value[999], Bynum_Split, Para_Test, Data_Test1, Data, Count_Script);
                    Box_Plot.End();

                    String_data += Box_Plot.Jmp_Script;
                    Count_Script++;
                    #endregion

                }

                else
                {
                    #region

                    BoxPlots Box_Plot = new BoxPlots();


                    Box_Plot.DT_Open("");
                    Box_Plot.Y(Split[1]);
                    Box_Plot.X(Data);

                    Box_Plot.Setting("");

                    Box_Plot.SendReport();

                    Dictionary<string, CSV_Class.For_Box> Data_Test = new Dictionary<string, CSV_Class.For_Box>();

                    Find_Sequence_Parameter(data, Data.Value, ref Data_Test);

                    Box_Plot.Dispatch_Set_MinMax(Split[1], Data_Test , Data);

                    var varList = Data_Test.Keys.ToList();
                    varList.Sort();

                    string[] List = varList.ToArray();

                    Box_Plot.Dispatch_Set_SpecLine(Split[1],List, Data_Test);

                    Box_Plot.End();

                    Count_Script++;

                    String_data = Box_Plot.Jmp_Script;
                    #endregion
                }


            }



            return String_data;
        }

        public string Distribution(string Parameter, Dictionary<string, CSV_Class.For_Box> data, Dictionary<string, CSV_Class.For_Box> X_data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
        {

            Distribution Dist = new Distribution();

            Dist.Dist();

            int i = 0;
            bool flag = false;
            foreach (KeyValuePair<string, CSV_Class.For_Box> _D in data)
            {
                Dist.Continuous_Distribution();
                Dist.Column(_D.Key);
                Dist.Horizontal_Layout();
                Dist.Vertical();
                Dist.Capability_Analysis(_D, OrderbySequence, Customer_enable, NPI_enable, CPK_enable,  CPK_Value);
             //   Dist.SendtoReport();
              //  Dist.Dispatch();
              //  Dist.Spec(_D.Value.Apple_Spec_Min, _D.Value.Apple_Spec_Max, _D.Value.Broadcom_Spec_Min, _D.Value.Broadcom_Spec_Max);

                if (i == data.Count - 1)
                {
                    flag = true;
                }
                Dist.End(flag);
                i++;

            }

            return Dist.Jmp_Script;
        }

        public string Distribution_By_X(string Parameter, Dictionary<string, CSV_Class.For_Box> data, Dictionary<string, CSV_Class.For_Box> By_data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
        {
            List<string>[] Para_Test;
            int Start_index = data.Count - OrderbySequence.Count;

            Find_Contain_Key(By_data, OrderbySequence , out Para_Test);

            Distribution_By_X Dist_By_X = new Distribution_By_X();

            Dist_By_X.Dist();

            int i = 0;
            bool flag = false;

            foreach (KeyValuePair<string, CSV_Class.For_Box> _D in data)
            {
                Dist_By_X.Continuous_Distribution(_D, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value);
            }

            Dist_By_X.By(By_data, Start_index);

      
   
            for(i = 0; i < Para_Test[0].Count; i ++)
            {
                bool Flag = false;

                if(i == Para_Test[0].Count - 1)
                {
                    Flag = true;
                }
                Dist_By_X.Capability_Analysis(Para_Test[0][i], data, OrderbySequence, Customer_enable, NPI_enable, CPK_enable, CPK_Value, Flag);
            }


               // Dist_By_X.End(flag);
                i++;

            return Dist_By_X.Jmp_Script;
        }

        public string FitYbyX(string Parameter, Dictionary<string, CSV_Class.For_Box> data, Dictionary<string, CSV_Class.For_Box> X_data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
        {

            FitYbyX Fit = new FitYbyX();

            Fit.Fit_Group();

            int i = 0;
            bool flag = false;
            foreach( KeyValuePair<string, CSV_Class.For_Box> _D in data)
            {
                Fit.Oneway();
                Fit.Y(_D.Key);
                Fit.X(X);
                Fit.Quantiles();
                Fit.BoxPlots();
                Fit.MeansandStdDev();
                Fit.StddevLines();
                Fit.GrandMean();
                Fit.SendtoReport();
                Fit.Dispatch();
                Fit.Spec(_D.Value.Apple_Spec_Min, _D.Value.Apple_Spec_Max, _D.Value.Broadcom_Spec_Min, _D.Value.Broadcom_Spec_Max);

                if( i == data.Count - 1)
                {
                    flag = true;
                }
                Fit.End(flag);
                i++;

            }




    

            return Fit.Jmp_Script;
        }

        public string FitYbyXs(string Parameter, Dictionary<string, CSV_Class.For_Box> data, Dictionary<string, CSV_Class.For_Box> X_data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, string[] X, string[] By)
        {

            FitYbyXs Fit = new FitYbyXs();

            Fit.Fit_Group();

            int i = 0;
            int k = 0;
     
            int end = data.Count * X_data.Count;
            bool flag = false;
            foreach (KeyValuePair<string, CSV_Class.For_Box> _D in data)
            {
                int y = 1;
                foreach (KeyValuePair<string, CSV_Class.For_Box> _X in X_data)
                {


                    Fit.Oneway();
                    Fit.Y(_D.Key);
                    Fit.X(X);

                    Fit.SendtoReport();
                    Fit.Dispatch(y);
                    Fit.Spec(_D.Value.Apple_Spec_Min, _D.Value.Apple_Spec_Max, _D.Value.Broadcom_Spec_Min, _D.Value.Broadcom_Spec_Max);

                    Fit.Dispatch(y + 1);
                    Fit.Spec(_X.Value.Apple_Spec_Min, _X.Value.Apple_Spec_Max, _X.Value.Broadcom_Spec_Min, _X.Value.Broadcom_Spec_Max);

                    if (i == data.Count - 1 && end - 1 == k)
                    {
                        flag = true;
                    }
                    Fit.End(flag);
                    y++;
                    k++;
                }
           
                i++;

            }






            return Fit.Jmp_Script;
        }

        public string DELETE()
        {
            string String_data = "";

            String_data = "dt = currentdatatable();";
            String_data += "close(dt, \"No Save\");";
            //String_data = "Close(currentdatatable(), No save);";

            return String_data;
        }

        public string SAVE(string Title)
        {
            string String_data = "";

            String_data = "dt = currentdatatable();";

            String_data += "dt << Save (\"" + Title + ".jmp\" ,jmp);";
            String_data += "dt << Save (\"" + Title + ".csv\" ,text);";
            String_data += "Close( currentdatatable(), No Save );";
            //String_data = "Close(currentdatatable(), No save);";

            return String_data;
        }

        public string Transpose(string Parameter, string FileName, List<Dictionary<string, double[]>[]> Gross, object[] ID)
        {
            string String_data = "";
            string File = FileName.Substring(FileName.LastIndexOf("\\") + 1);
            File = File.Replace(".csv", "");
            String_data = "Data Table(\"" + File + "\") <<";
            String_data += "Transpose(";
            String_data += "columns(";



            for (int i = 0; i < ID.Length - 1; i++)
            {
                String_data += ":Name( \"" + ID[i].ToString() + "\"),";
            }
            String_data += ":Name( \"" + ID[ID.Length - 1].ToString() + "\")";
            //foreach (object o item in ID)
            //{


            //}

            //foreach (Dictionary<string, double[]>[] item in Gross)
            //{
            //    foreach (Dictionary<string, double[]> items in item)
            //    {
            //        int j = 0;
            //        foreach (KeyValuePair<string, double[]> o in items)
            //        {
            //            for (int i = 0; i < o.Value.Length - 1; i++)
            //            {
            //                String_data += ":Name( \"" + Convert.ToString(i + 1) + "\"),";
            //            }
            //            String_data += ":Name( \"" + Convert.ToString(o.Value.Length) + "\")";
            //            break;
            //        }
            //        if(falg)
            //        {
            //            break;
            //        }

            //    }
            //    if (falg)
            //    {
            //        break;
            //    }

            //}

            String_data += "),";
            String_data += "Label( :\"parameter\" ),";

            String_data += "Output Table( \"Transpose of " + File + "\"));";
            String_data += "dt = data table(\"Data\");";
            String_data += "Close(dt, \"No Save\");";

            String_data += "dt = Data Table(\"Transpose of " + File + "\");";
            String_data += "dt << Save(\"C:\\temp\\dummy\\dummy.jmp\");";

            return String_data;
        }

        public string Distribution_for_Gross(string Parameter, Dictionary<string, DB_Class.DB_Editing.Gross> Gross, object[] ID)
        {
            string String_data = "";


            String_data = "dt = currentdatatable();";

            String_data += "Distribution(";
            String_data += "Stack(1),";

            int j = 0;

            foreach (KeyValuePair<string, DB_Class.DB_Editing.Gross> item in Gross)
            {

                String_data += "Continuous Distribution(";
                String_data += "Column(";

                String_data += ":Name( \"" + item.Key.ToString() + "\")),";
                String_data += "Horizontal Layout( 1 ),";
                String_data += "Vertical( 0 ),";

                if (item.Value.SpecL <= -100)
                {
                    if (item.Value.SpecH > 500)
                    {

                    }
                    else
                    {
                        String_data += " Capability Analysis(USL(" + item.Value.SpecH + "))";
                    }

                }
                if (item.Value.SpecH > 500)
                {
                    if (item.Value.SpecL < -500)
                    {

                    }
                    else
                    {
                        String_data += " Capability Analysis(LSL(" + item.Value.SpecL + "))";
                    }
                }

                if (item.Value.SpecL > -100 && item.Value.SpecH < 500)
                {
                    String_data += " Capability Analysis(LSL(" + item.Value.SpecL + "),USL(" + item.Value.SpecH + "))";
                }


                j++;
                if (j == Gross.Count)
                {
                    String_data += ")";
                }
                else
                {
                    String_data += "),";
                }
            }
            String_data += ");";


            return String_data;
        }

        public string Distribution_HideAndExclude(string Parameter, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
        {
            string String_data = "";

            if (Gross.Count != 0)
            {
                String_data = "dt = currentdatatable();";

                foreach (KeyValuePair<string, DB_Class.DB_Editing.Gross> item in Gross)
                {
                    String_data += "dt << select where( :Label == \"" + item + "\");";
                    String_data += "dt << hide and exclude;";
                }
            }

            return String_data;
        }
        public string Distribution_HideAndExclude_1(string Parameter, Dictionary<string, DB_Class.DB_Editing.Gross> Gross)
        {
            string String_data = "";

            if (Gross.Count != 0)
            {
                String_data = "dt = currentdatatable();";

                //foreach (string item in Gross)
                //{
                //    String_data += "dt << select where( :Label == " + item + ");";
                //    String_data += "dt << hide and exclude;";
                //}
            }

            return String_data;
        }

        public void Find_Sequence_Parameter(Dictionary<string, CSV_Class.For_Box> data, Dictionary<int, string> OrderbySequence, ref Dictionary<string, CSV_Class.For_Box> Data_Test)
        {
            Dictionary<string, CSV_Class.For_Box> Data_For_BoxPlot = new Dictionary<string, CSV_Class.For_Box>();

            string[] Datas = new string[0];
            List<string> test = new List<string>();
            int i = 0;

            foreach (KeyValuePair<string, CSV_Class.For_Box> T in data)
            {
                i = 0;
                Dictionary<string, List<string>> Data_Dic = new Dictionary<string, List<string>>();

                string[] _D = new string[0];

                foreach (KeyValuePair<int, string> D in OrderbySequence)
                {
                    List<string> Data_List = new List<string>();

                    int Dic_Count = 0;
                    if(OrderbySequence.Keys.Contains(888))
                    {
                        Dic_Count = OrderbySequence.Count - 2;
                        _D = OrderbySequence[888].Split('>');

                    }
                    else
                    {
                        Dic_Count = OrderbySequence.Count - 1;
                    }
                    if (i < Dic_Count)
                    {
                        if (i != 0)
                        {
                        

                            if (D.Value.ToUpper() == "SITE")
                            {
                                CSV_Class.For_Box Data = T.Value;

                                string[] Dummy = Data.SITE_ID.Distinct().ToArray();

                                Data_List = Data_List.Concat(Dummy).ToList();
                                Data_Dic.Add(D.Value, Data_List);

                            }
                            else if (D.Value.ToUpper() == "LOT")
                            {
                                CSV_Class.For_Box Data = T.Value;

                                string[] Dummy = Data.LOT_ID.Distinct().ToArray();
                                Data_List = Data_List.Concat(Dummy).ToList();
                                Data_Dic.Add(D.Value, Data_List);
                            }
                            else if (D.Value.ToUpper() == "WAFER")
                            {
                                CSV_Class.For_Box Data = T.Value;

                                string[] Dummy = Data.WAFER_ID.Distinct().ToArray();
                                Data_List = Data_List.Concat(Dummy).ToList();
                                Data_Dic.Add(D.Value, Data_List);
                            }
                            else if (D.Value == "")
                            {



                            }
                            else
                            {
                                CSV_Class.For_Box Data = T.Value;

                                string[] Dummy = Data.Parameter.Split('_');
                                int Index = D.Key;

                                Data_List.Add(Dummy[Index]);
                                Data_Dic.Add(D.Value, Data_List);
                            }
                        }
                    }

                    if( i == Dic_Count - 1)
                    {
                        if(_D.Length != 0)
                        {
                            for(int u = 0; u < _D.Length; u++)
                            {
                                Data_List = new List<string>();

                                CSV_Class.For_Box Data = T.Value;

                                string[] Dummy = Data.Parameter.Split('_');
                                int Index = Convert.ToInt16(_D[u]);

                                int t = (int)Enum.Parse(typeof(BoxPlot), _D[u]);

                                if (t >= 20)
                                {
                                    var content = (BoxPlot)Enum.Parse(typeof(BoxPlot), _D[u]);
                                    Data_List.Add(Convert.ToString(content));
                                }
                                else
                                {
                                    Data_List.Add(Dummy[Index]);
                                }
                           

                                Data_Dic.Add(t.ToString(), Data_List);
                            }

                        }

                    }
                    i++;

                }
        
                int j = 0;
                int jj = 0;
                int Max = 0;

                foreach (KeyValuePair<string, List<string>> S in Data_Dic)
                {
                    j = S.Value.Count;

                    if (j > jj)
                    {
                        Max = j;
                    }
                    jj = j;
                }

                string[] Array = new string[Max];

                j = 0;

                foreach (KeyValuePair<string, List<string>> S in Data_Dic)
                {
                    if (S.Value.Count > 1)
                    {
                        for (int k = 0; k < S.Value.Count; k++)
                        {
                            Array[k] += S.Value[k];
                        }
                    }
                    else
                    {
                        for (int k = 0; k < Array.Length; k++)
                        {
                            Array[k] += S.Value[0];
                        }
                    }
                }

                test = test.Concat(Array).ToList();

                for (int l = 0; l < test.Count; l++)
                {
                    if (!Data_Test.Keys.Contains(test[l]))
                    {
                        Data_Test.Add(test[l], T.Value);
                    }

                }
                test = new List<string>();
            }

            //var varList = Data_Test.Keys.ToList();
            //varList.Sort();


            i++;


        }

        public void Find_Contain_Key(Dictionary<string, CSV_Class.For_Box> data, Dictionary<int, Dictionary<int, string>> OrderbySequence, out List<string>[] Para_Test)
        {
            int Start_index = data.Count - OrderbySequence.Count;
            int i = 0;

            List<string> Test1 = new List<string>();

            Dictionary<string, string[]> dummy_Test = new Dictionary<string, string[]>();

            foreach (KeyValuePair<string, CSV_Class.For_Box> Data in data)
            {

                CSV_Class.For_Box ss = Data.Value;

                if (ss.data == null)
                {

                    dummy_Test.Add(Data.Key.ToString(), Array.ConvertAll<object, string>(ss.data_object, Convert.ToString));
                }
                else
                {

                    dummy_Test.Add(Data.Key.ToString(), Array.ConvertAll<double, string>(ss.data, Convert.ToString));
                }


            }

            string[] key = dummy_Test.Keys.ToArray();
            int k = 0;

            foreach (KeyValuePair<string, string[]> T in dummy_Test)
            {
                i = 0;

                if(dummy_Test.Count == 1)
                {
                 
                    for (int h = 0; h < T.Value.Length; h++)
                    {
                        Test1.Add("");
                        string[] s = dummy_Test[key[0]];

                        if(s[h] == "PT9133008440-E")
                        {

                        }
                        Test1[k] += key[i] + "," + s[h];

                        k++;
                
                    }
                }
                else
                {
                    for (int h = 0; h < T.Value.Length; h++)
                    {
                        Test1.Add("");
                        string[] s = dummy_Test[key[i]];

                        if (s[h] == "PT9133008440-E")
                        {

                        }

                        Test1[k] += key[i] + "," + s[h] + ",";

                        i++;

                        s = dummy_Test[key[i]];

                        Test1[k] += key[i] + "," + s[h];

                        i = 0;
                        k++;
                    }
                }
             

                break;

            }

            Test1 = Test1.Distinct().ToList();

            Test1 = Test1.Distinct().ToList();
            Para_Test = new List<string>[1];

            Para_Test[0] = new List<string>();

            for (int m = 0; m < Test1.Count; m++)
            {
                Para_Test[0].Add(Test1[m]);
                //    For_Special[Count_Script].Add(Test1[0][m]);
            }
        }

    }

    public class Distribution
    {
        public string Jmp_Script;

        public Distribution()
        {
            Jmp_Script = "";
        }
        public void Dist()
        {
            Jmp_Script = "Distribution(";
            Jmp_Script += "\n";
            Jmp_Script += "Stack(1),";
            Jmp_Script += "\n";
        }

        public void Continuous_Distribution()
        {
            Jmp_Script += "Continuous Distribution(";
            Jmp_Script += "\n";
        }

        public void Column(string Name)
        {
            Jmp_Script += "Column( :NAME(\"" + Name + "\")),";
            Jmp_Script += "\n";

        }
        public void Horizontal_Layout()
        {

            Jmp_Script += "Horizontal Layout(1),";
            Jmp_Script += "\n";
        }
        public void Vertical()
        {

            Jmp_Script += "Vertical(0),";
            Jmp_Script += "\n";
        }

        public void Capability_Analysis(KeyValuePair<String, CSV_Class.For_Box> data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
        {

            string Spec_Ig = "";
            string[] split = data.Key.Split('_');
            bool Flag = false;

            if (data.Value.data_object != null)
            {
                Flag = true;

                Jmp_Script += "\n";
                Jmp_Script += "SendToReport(";
                Jmp_Script += "\n";
                Jmp_Script += "Dispatch(";
                Jmp_Script += "\n";
                Jmp_Script += "{ \"" + data.Key + "\" },";
                Jmp_Script += "" + "1" + ",";
                Jmp_Script += "),";
                Jmp_Script += "\n";
                Jmp_Script += "Dispatch(";
                Jmp_Script += "\n";
                Jmp_Script += "{\"" + data.Key + "\"},";
                Jmp_Script += "\n";
                Jmp_Script += " \"Distrib Histogram\",";
                Jmp_Script += "\n";
                Jmp_Script += "FrameBox,";
                Jmp_Script += "\n";
                Jmp_Script += "{";
                Jmp_Script += "DispatchSeg(LabelSeg(1), { Font( \"Segoe UI\"" + ", 7,  \"Plain\")} ),";
                Jmp_Script += "\n";
                Jmp_Script += "DispatchSeg(LabelSeg(2), {  Font( \"Segoe UI\"" + ", 7,  \"Plain\")} )}";
                Jmp_Script += "\n";
                Jmp_Script += ")";
                Jmp_Script += ")";
                Jmp_Script += ")";
            }
            else
            {



                foreach (KeyValuePair<int, Dictionary<int, string>> Data in OrderbySequence)
                {
                    if (Data.Value[999] == split[1])
                    {

                        if (Data.Value.Keys.Contains(777))
                        {
                            Spec_Ig = Data.Value[777];
                            break;
                        }
                    }

                }
                string SpecMin = "";
                string SpecHigh = "";

                if (NPI_enable)
                {
                    SpecMin = data.Value.Broadcom_Spec_Min;
                    SpecHigh = data.Value.Broadcom_Spec_Max;
                }
                else if (Customer_enable)
                {
                    SpecMin = data.Value.Apple_Spec_Min;
                    SpecHigh = data.Value.Apple_Spec_Max;
                }

                double average = data.Value.data.Average();

                Array.Sort(data.Value.data);
                double Median = 0f;

                if (data.Value.data.Length % 2 == 0)
                {
                    double i = data.Value.data[((data.Value.data.Length / 2) - 1)];
                    double j = data.Value.data[(data.Value.data.Length) / 2];
                    double Ave = (i + j) / 2;
                    Median = Ave;
                }
                else
                {

                    int GetMedian_i = (data.Value.data.Length) / 2;
                    Median = data.Value.data[GetMedian_i];
                }

                double minusSquareSummary = 0.0;

                foreach (double source in data.Value.data)
                {
                    minusSquareSummary += (source - average) * (source - average);
                }

                double stdev = Math.Sqrt(minusSquareSummary / (data.Value.data.Length - 1));

                var chart = new System.Web.UI.DataVisualization.Charting.Chart();
                double result = chart.DataManipulator.Statistics.InverseTDistribution(.05, data.Value.data.Length - 1);

                double Confidence_Interval = result * (stdev / Math.Sqrt(data.Value.data.Length));

                double Min95per = average - Confidence_Interval;
                double Max95per = average + Confidence_Interval;

                if (CPK_enable)
                {
                    SpecMin = Convert.ToString(average - CPK_Value * 3 * stdev);
                    SpecHigh = Convert.ToString(average + CPK_Value * 3 * stdev);
                }

                if (Spec_Ig.ToUpper() == "MIN")
                {
                    if (Convert.ToDouble(SpecMin) <= -999 && Convert.ToDouble(SpecHigh) >= 999)
                    {

                    }
                    else
                    {
                        Jmp_Script += " Capability Analysis(USL(" + SpecHigh + "))";
                        Jmp_Script += ",";
                    }


                    Jmp_Script += "SendToReport(";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";
                    Jmp_Script += "{ \"" + data.Key + "\" },";
                    Jmp_Script += "\n";
                    Jmp_Script += "" + "1" + ",";
                    Jmp_Script += "),";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";

                    Jmp_Script += "{\"" + data.Key + "\"},";
                    Jmp_Script += "\n";

                    Jmp_Script += " \"Distrib Histogram\",";
                    Jmp_Script += "\n";
                    Jmp_Script += "FrameBox,";
                    Jmp_Script += "\n";
                    Jmp_Script += "{";
                    Jmp_Script += "DispatchSeg(LabelSeg(1), { Font( \"Segoe UI\"" + ", 7,  \"Plain\")} ),";
                    Jmp_Script += "\n";
                    Jmp_Script += "DispatchSeg(LabelSeg(2), {  Font( \"Segoe UI\"" + ", 7,  \"Plain\")} )}";
                    Jmp_Script += "\n";
                    Jmp_Script += ")";
                    Jmp_Script += "\n";
                    Jmp_Script += ")";
                    Jmp_Script += "\n";
                    Jmp_Script += ")";
                }
                else if (Spec_Ig.ToUpper() == "MAX")
                {
                    if (Convert.ToDouble(SpecMin) <= -999 && Convert.ToDouble(SpecHigh) >= 999)
                    {

                    }
                    else
                    {
                        Jmp_Script += " Capability Analysis(LSL(" + SpecMin + "))";
                        Jmp_Script += ",";
                    }



                    Jmp_Script += "\n";
                    Jmp_Script += "SendToReport(";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";
                    Jmp_Script += "{ \"" + data.Key + "\" },";
                    Jmp_Script += "" + "1" + ",";
                    Jmp_Script += "),";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";
                    Jmp_Script += "{\"" + data.Key + "\"},";
                    Jmp_Script += "\n";
                    Jmp_Script += " \"Distrib Histogram\",";
                    Jmp_Script += "\n";
                    Jmp_Script += "FrameBox,";
                    Jmp_Script += "\n";
                    Jmp_Script += "{";
                    Jmp_Script += "DispatchSeg(LabelSeg(1), { Font( \"Segoe UI\"" + ", 7,  \"Plain\")} ),";
                    Jmp_Script += "\n";
                    Jmp_Script += "DispatchSeg(LabelSeg(2), {  Font( \"Segoe UI\"" + ", 7,  \"Plain\")} )}";
                    Jmp_Script += "\n";
                    Jmp_Script += ")";
                    Jmp_Script += ")";
                    Jmp_Script += ")";
                }


                else if (Spec_Ig.ToUpper() == "MIN>MAX" || Spec_Ig.ToUpper() == "MAX>MIN")
                {
                    if (Convert.ToDouble(SpecMin) <= -999 && Convert.ToDouble(SpecHigh) >= 999)
                    {

                    }
                    else
                    {
                        //   Jmp_Script += " Capability Analysis(LSL(" + SpecMin + "), USL(" + SpecHigh + "))";
                        //   Jmp_Script += ",";
                    }

                    Jmp_Script += "SendToReport(";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";
                    Jmp_Script += "{ \"" + data.Key + "\" },";
                    Jmp_Script += "\n";
                    Jmp_Script += "" + "1" + ",";
                    Jmp_Script += "ScaleBox,";
                    Jmp_Script += "\n";
                    Jmp_Script += "{ Min(" + Min95per + "), Max(" + Max95per + ")}";
                    Jmp_Script += "\n";
                    Jmp_Script += "),";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";

                    Jmp_Script += "{\"" + data.Key + "\"},";
                    Jmp_Script += "\n";
                    Jmp_Script += " \"Distrib Histogram\",";
                    Jmp_Script += "\n";
                    Jmp_Script += "FrameBox,";
                    Jmp_Script += "\n";
                    Jmp_Script += "{";
                    Jmp_Script += "DispatchSeg(LabelSeg(1), { Font( \"Segoe UI\"" + ", 7,  \"Plain\")} ),";
                    Jmp_Script += "\n";
                    Jmp_Script += "DispatchSeg(LabelSeg(2), {  Font( \"Segoe UI\"" + ", 7,  \"Plain\")} )}";
                    Jmp_Script += "\n";
                    Jmp_Script += ")";
                    Jmp_Script += ")";
                    Jmp_Script += ")";
                }

                else if (Spec_Ig.ToUpper() == "")
                {
                    if (Convert.ToDouble(SpecMin) <= -999 && Convert.ToDouble(SpecHigh) >= 999)
                    {

                    }
                    else
                    {
                        Jmp_Script += " Capability Analysis(LSL(" + SpecMin + "), USL(" + SpecHigh + "))";
                        Jmp_Script += ",";
                    }



                    Jmp_Script += "\n";
                    Jmp_Script += "SendToReport(";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";
                    Jmp_Script += "{ \"" + data.Key + "\" },";
                    Jmp_Script += "" + "1" + ",";
                    Jmp_Script += "\n";
                    Jmp_Script += "ScaleBox,";
                    Jmp_Script += "\n";
                    Jmp_Script += "{ Min(" + Min95per + "), Max(" + Max95per + ")}";
                    Jmp_Script += "),";
                    Jmp_Script += "\n";
                    Jmp_Script += "Dispatch(";
                    Jmp_Script += "\n";

                    Jmp_Script += "{\"" + data.Value.Parameter + "\"},";
                    Jmp_Script += "\n";
                    Jmp_Script += " \"Distrib Histogram\",";
                    Jmp_Script += "\n";
                    Jmp_Script += "FrameBox,";
                    Jmp_Script += "\n";
                    Jmp_Script += "{";
                    Jmp_Script += "DispatchSeg(LabelSeg(1), { Font( \"Segoe UI\"" + ", 7,  \"Plain\")} ),";
                    Jmp_Script += "\n";
                    Jmp_Script += "DispatchSeg(LabelSeg(2), {  Font( \"Segoe UI\"" + ", 7,  \"Plain\")} )}";
                    Jmp_Script += "\n";
                    Jmp_Script += ")";
                    Jmp_Script += ")";
                    Jmp_Script += ")";
                }

                Jmp_Script += "\n";
            }

        }
        public void SendtoReport()
        {
            Jmp_Script += "	SendToReport(";
            Jmp_Script += "\n";
        }
        public void Dispatch()
        {
            Jmp_Script += "Dispatch( { }, \"1\", ScaleBox,";
            Jmp_Script += "\n";
        }
        public void Spec(string CustomerSpecMin, string CustomerSpecHigh, string BroadcomSpecMin, string BroadcomSpecHigh)
        {
            Jmp_Script += "{Add Ref Line( " + BroadcomSpecMin + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + BroadcomSpecHigh + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecMin + ", \"Solid\", \"Red\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecHigh + ", \"Solid\", \"Red\", \"\", 3 )}),";
            Jmp_Script += "\n";

            Jmp_Script += "Dispatch( { }, \"Oneway Plot\", FrameBox,";
            Jmp_Script += "\n";
            Jmp_Script += "DispatchSeg( Box Plot Seg(1),{ Box Type(\"Outlier\"), Line Color(\"Red\")}";
            Jmp_Script += "\n";

        }
        public void End(bool Flag)
        {
            if (Flag)
            {
                Jmp_Script += ");";
            }
            else
            {
                Jmp_Script += ",";
            }


        }

    }

    public class Distribution_By_X
    {
        public string Jmp_Script;

        public Distribution_By_X()
        {
            Jmp_Script = "";
        }
        public void Dist()
        {
            Jmp_Script = "Distribution(";
            Jmp_Script += "\n";
            Jmp_Script += "Stack(1),";
            Jmp_Script += "\n";
        }

        public void Continuous_Distribution(KeyValuePair<string, CSV_Class.For_Box> Data , Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value)
        {
            string Spec_Ig = "";
            string split = Data.Key.Split('_')[1];
            foreach (KeyValuePair<int, Dictionary<int, string>> Datas in OrderbySequence)
            {
                if (Datas.Value[999] == split)
                {

                    if (Datas.Value.Keys.Contains(777))
                    {
                        Spec_Ig = Datas.Value[777];
                        break;
                    }
                }

            }
            string SpecMin = "";
            string SpecHigh = "";

            if (NPI_enable)
            {
                SpecMin = Data.Value.Broadcom_Spec_Min;
                SpecHigh = Data.Value.Broadcom_Spec_Max;
            }
            else if (Customer_enable)
            {
                SpecMin = Data.Value.Apple_Spec_Min;
                SpecHigh = Data.Value.Apple_Spec_Max;
            }

            double average = Data.Value.data.Average();

            Array.Sort(Data.Value.data);
            double Median = 0f;

            if (Data.Value.data.Length % 2 == 0)
            {
                double i = Data.Value.data[((Data.Value.data.Length / 2) - 1)];
                double j = Data.Value.data[(Data.Value.data.Length) / 2];
                double Ave = (i + j) / 2;
                Median = Ave;
            }
            else
            {

                int GetMedian_i = (Data.Value.data.Length) / 2;
                Median = Data.Value.data[GetMedian_i];
            }

            double minusSquareSummary = 0.0;

            foreach (double source in Data.Value.data)
            {
                minusSquareSummary += (source - average) * (source - average);
            }

            double stdev = Math.Sqrt(minusSquareSummary / (Data.Value.data.Length - 1));

            var chart = new System.Web.UI.DataVisualization.Charting.Chart();
            double result = chart.DataManipulator.Statistics.InverseTDistribution(.05, Data.Value.data.Length - 1);

            double Confidence_Interval = result * (stdev / Math.Sqrt(Data.Value.data.Length));

            double Min95per = average - Confidence_Interval;
            double Max95per = average + Confidence_Interval;

            if (CPK_enable)
            {
                SpecMin = Convert.ToString(average - CPK_Value * 3 * stdev);
                SpecHigh = Convert.ToString(average + CPK_Value * 3 * stdev);
            }

            Jmp_Script += "Continuous Distribution(";
            Jmp_Script += "\n";
            Jmp_Script += "Column( :NAME(\"" + Data.Key + "\")),";
            Jmp_Script += "\n";
            Jmp_Script += "Horizontal Layout(1),";
            Jmp_Script += "\n";
            Jmp_Script += "Vertical(0)";
            Jmp_Script += "\n";

            if(SpecHigh == "" && SpecMin == "")
            {
                Jmp_Script += "),";
            }

            else if (Spec_Ig.ToUpper() == "MIN")
            {
                Jmp_Script += ",Capability Analysis(USL(" + SpecHigh + ")))" + ",";
   
            }
            else if (Spec_Ig.ToUpper() == "MAX")
            {
                Jmp_Script += ",Capability Analysis(LSL(" + SpecMin + ")))" + ",";

            }


            else if (Spec_Ig.ToUpper() == "MIN>MAX" || Spec_Ig.ToUpper() == "MAX>MIN")
            {
                Jmp_Script += "),";

            }

            else if (Spec_Ig.ToUpper() == "")
            {

                Jmp_Script += ",Capability Analysis(LSL(" + SpecMin + "), USL(" + SpecHigh + "))),";

            }
            Jmp_Script += "\n";
        }

        public void By(Dictionary<string, CSV_Class.For_Box> Data, int Count)
        {
            int i = 0;

            Jmp_Script += "By(";

            foreach (KeyValuePair<string, CSV_Class.For_Box> _D in Data)
            {

                Jmp_Script += ":" + _D.Key;

                if (i == Data.Count - 1)
                {
                    Jmp_Script += "),";
                }
                else
                {
                    Jmp_Script += ",";
                }
                i++;
            }

        }

        public void Capability_Analysis(string Para_Test, Dictionary<String, CSV_Class.For_Box> data, Dictionary<int, Dictionary<int, string>> OrderbySequence, bool Customer_enable, bool NPI_enable, bool CPK_enable, double CPK_Value, bool Flag)
        {



            int k = 0;

            string[] dummy = Para_Test.Split(',');

            Jmp_Script += "\n";

            foreach (KeyValuePair<String, CSV_Class.For_Box> _D in data)
            {
                string Name = "";

                Jmp_Script += "\n";
                Jmp_Script += "SendToByGroup( {";
                Jmp_Script += "\n";

                for (int h = 0; h < dummy.Length; h++)
                {

                    if(dummy.Length == 2)
                    {
                        if (h != dummy.Length - 2)
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" ,";
                            Name += dummy[h] + " = " + dummy[h + 1] + ",";
                        }
                        else
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" },";
                            Name += dummy[h] + " = " + dummy[h + 1] + "\"";
                        }
                    }
                    else
                    {
                        if (h != dummy.Length - 2)
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" ,";
                            Name += dummy[h] + " = " + dummy[h + 1] + ",";
                        }
                        else
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" },";
                            Name += dummy[h] + " = " + dummy[h + 1] + "\"";
                        }
                    }

                    h++;
                }

                Jmp_Script += "SendToReport(";
                Jmp_Script += "\n";
                Jmp_Script += "Dispatch(";
                Jmp_Script += "\n";
                Jmp_Script += "{\"Distributions " + Name + ",";
                Jmp_Script += "\n";

                Jmp_Script += "\"" + _D.Key + "\" },";
                Jmp_Script += "\n";
                Jmp_Script += "\"1 \",";

                Jmp_Script += "\n";
                Jmp_Script += "Dispatch(";
                Jmp_Script += "\n";

                Jmp_Script += "{\"" + _D.Value.Parameter + "\"},";
                Jmp_Script += "\n";

                Jmp_Script += " \"Distrib Histogram\",";
                Jmp_Script += "\n";
                Jmp_Script += "FrameBox,";
                Jmp_Script += "\n";
                Jmp_Script += "{";
                Jmp_Script += "DispatchSeg(LabelSeg(1), { Font( \"Segoe UI\"" + ", 7,  \"Plain\")} ),";
                Jmp_Script += "\n";
                Jmp_Script += "DispatchSeg(LabelSeg(2), {  Font( \"Segoe UI\"" + ", 7,  \"Plain\")} )}";
                Jmp_Script += "\n";
                Jmp_Script += ")";
                Jmp_Script += "\n";
                Jmp_Script += ")";
                Jmp_Script += "\n";
                Jmp_Script += ")";


                if (Flag && k == data.Count - 1)
                {
                    Jmp_Script += "));";
                }
                else
                {
                    Jmp_Script += "),";
                }
                k++;

            }




        }


        public void SendtoReport()
        {
            Jmp_Script += "	SendToReport(";
            Jmp_Script += "\n";
        }
        public void Dispatch()
        {
            Jmp_Script += "Dispatch( { }, \"1\", ScaleBox,";
            Jmp_Script += "\n";
        }
        public void Spec(string CustomerSpecMin, string CustomerSpecHigh, string BroadcomSpecMin, string BroadcomSpecHigh)
        {
            Jmp_Script += "{Add Ref Line( " + BroadcomSpecMin + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + BroadcomSpecHigh + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecMin + ", \"Solid\", \"Red\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecHigh + ", \"Solid\", \"Red\", \"\", 3 )}),";
            Jmp_Script += "\n";

            Jmp_Script += "Dispatch( { }, \"Oneway Plot\", FrameBox,";
            Jmp_Script += "\n";
            Jmp_Script += "DispatchSeg( Box Plot Seg(1),{ Box Type(\"Outlier\"), Line Color(\"Red\")}";
            Jmp_Script += "\n";

        }
        public void End(bool Flag)
        {
            if (Flag)
            {
                Jmp_Script += ");";
            }
            else
            {
                Jmp_Script += ",";
            }


        }

    }

    public class FitYbyX
    {
        public string Jmp_Script;

        public FitYbyX()
        {
            Jmp_Script = "";
        }
        public void Fit_Group()
        {
            Jmp_Script = "Fit Group(";
            Jmp_Script += "\n";
        }

        public void Oneway()
        {
            Jmp_Script += "Oneway(";
            Jmp_Script += "\n";
        }

        public void Y(string Name)
        {
            Jmp_Script += "Y( :Name( \"" + Name + "\")),";
            Jmp_Script += "\n";

        }
        public void X(string[] Parameter)
        {

            Jmp_Script += "X( " + Parameter[0] + "),";
            Jmp_Script += "\n";
        }

        public void Quantiles()
        {
            Jmp_Script += "Quantiles( 1 ),";
            Jmp_Script += "\n";
        }
        public void BoxPlots()
        {
            Jmp_Script += "Box Plots( 0 ),";
            Jmp_Script += "\n";
        }
        public void MeansandStdDev()
        {
            Jmp_Script += "Means and Std Dev( 1 ),";
            Jmp_Script += "\n";
        }
        public void StddevLines()
        {
            Jmp_Script += "Std Dev Lines( 1 ),";
            Jmp_Script += "\n";
        }
        public void GrandMean()
        {
            Jmp_Script += "Grand Mean( 0 ),";
            Jmp_Script += "\n";
        }
        public void SendtoReport()
        {
            Jmp_Script += "	SendToReport(";
            Jmp_Script += "\n";
        }
        public void Dispatch()
        {
            Jmp_Script += "Dispatch( { }, \"1\", ScaleBox,";
            Jmp_Script += "\n";
        }
        public void Spec(string CustomerSpecMin, string CustomerSpecHigh, string BroadcomSpecMin, string BroadcomSpecHigh)
        {
            Jmp_Script += "{Add Ref Line( " + BroadcomSpecMin + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + BroadcomSpecHigh + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecMin + ", \"Solid\", \"Red\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecHigh + ", \"Solid\", \"Red\", \"\", 3 )}),";
            Jmp_Script += "\n";

            Jmp_Script += "Dispatch( { }, \"Oneway Plot\", FrameBox,";
            Jmp_Script += "\n";
            Jmp_Script += "DispatchSeg( Box Plot Seg(1),{ Box Type(\"Outlier\"), Line Color(\"Red\")}";
            Jmp_Script += "\n";

        }
        public void End(bool Flag)
        {
            if(Flag)
            {
                Jmp_Script += "))),));";
            }
            else
            {
                Jmp_Script += ")))),";
            }
      

        }

    }
    public class FitYbyXs
    {
        public string Jmp_Script;

        public FitYbyXs()
        {
            Jmp_Script = "";
        }
        public void Fit_Group()
        {
            Jmp_Script = "Fit Group(";
            Jmp_Script += "\n";
        }

        public void Oneway()
        {
            Jmp_Script += "Bivariate(";
            Jmp_Script += "\n";
        }

        public void Y(string Name)
        {
            Jmp_Script += "Y( :Name( \"" + Name + "\")),";
            Jmp_Script += "\n";

        }
        public void X(string[] Parameter)
        {

            Jmp_Script += "X( " + Parameter[0] + "),";
            Jmp_Script += "\n";
        }

      
        public void SendtoReport()
        {
            Jmp_Script += "	SendToReport(";
            Jmp_Script += "\n";
        }
        public void Dispatch(int Count)
        {
            Jmp_Script += "Dispatch( { }, \"" + Count + "\", ScaleBox,";
            Jmp_Script += "\n";
        }
        public void Spec(string CustomerSpecMin, string CustomerSpecHigh, string BroadcomSpecMin, string BroadcomSpecHigh)
        {
            Jmp_Script += "{Add Ref Line( " + BroadcomSpecMin + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + BroadcomSpecHigh + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecMin + ", \"Solid\", \"Red\", \"\", 3 ),";
            Jmp_Script += "\n";
            Jmp_Script += "Add Ref Line( " + CustomerSpecHigh + ", \"Solid\", \"Red\", \"\", 3 )}),";
            Jmp_Script += "\n";

            //Jmp_Script += "Dispatch( { }, \"Oneway Plot\", FrameBox,";
            //Jmp_Script += "\n";
            //Jmp_Script += "DispatchSeg( Box Plot Seg(1),{ Box Type(\"Outlier\"), Line Color(\"Red\")}";
            //Jmp_Script += "\n";

        }
        public void End(bool Flag)
        {
            if (Flag)
            {
                Jmp_Script += ")));";
            }
            else
            {
                Jmp_Script += ")),";
            }


        }

    }


    public class BoxPlots
    {
        public string Jmp_Script;

        public BoxPlots()
        {
            Jmp_Script = "";
        }
        public void DT_Open(string Path)
        {
            Jmp_Script += "dt = current data table();";
            Jmp_Script += "dist = dt << Variability Chart(";
        }
        public void Y(string Parameter)
        {
            Jmp_Script += "Y( :Name( \"" + Parameter + "\")),";
        }
        public void X(KeyValuePair<int, Dictionary<int, string>> Data)
        {
            Jmp_Script += "x(:";
            int i = 0;
            int Count = Data.Value.Count - 2;
            foreach (KeyValuePair<int, string> S in Data.Value)
            {
                if (Count == i)
                {
                    Jmp_Script += S.Value + "),";
                    break;
                }
                if (Count - 1 == i)
                {
                    Jmp_Script += S.Value + ", :";
                }
                else if (i != 0)
                {
                    Jmp_Script += S.Value + ", :";
                }
                i++;

            }

            Jmp_Script += "\n";
        }
        public void Setting(string Parameter)
        {
            Jmp_Script += "Max Iter( 100 )," +
                           "Conv Limit(0.00000001)," +
                           "Number Integration Abscissas(128)," +
                           "Number Function Evals(65536)," +
                           "Analysis Type( \"Choose best analysis (EMS REML Bayesian) \")," +
                           " Std Dev Chart(0)," +
                           "Show Box Plots(1),";
            Jmp_Script += "\n";
        }
        public void SendReport()
        {
            Jmp_Script += "SendToReport(";
            Jmp_Script += "\n";
        }
        public void Dispatch_Set_MinMax(string Parameter, Dictionary<string, CSV_Class.For_Box> Data_Test, KeyValuePair<int, Dictionary<int, string>> Data)
        {
            string Ignore_Spec = "";
            if (Data.Value.Keys.Contains(777))
            {
                Ignore_Spec = Data.Value[777];
            }
         

            var varList = Data_Test.Keys.ToList();
            varList.Sort();

            int For_Spec_Line = 0;
            int kkk = 0;


            List<double> Datas = new List<double>();
            List<double> Spec_Min = new List<double>();
            List<double> Spec_Max = new List<double>();

            double Broadcom_Min = 0f;
            double Broadcom_Max = 0f;
            double Data_Min_Value = 0f;
            double Data_Max_Value = 0f;


            foreach (string item in Data_Test.Keys)
            {
                CSV_Class.For_Box Test = Data_Test[varList[For_Spec_Line].ToString()];


                    Broadcom_Min = Convert.ToDouble(Test.Broadcom_Spec_Min);
                    Broadcom_Max = Convert.ToDouble(Test.Broadcom_Spec_Max);

                    Data_Min_Value = Test.data.Min();
                    Data_Max_Value = Test.data.Max();

                    Datas.Add(Data_Min_Value);
                    Datas.Add(Data_Min_Value);

                    Datas.Add(Broadcom_Min);
                    Datas.Add(Broadcom_Max);

                    Spec_Min.Add(Convert.ToDouble(Test.Apple_Spec_Min));
                    Spec_Min.Add(Convert.ToDouble(Data_Min_Value));

                    Spec_Max.Add(Convert.ToDouble(Test.Apple_Spec_Max));
                    Spec_Max.Add(Convert.ToDouble(Data_Max_Value));
                


                For_Spec_Line++;
            }

            double Data_Min = Datas.Min();
            double Data_Max = Datas.Max();



            Spec_Min = Spec_Min.Distinct().ToList();
            Spec_Max = Spec_Max.Distinct().ToList();

            Spec_Min.Remove(-999);
            Spec_Max.Remove(999);


            if (Ignore_Spec == "")
            {
                Datas = Datas.Concat(Spec_Min).ToList();
                Datas = Datas.Concat(Spec_Max).ToList();

                if (Datas.Max() < 0)
                {
                    Data_Max = Datas.Max() * 0.95;
                }
                else
                {
                    Data_Max = Datas.Max() * 1.05;
                }


                if (Datas.Min() < 0)
                {
                    Data_Min = Datas.Min() * 1.05;
                }
                else
                {
                    Data_Min = Datas.Min() * 0.95;
                }

            }
            else if (Ignore_Spec.ToUpper().Trim() == "MIN")
            {
                Datas = Datas.Concat(Spec_Max).ToList();

                if (Datas.Max() < 0)
                {
                    Data_Max = Datas.Max() * 0.95;
                }
                else
                {
                    Data_Max = Datas.Max() * 1.05;
                }

                if (Datas.Min() < 0)
                {
                    Data_Min = Datas.Min() * 1.05;
                }
                else
                {
                    Data_Min = Datas.Min() * 0.95;
                }
            }
            else if (Ignore_Spec.ToUpper().Trim() == "MAX")
            {
                Datas = Datas.Concat(Spec_Min).ToList();

                if (Datas.Max() < 0)
                {
                    Data_Max = Datas.Max() * 0.95;
                }
                else
                {
                    Data_Max = Datas.Max() * 1.05;
                }

                if (Datas.Min() < 0)
                {
                    Data_Min = Datas.Min() * 1.05;
                }
                else
                {
                    Data_Min = Datas.Min() * 0.95;
                }
            }
            else if (Ignore_Spec.ToUpper().Trim() == "MAX/MIN" || Ignore_Spec.ToUpper().Trim() == "MIN/MAX")
            {

            }

            List<double> Edit_Data = new List<double>();
            Edit_Data.Add(Data_Max);
            Edit_Data.Add(Data_Min);

            var Edit = Edit_Data.OrderByDescending(k => k).ToArray();


            if (Ignore_Spec == "")
            {
                Jmp_Script += "Dispatch({\"Variability chart for " + Parameter + "\"},\"2\", ScaleBox, { Min(" + Edit[1] + " ), Max(" + Edit[0] + "), Minor Ticks(1)}),";
                Jmp_Script += "\n";
            }

            else if (Ignore_Spec.ToUpper().Trim() == "MIN")
            {
                Jmp_Script += "Dispatch({\"Variability chart for " + Parameter + "\"},\"2\", ScaleBox, {Max(" + Edit[0] + "), Minor Ticks(1)}),";
                Jmp_Script += "\n";
            }
            else if (Ignore_Spec.ToUpper().Trim() == "MAX")
            {
                Jmp_Script += "Dispatch({\"Variability chart for " + Parameter + "\"},\"2\", ScaleBox, { Min(" + Edit[1] + " ), Minor Ticks(1)}),";
                Jmp_Script += "\n";
            }
            else if (Ignore_Spec.ToUpper().Trim() == "MAX/MIN" || Ignore_Spec.ToUpper().Trim() == "MIN/MAX")
            {
              //  Jmp_Script += "Dispatch({\"Variability chart for " + Parameter + "\"}, \"Variability chart\",";

            }

        }
        public void Dispatch_Set_SpecLine(string Parameter , string[] List, Dictionary<string, CSV_Class.For_Box> Data_Test)
        {

            int Line = 0;

            Jmp_Script += "Dispatch({\"Variability chart for " + Parameter + "\"},\"Variability chart\", Framebox,{frame size ( 1000,300), Add Graphics Script(";
            Jmp_Script += "\n";
            foreach (string item in Data_Test.Keys)
            {
                CSV_Class.For_Box Test = Data_Test[List[Line].ToString()];


                Jmp_Script += "Line style(\"solid\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen color(\"Blue\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen size(3);";
                Jmp_Script += "\n";
                Jmp_Script += "Line ({ " + Line + " , " + Test.Broadcom_Spec_Min + " },{" + (Line + 1) + " , " + Test.Broadcom_Spec_Min + "});";
                Jmp_Script += "\n";
                Jmp_Script += "Line style(\"solid\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen color(\"Blue\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen size(3);";
                Jmp_Script += "\n";
                Jmp_Script += "Line ({ " + Line + " , " + Test.Broadcom_Spec_Max + " },{" + (Line + 1) + " , " + Test.Broadcom_Spec_Max + "});";
                Jmp_Script += "\n";
                Jmp_Script += "Line style(\"solid\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen color(\"Red\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen size(3);";
                Jmp_Script += "\n";
                Jmp_Script += "Line ({ " + Line + " , " + Test.Apple_Spec_Min + " },{" + (Line + 1) + " , " + Test.Apple_Spec_Min + "});";
                Jmp_Script += "\n";

                Jmp_Script += "Line style(\"solid\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen color(\"Red\");";
                Jmp_Script += "\n";
                Jmp_Script += "pen size(3);";
                Jmp_Script += "\n";
                Jmp_Script += "Line ({ " + Line + " , " + Test.Apple_Spec_Max + " },{" + (Line + 1) + " , " + Test.Apple_Spec_Max + "});";
                Jmp_Script += "\n";
                Line++;
            }
        }
        public void End()
        {
            Jmp_Script += ")}";
            Jmp_Script += ")));";

        }


        public void Spec(string CustomerSpecMin, string CustomerSpecHigh, string BroadcomSpecMin, string BroadcomSpecHigh)
        {
            Jmp_Script += "{Add Ref Line( " + BroadcomSpecMin + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "Add Ref Line( " + BroadcomSpecHigh + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "Add Ref Line( " + CustomerSpecMin + ", \"Solid\", \"Red\", \"\", 3 ),";
            Jmp_Script += "Add Ref Line( " + CustomerSpecHigh + ", \"Solid\", \"Red\", \"\", 3 )}),";

            Jmp_Script += "Dispatch( { }, \"Oneway Plot\", FrameBox,";
            Jmp_Script += "{DispatchSeg( Box Plot Seg(1),{ Box Type(\"Outlier\"), Line Color(\"Red\")}";

        }
  

    }

    public class BoxPlots_By
    {
        public string Jmp_Script;

        public BoxPlots_By()
        {
            Jmp_Script = "";
        }
        public void DT_Open(string Path)
        {
            Jmp_Script += "dt = current data table();";
            Jmp_Script += "dist = dt << Variability Chart(";
        }
        public void SendToByGroup(List<string>[] Para_Test, string[] Bynum_Split, string Key, int Count_Script)
        {
            int k = 0;
            for (int p = 0; p < Para_Test[Count_Script].Count; p++)
            {

                string[] dummy = Para_Test[Count_Script][p].Split(',');

                if (Bynum_Split.Length == 1)
                {
                    for (int h = 0; h < dummy.Length; h++)
                    {
                        var by1 = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[h]);


                        Jmp_Script += "SendToByGroup( { :" + by1 + " == \"" + dummy[h + 1] + "\"} ), Y ( :Name( \"" + Key + "\")),";
                        Jmp_Script += "\n";
                        h++;
                    }
                }
                else
                {
                    Jmp_Script += "SendToByGroup( {";
                    Jmp_Script += "\n";
                    for (int h = 0; h < dummy.Length; h++)
                    {
                       // var by1 = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[h]);

                        if (h != dummy.Length - 2)
                        //    if (h != dummy.Length - 1)
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" ,";
                        }
                        else 
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" }),";
                        }
                        h++;
                    }

                    Jmp_Script += " Y ( :Name( \"" + Key + "\")),";
                    Jmp_Script += "\n";
                }

                k++;
            }



        }

        public void X(KeyValuePair<int, Dictionary<int, string>> Data)
        {
            Jmp_Script += "x(:";
            Jmp_Script += "\n";
            int i = 0;
            int Count = Data.Value.Count - 3;
            foreach (KeyValuePair<int, string> S in Data.Value)
            {
                if (Count == i)
                {
                    Jmp_Script += S.Value + "),";
                    break;
                }
                if (Count - 1 == i)
                {
                    Jmp_Script += S.Value + ", :";
                }
                else if (i != 0)
                {
                    Jmp_Script += S.Value + ", :";
                }
                i++;

            }

            Jmp_Script += "\n";
        }
        public void Setting(string[] Bynum_Split)
        {

            string Parameter = "";
            for (int h = 0; h < Bynum_Split.Length; h++)
            {
                var by1 = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[h]);

                if (Bynum_Split.Length == 1)
                {

                    Parameter += ":" + by1;

                }
                else
                {
                    if (h != Bynum_Split.Length - 1)
                    {
                        Parameter += ":" + by1 + ",";
                    }
                    else
                    {
                        Parameter += ":" + by1;
                    }
                }

            }

            Jmp_Script += "Max Iter( 100 )," +
                           "Conv Limit(0.00000001)," +
                           "Number Integration Abscissas(128)," +
                           "Number Function Evals(65536)," +
                           "Analysis Type( \"Choose best analysis (EMS REML Bayesian) \")," +
                           " Std Dev Chart(0)," +
                           "Show Box Plots(1),";
            Jmp_Script += "\n";
            Jmp_Script += "By(" + Parameter + "),";
            Jmp_Script += "\n";
        }
        public void SendReport()
        {
            Jmp_Script += "SendToByReport(";
            Jmp_Script += "\n";
        }
        public void Dispatch_Set_MinMax(string Parameter,string[] Bynum_Split, List<string>[] Para_Test, Dictionary<string, CSV_Class.For_Box>[] Data_Test1, KeyValuePair<int, Dictionary<int, string>> Data, int Count_Script)
        {


            for (int p = 0; p < Data_Test1.Length; p++)
            {

                string Ignore_Spec = "";
                if (Data.Value.Keys.Contains(777))
                {
                    Ignore_Spec = Data.Value[777];
                }


                var varList_test = Data_Test1[p].Keys.ToList();
                varList_test.Sort();


                int For_Spec_Line = 0;
                int kkk = 0;


                List<double> Datas = new List<double>();
                List<double> Spec_Min = new List<double>();
                List<double> Spec_Max = new List<double>();


                Dictionary<string, CSV_Class.For_Box> items = Data_Test1[p];

                varList_test = items.Keys.ToList();
                varList_test.Sort();

                double Broadcom_Min = 0f;
                double Broadcom_Max = 0f;
                double Data_Min_Value = 0f;
                double Data_Max_Value = 0f;

                foreach (CSV_Class.For_Box item in items.Values)
                {

                    CSV_Class.For_Box Test = items[varList_test[For_Spec_Line].ToString()];



                    Broadcom_Min = Convert.ToDouble(Test.Broadcom_Spec_Min);
                    Broadcom_Max = Convert.ToDouble(Test.Broadcom_Spec_Max);

                    Data_Min_Value = Test.data.Min();
                    Data_Max_Value = Test.data.Max();

                    Datas.Add(Data_Min_Value);
                    Datas.Add(Data_Min_Value);

                    Datas.Add(Broadcom_Min);
                    Datas.Add(Broadcom_Max);

                    Spec_Min.Add(Convert.ToDouble(Test.Apple_Spec_Min));
                    Spec_Min.Add(Convert.ToDouble(Data_Min_Value));

                    Spec_Max.Add(Convert.ToDouble(Test.Apple_Spec_Max));
                    Spec_Max.Add(Convert.ToDouble(Data_Max_Value));


                    For_Spec_Line++;
                }



                double Data_Min = Datas.Min();
                double Data_Max = Datas.Max();

                Spec_Min = Spec_Min.Distinct().ToList();
                Spec_Max = Spec_Max.Distinct().ToList();

                Spec_Min.Remove(-999);
                Spec_Max.Remove(999);


                if (Ignore_Spec == "")
                {
                    Datas = Datas.Concat(Spec_Min).ToList();
                    Datas = Datas.Concat(Spec_Max).ToList();

                    if (Datas.Max() < 0)
                    {
                        Data_Max = Datas.Max() * 0.95;
                    }
                    else
                    {
                        Data_Max = Datas.Max() * 1.05;
                    }

                    if (Datas.Min() < 0)
                    {
                        Data_Min = Datas.Min() * 1.05;
                    }
                    else
                    {
                        Data_Min = Datas.Min() * 0.95;
                    }

                }
                else if (Ignore_Spec.ToUpper().Trim() == "MIN")
                {
                    Datas = Datas.Concat(Spec_Max).ToList();

                    if (Datas.Max() < 0)
                    {
                        Data_Max = Datas.Max() * 0.95;
                    }
                    else
                    {
                        Data_Max = Datas.Max() * 1.05;
                    }

                    if (Datas.Min() < 0)
                    {
                        Data_Min = Datas.Min() * 1.05;
                    }
                    else
                    {
                        Data_Min = Datas.Min() * 0.95;
                    }

                }
                else if (Ignore_Spec.ToUpper().Trim() == "MAX")
                {
                    Datas = Datas.Concat(Spec_Min).ToList();
                    if (Datas.Max() < 0)
                    {
                        Data_Max = Datas.Max() * 0.95;
                    }
                    else
                    {
                        Data_Max = Datas.Max() * 1.05;
                    }

                    if (Datas.Min() < 0)
                    {
                        Data_Min = Datas.Min() * 1.05;
                    }
                    else
                    {
                        Data_Min = Datas.Min() * 0.95;
                    }
                }
                else if (Ignore_Spec.ToUpper().Trim() == "MAX/MIN" || Ignore_Spec.ToUpper().Trim() == "MIN/MAX")
                {

                }

                List<double> Edit_Data = new List<double>();
                Edit_Data.Add(Data_Max);
                Edit_Data.Add(Data_Min);


                var Edit = Edit_Data.OrderByDescending(kh => kh).ToArray();

                For_Spec_Line = 0;
                kkk = 0;

                string[] dummy = Para_Test[Count_Script][p].Split(',');
                string by = "";

                Jmp_Script += "SendToByGroup( {";
                Jmp_Script += "\n";

                for (int h = 0; h < dummy.Length; h++)
                {
            

                    if (Bynum_Split.Length == 1)
                    {
                        var by1 = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[h]);

                        Jmp_Script += ":" + by1 + " == \"" + dummy[h + 1] + "\" },";


                    }
                    else
                    {


                        if (h != dummy.Length - 2)
                        //    if (h != dummy.Length - 1)
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" ,";
                        }
                        else
                        {
                            Jmp_Script += ":" + dummy[h] + " == \"" + dummy[h + 1] + "\" },";
                        }

                        Jmp_Script += "\n";
                    }

                    Jmp_Script += "\n";

                  //  by = Convert.ToString(by1);
                    h++;
                }

                Jmp_Script += "SendToReport(Dispatch({\"Variability Gauge ";

                string by_test = "";
                for (int h = 0; h < dummy.Length; h++)
                {
        
                    if (Bynum_Split.Length == 1)
                    {
                        var by1 = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[h]);
                        Jmp_Script += by1 + " = " + dummy[h + 1];
                        by_test = Convert.ToString(by1);

                    }
                    else
                    {
                        if (h != dummy.Length - 2)
                        {
                            Jmp_Script += dummy[h] + " = " + dummy[h + 1] + ",";
                         //   by_test = Convert.ToString(by1);
                        }
                        else
                        {
                            Jmp_Script += dummy[h] + " = " + dummy[h + 1] ;
                           // by_test = Convert.ToString(by1);
                        }
                    }

                    h++;
                }


                Jmp_Script += "\"";
                if (Ignore_Spec == "")
                {
                    Jmp_Script += ",\"Variability chart for " + Parameter + "\"},\"2\", ScaleBox, { Min(" + Edit[1] + " ), Max(" + Edit[0] + "), Minor Ticks(1)}),";
                    Jmp_Script += "\n";
                }

                else if (Ignore_Spec.ToUpper().Trim() == "MIN")
                {
                    Jmp_Script += ",\"Variability chart for " + Parameter + "\"},\"2\", ScaleBox, {Max(" + Edit[0] + "), Minor Ticks(1)}),";
                    Jmp_Script += "\n";
                }
                else if (Ignore_Spec.ToUpper().Trim() == "MAX")
                {
                    Jmp_Script += ",\"Variability chart for " + Parameter + "\"},\"2\", ScaleBox, { Min(" + Edit[1] + " ), Minor Ticks(1)}),";
                    Jmp_Script += "\n";
                }
                else if (Ignore_Spec.ToUpper().Trim() == "MAX/MIN" || Ignore_Spec.ToUpper().Trim() == "MIN/MAX")
                {
                    //  Jmp_Script += "Dispatch({\"Variability chart for " + Parameter + "\"}, \"Variability chart\",";

                }


                Jmp_Script += "Dispatch({\"Variability Gauge ";

                for (int h = 0; h < dummy.Length; h++)
                {
           

                    if (Bynum_Split.Length == 1)
                    {
                        var by1 = (BoxPlot)Enum.Parse(typeof(BoxPlot), Bynum_Split[h]);

                        Jmp_Script += by1 + " = " + dummy[h + 1];

                    }
                    else
                    {
                        if (h != dummy.Length - 2)
                        {
                            Jmp_Script += dummy[h] + " = " + dummy[h + 1] + ",";
                        }
                        else
                        {
                            Jmp_Script += dummy[h] + " = " + dummy[h + 1];
                        }
                    }
                    h++;

                }
                Jmp_Script += "\",\"Variability Chart for " + Data.Value[999] + "\"},";

                Jmp_Script += "\n";


                Jmp_Script += "\"Variability Chart\",";
                Jmp_Script += "\n";
                Jmp_Script += "FrameBox,";
                Jmp_Script += "\n";
                Jmp_Script += "{Frame Size(1000,300),";
                Jmp_Script += "\n";
                Jmp_Script += "Add Graphics Script(";
                Jmp_Script += "\n";



                For_Spec_Line = 0;
                int k = 0;

                foreach (CSV_Class.For_Box item in Data_Test1[p].Values)
                {
                    var varList = Data_Test1[p].Keys.ToList();
                    varList.Sort();

                    CSV_Class.For_Box Test = Data_Test1[p][varList[k].ToString()];


                    Jmp_Script += "Line style(\"solid\");";
                    Jmp_Script += "pen color(\"Blue\");";
                    Jmp_Script += "pen size(3);";
                    Jmp_Script += "Line ({ " + For_Spec_Line + " , " + Test.Broadcom_Spec_Min + " },{" + (For_Spec_Line + 1) + " , " + Test.Broadcom_Spec_Min + "});";

                    Jmp_Script += "\n";
                    Jmp_Script += "Line style(\"solid\");";
                    Jmp_Script += "pen color(\"Blue\");";
                    Jmp_Script += "pen size(3);";
                    Jmp_Script += "Line ({ " + For_Spec_Line + " , " + Test.Broadcom_Spec_Max + " },{" + (For_Spec_Line + 1) + " , " + Test.Broadcom_Spec_Max + "});";

                    Jmp_Script += "\n";
                    Jmp_Script += "Line style(\"solid\");";
                    Jmp_Script += "pen color(\"Red\");";
                    Jmp_Script += "pen size(3);";
                    Jmp_Script += "Line ({ " + For_Spec_Line + " , " + Test.Apple_Spec_Min + " },{" + (For_Spec_Line + 1) + " , " + Test.Apple_Spec_Min + "});";

                    Jmp_Script += "\n";
                    Jmp_Script += "Line style(\"solid\");";
                    Jmp_Script += "pen color(\"Red\");";
                    Jmp_Script += "pen size(3);";
                    Jmp_Script += "Line ({ " + For_Spec_Line + " , " + Test.Apple_Spec_Max + " },{" + (For_Spec_Line + 1) + " , " + Test.Apple_Spec_Max + "});";
                    Jmp_Script += "\n";
                    For_Spec_Line++;
                    k++;

                }
                if ( p == Data_Test1.Length - 1)
                {
                    Jmp_Script += ")}))),";
                    Jmp_Script += "\n";
                }
                else
                {
                    Jmp_Script += ")}))),";
                    Jmp_Script += "\n";
                }
            
            }
   
        }
        public void Dispatch_Set_SpecLine(string Parameter, string[] List, Dictionary<string, CSV_Class.For_Box>[] Data_Test1)
        {

         
        }
        public void End()
        {
            Jmp_Script += ");";

        }

        public void Spec(string CustomerSpecMin, string CustomerSpecHigh, string BroadcomSpecMin, string BroadcomSpecHigh)
        {
            Jmp_Script += "{Add Ref Line( " + BroadcomSpecMin + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "Add Ref Line( " + BroadcomSpecHigh + ", \"Solid\", \"Blue\", \"\", 3 ),";
            Jmp_Script += "Add Ref Line( " + CustomerSpecMin + ", \"Solid\", \"Red\", \"\", 3 ),";
            Jmp_Script += "Add Ref Line( " + CustomerSpecHigh + ", \"Solid\", \"Red\", \"\", 3 )}),";

            Jmp_Script += "Dispatch( { }, \"Oneway Plot\", FrameBox,";
            Jmp_Script += "{DispatchSeg( Box Plot Seg(1),{ Box Type(\"Outlier\"), Line Color(\"Red\")}";

        }


    }

    public enum BoxPlot
    {
        Label,
        Identifier = 0,
        Parameter = 1,
        Measuer = 2,
        Band = 3,
        Pmode = 4,
        Modulation = 5,
        Waveform = 6,
        Power_Identifier = 7,
        Pout = 8,
        Frequency = 9,
        Vcc = 10,
        Vdd = 11,
        DAC1 = 12,
        DAC2 = 13,
        TX = 14,
        ANT = 15,
        RX = 16,
        Extra = 17,
        Note1 = 18,
        SpecNumber = 19,
        Site = 20,
        Lot = 21,
        Wafer = 22

    }

}
