using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
namespace CSV_Class
{
    public class CSV
    {
        public class FCM_Automation_CSV : INT
        {
            public FileStream OepnSepcFS { get; set; }
            public FileStream OepnSepcFS2 { get; set; }
            public StreamReader StreamReader { get; set; }
            public StreamReader StreamReader2 { get; set; }
            public StreamWriter StreamWrite { get; set; }
            public string[] Get_String { get; set; }
            public int TheFirst_trashes { get; set; }
            public int TheEnd_trashes { get; set; }
            public void Write_Open(string FilePath)
            {
                StreamWrite = new StreamWriter(FilePath, false, Encoding.Default);
            }
            public void Read_Open(string Filename)
            {
                OepnSepcFS = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader = new StreamReader(OepnSepcFS);
            }
            public void Read_Open2(string Filename)
            {
                OepnSepcFS2 = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader2 = new StreamReader(OepnSepcFS2);
            }
            public void Write_Close()
            {
                StreamWrite.Close();
            }
            public void Read_Close()
            {
                StreamReader.Close();
            }
            public void Read2_Close()
            {
                StreamReader2.Close();
            }
            public string[] Read()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');
            }
            public string[] Read2()
            {
                return Get_String = StreamReader2.ReadLine().Trim().Split(',');
            }
            public string Read_Cloth_Spec()
            {
                string Get_Strings = StreamReader.ReadLine().Trim();
                return Get_Strings;
            }
            public string[] Read_Test()
            {
                //return Get_String = StreamReader.ReadLine().Trim().Split(',');
                return null;
            }

            public void Write(Dictionary<string, For_Box> Data)
            {
                StreamWrite.WriteLine("");
            }
            public void Write(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
            public void Write(string Parameter, string data)
            {
                StreamWrite.WriteLine(Parameter + "\t" + data);
            }
            public void Write(string Parameter, object[] id, object[] data)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0) StreamWrite.WriteLine(Parameter);
                    else
                    {
                        StreamWrite.WriteLine(data[i - 1]);
                    }
                }

            }
            public void Write(string Parameter, object[] id, object[] data, object[] Lot, string Variation)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }
            }
            public void Write(string Parameter, string Data1, string Data2)
            {
                StreamWrite.WriteLine(Parameter + "," + Data1 + "," + Data2);
            }

            public void Write(string String, string Key, int dummy)
            {

                //foreach (KeyValuePair<string, object> o in data)
                //{
                //    DT.NewColumn(o.Key.ToString(), JMP.colDataTypeConstants.dtTypeCharacter, JMP.colModelTypeConstants.colModelTypeNominal, 20);
                //}
            }
            public void Write_For_Result(string Parameter)
            {
                StreamWrite.Write(Parameter);
            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, string Option)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(Test_Data.ID[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(Option + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j]);
                        // StreamWrite.WriteLine("1");
                    }

                    i++;

                }



            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, KeyValuePair<int, Dictionary<int, string>> OrderbySequence)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(id[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(split[1].ToString() + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j] + ',');
                    }

                    i++;

                }




            }
            public void WriteScript(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
        }
        public class Yield_CSV : INT
        {
            public FileStream OepnSepcFS { get; set; }

            public FileStream OepnSepcFS2 { get; set; }
            public StreamReader StreamReader { get; set; }
            public StreamReader StreamReader2 { get; set; }
            public StreamWriter StreamWrite { get; set; }
            public string[] Get_String { get; set; }
            public int TheFirst_trashes { get; set; }
            public int TheEnd_trashes { get; set; }
            public void Write_Open(string FilePath)
            {
                StreamWrite = new StreamWriter(FilePath, false, Encoding.Default);
            }
            public void Read_Open(string Filename)
            {
                OepnSepcFS = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader = new StreamReader(OepnSepcFS);
            }
            public void Read_Open2(string Filename)
            {
                OepnSepcFS2 = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader2 = new StreamReader(OepnSepcFS2);
            }
            public void Write_Close()
            {
                StreamWrite.Close();
            }
            public void Read_Close()
            {
                StreamReader.Close();
            }
            public void Read2_Close()
            {
                StreamReader2.Close();
            }
            public string[] Read()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');

            }
            public string[] Read2()
            {
                return Get_String = StreamReader2.ReadLine().Trim().Split(',');

            }
            public string Read_Cloth_Spec()
            {
                string Get_Strings = StreamReader.ReadLine().Trim();
                return Get_Strings;
            }
            public string[] Read_Test()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');
            }
            public void GetSplit(string s, char c)
            {
                int l = s.Length;
                int i = 0, j = s.IndexOf(c, 0, l);
                if (j == -1) // No such substring
                {
                    //    yield return s; // Return original and break
                    //    yield break;
                }

                while (j != -1)
                {
                    if (j - i > 0) // Non empty? 
                    {
                        s.Substring(i, j - i); // Return non-empty match
                    }
                    i = j + 1;
                    j = s.IndexOf(c, i, l - i);
                }

                if (i < l) // Has remainder?
                {
                    // yield return s.Substring(i, l - i); // Return remaining trail
                }
            }

            public void Write(Dictionary<string, For_Box> Data)
            {
                int i = 0;

                List<string[]> _D = new List<string[]>();


                i = 0;

                foreach (KeyValuePair<string, For_Box> T in Data)
                {
                    _D.Add(T.Value.ID.ToArray());
                    _D.Add(T.Value.LOT_ID.ToArray());
                    _D.Add(T.Value.SITE_ID.ToArray());
                    _D.Add(T.Value.WAFER_ID.ToArray());
                    break;
                }

                foreach (KeyValuePair<string, For_Box> T in Data)
                {
                    System.Type type;

                    if (T.Value.data == null)
                    {
                        type = T.Value.data_object.GetType();
                    }
                    else
                    {
                        type = T.Value.data.GetType();
                    }


                    if (type.Name == "Double[]")
                    {
                        string[] doubles = Array.ConvertAll<double, string>(T.Value.data, Convert.ToString);
                        _D.Add(doubles);
                    }
                    else
                    {
                        string[] doubles = T.Value.data_object.Cast<string>().ToArray();
                        _D.Add(doubles);
                    }



                }

                i = 0;
                foreach (KeyValuePair<string, For_Box> T in Data)
                {
                    if (Data.Count == 1)
                    {

                        StreamWrite.Write("Label,LOT,SITE,WAFER,");
                        StreamWrite.WriteLine(T.Key);


                    }
                    else
                    {
                        if (i == 0)
                        {
                            StreamWrite.Write("Label,LOT,SITE,WAFER,");
                            StreamWrite.Write(T.Key + ",");
                        }
                        else if (i == Data.Count - 1)
                        {
                            StreamWrite.WriteLine(T.Key);
                        }
                        else
                        {
                            StreamWrite.Write(T.Key + ",");

                        }
                    }

                    i++;
                }

                for (int j = 0; j < _D[0].Length; j++)
                {
                    if (Data.Count == 1)
                    {
                        for (i = 0; i < _D.Count; i++)
                        {
                            if (i != _D.Count - 1)
                            {
                                StreamWrite.Write(_D[i][j] + ",");
                            }

                            else if (i == _D.Count - 1)
                            {
                                StreamWrite.WriteLine(_D[i][j]);
                            }

                        }
                    }
                    else
                    {
                        for (i = 0; i < _D.Count; i++)
                        {
                            if (i != _D.Count - 1)
                            {
                                StreamWrite.Write(_D[i][j] + ",");
                            }

                            else if (i == _D.Count - 1)
                            {
                                StreamWrite.WriteLine(_D[i][j]);
                            }

                        }
                    }

                }

            }
            public void Write(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
            public void Write(string Parameter, string data)
            {
                StreamWrite.WriteLine(Parameter + "\t" + data);
            }
            public void Write(string Parameter, object[] id, object[] data)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }

            }
            public void Write(string Parameter, object[] id, object[] data, object[] Lot, string Variation)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label," + Variation + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.Write(Lot[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());

                    }
                }
            }
            public void Write(string Parameter, string Data1, string Data2)
            {
                StreamWrite.WriteLine(Parameter + "\t" + Data1 + "\t" + Data2);
            }
            public void Write(string String, string Key, int dummy)
            {

                StreamWrite.WriteLine(String);

                //int count = 0;
                //bool falg = false;

                //foreach (Dictionary<string, double[]>[] item in Data)
                //{

                //    foreach (Dictionary<string, double[]> items in item)
                //    {
                //        int j = 0;
                //        falg = false;
                //        foreach (KeyValuePair<string, double[]> o in items)
                //        {
                //            StreamWrite.Write("Parameter" + ',');
                //            for (int i = 0; i < o.Value.Length - 1; i++)
                //            {
                //                StreamWrite.Write((id[i]).ToString() + ',');
                //            }
                //            StreamWrite.WriteLine(id[id.Length -1].ToString());
                //            falg = true;
                //            break;
                //        }
                //        if (falg) break;
                //    }
                //    if (falg) break;
                //}


                //foreach (Dictionary < string, double[]>[] item in Data)
                //{
                //    foreach (Dictionary<string, double[]> items in item)
                //    {
                //        int j = 0;
                //        foreach(KeyValuePair<string,double[]> o in items)
                //        {
                //            StreamWrite.Write(o.Key.ToString() + ',');
                //            for (int i = 0; i < o.Value.Length - 1; i ++)
                //            {
                //                StreamWrite.Write(o.Value[i].ToString() + ',');
                //            }
                //            StreamWrite.WriteLine(o.Value[o.Value.Length - 1].ToString());
                //        }
                //    }

                //}
            }
            public void Write_For_Result(string Parameter)
            {
                StreamWrite.Write(Parameter);
            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, KeyValuePair<int, Dictionary<int, string>> OrderbySequence)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(id[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(split[1].ToString() + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j] + ',');
                    }

                    i++;

                }




            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, string Option)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(Test_Data.ID[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(split[1].ToString() + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j]);
                    }

                    i++;

                }




            }
            public void WriteScript(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
        }
        public class BOXPLOT : INT
        {
            public FileStream OepnSepcFS { get; set; }

            public FileStream OepnSepcFS2 { get; set; }
            public StreamReader StreamReader { get; set; }
            public StreamReader StreamReader2 { get; set; }
            public StreamWriter StreamWrite { get; set; }
            public string[] Get_String { get; set; }
            public int TheFirst_trashes { get; set; }
            public int TheEnd_trashes { get; set; }
            public void Write_Open(string FilePath)
            {
                StreamWrite = new StreamWriter(FilePath, false, Encoding.Default);
            }
            public void Read_Open(string Filename)
            {
                OepnSepcFS = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader = new StreamReader(OepnSepcFS);
            }
            public void Read_Open2(string Filename)
            {
                OepnSepcFS2 = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader2 = new StreamReader(OepnSepcFS2);
            }
            public void Write_Close()
            {
                StreamWrite.Close();
            }
            public void Read_Close()
            {
                StreamReader.Close();
            }
            public void Read2_Close()
            {
                StreamReader2.Close();
            }
            public string[] Read()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');

            }
            public string[] Read2()
            {
                return Get_String = StreamReader2.ReadLine().Trim().Split(',');

            }
            public string Read_Cloth_Spec()
            {
                return null;
            }
            public string[] Read_Test()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');
            }
            public void GetSplit(string s, char c)
            {
                int l = s.Length;
                int i = 0, j = s.IndexOf(c, 0, l);
                if (j == -1) // No such substring
                {
                    //    yield return s; // Return original and break
                    //    yield break;
                }

                while (j != -1)
                {
                    if (j - i > 0) // Non empty? 
                    {
                        s.Substring(i, j - i); // Return non-empty match
                    }
                    i = j + 1;
                    j = s.IndexOf(c, i, l - i);
                }

                if (i < l) // Has remainder?
                {
                    // yield return s.Substring(i, l - i); // Return remaining trail
                }
            }

            public void Write(Dictionary<string, For_Box> Data)
            {
                StreamWrite.WriteLine("");
            }
            public void Write(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
            public void Write(string Parameter, string data)
            {
                StreamWrite.WriteLine(Parameter + "\t" + data);
            }
            public void Write(string Parameter, object[] id, object[] data)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }

            }
            public void Write(string Parameter, object[] id, object[] data, object[] Lot, string Variation)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label,LOT" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.Write(Lot[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());

                    }
                }
            }
            public void Write(string Parameter, string Data1, string Data2)
            {
                StreamWrite.WriteLine(Parameter + "\t" + Data1 + "\t" + Data2);
            }
            public void Write(string String, string Key, int dummy)
            {


            }
            public void Write_For_Result(string Parameter)
            {
                StreamWrite.Write(Parameter);
            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, string Option)
            {

            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, KeyValuePair<int, Dictionary<int, string>> OrderbySequence)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(id[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(split[1].ToString() + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j] + ',');
                    }

                    i++;

                }




            }
            public void WriteScript(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
        }
        public class GETSPEC : INT
        {
            public FileStream OepnSepcFS { get; set; }
            public FileStream OepnSepcFS2 { get; set; }
            public StreamReader StreamReader { get; set; }
            public StreamReader StreamReader2 { get; set; }
            public StreamWriter StreamWrite { get; set; }
            public string[] Get_String { get; set; }
            public int TheFirst_trashes { get; set; }
            public int TheEnd_trashes { get; set; }
            public void Write_Open(string FilePath)
            {
                StreamWrite = new StreamWriter(FilePath, false, Encoding.Default);
            }
            public void Read_Open(string Filename)
            {
                OepnSepcFS = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader = new StreamReader(OepnSepcFS);
            }
            public void Read_Open2(string Filename)
            {
                OepnSepcFS2 = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader2 = new StreamReader(OepnSepcFS2);
            }
            public void Write_Close()
            {
                StreamWrite.Close();
            }
            public void Read_Close()
            {
                StreamReader.Close();
            }

            public void Read2_Close()
            {
                StreamReader2.Close();
            }
            public string[] Read()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');


            }
            public string[] Read2()
            {
                return Get_String = StreamReader2.ReadLine().Trim().Split(',');


            }
            public string Read_Cloth_Spec()
            {
                string Get_Strings = StreamReader.ReadLine().Trim();
                return Get_Strings;
            }
            public string[] Read_Test()
            {

                string Get_String1 = StreamReader.ReadLine();

                GetSplit(Get_String1, ',');

                return null;
            }
            public static IEnumerable<string> GetSplit(string s, char c)
            {
                int l = s.Length;
                int i = 0, j = s.IndexOf(c, 0, l);
                if (j == -1) // No such substring
                {
                    yield return s; // Return original and break
                    yield break;
                }

                while (j != -1)
                {
                    if (j - i > 0) // Non empty? 
                    {
                        yield return s.Substring(i, j - i); // Return non-empty match
                    }
                    i = j + 1;
                    j = s.IndexOf(c, i, l - i);
                }

                if (i < l) // Has remainder?
                {
                    yield return s.Substring(i, l - i); // Return remaining trail
                }
            }

            public void Write(Dictionary<string, For_Box> Data)
            {
                StreamWrite.WriteLine("");
            }
            public void Write(string Parameter)
            {
            }
            public void Write(string Parameter, string data)
            {
                StreamWrite.WriteLine(Parameter + "\t" + data);
            }
            public void Write(string Parameter, object[] id, object[] data)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0) StreamWrite.WriteLine(Parameter);
                    else
                    {
                        StreamWrite.WriteLine(data[i - 1]);
                    }
                }
            }
            public void Write(string Parameter, object[] id, object[] data, object[] Lot, string Variation)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }
            }
            public void Write(string Parameter, string Data1, string Data2)
            {
                StreamWrite.WriteLine(Parameter + "," + Data1 + "," + Data2);
            }
            public void Write_For_Result(string Parameter)
            {
                StreamWrite.Write(Parameter);
            }
            public void Write(string String, string Key, int dummy)
            {

            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, string Option)
            {

            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, KeyValuePair<int, Dictionary<int, string>> OrderbySequence)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(id[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(split[1].ToString() + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j] + ',');
                    }

                    i++;

                }




            }
            public void WriteScript(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
        }

        public class MERGE : INT
        {
            public FileStream OepnSepcFS { get; set; }

            public FileStream OepnSepcFS2 { get; set; }
            public StreamReader StreamReader { get; set; }
            public StreamReader StreamReader2 { get; set; }
            public StreamWriter StreamWrite { get; set; }
            public string[] Get_String { get; set; }
            public int TheFirst_trashes { get; set; }
            public int TheEnd_trashes { get; set; }
            public void Write_Open(string FilePath)
            {
                StreamWrite = new StreamWriter(FilePath, false, Encoding.Default);
            }
            public void Read_Open(string Filename)
            {
                OepnSepcFS = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader = new StreamReader(OepnSepcFS);
            }
            public void Read_Open2(string Filename)
            {
                OepnSepcFS2 = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader2 = new StreamReader(OepnSepcFS2);
            }
            public void Write_Close()
            {
                StreamWrite.Close();
            }
            public void Read_Close()
            {
                StreamReader.Close();
            }
            public void Read2_Close()
            {
                StreamReader2.Close();
            }
            public string[] Read()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');

            }
            public string[] Read2()
            {
                return Get_String = StreamReader2.ReadLine().Trim().Split(',');

            }

            public string Read_Cloth_Spec()
            {
                string Get_Strings = "";

                for (int i = 0; i < Get_String.Length; i++)
                {
                    if (i < Get_String.Length - 1)
                    {
                        Get_Strings += Get_String[i] + ",";
                    }
                    else
                    {
                        Get_Strings += Get_String[i];
                    }

                }
                return Get_Strings;
            }
            public string[] Read_Test()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split(',');
            }
            public void GetSplit(string s, char c)
            {
                int l = s.Length;
                int i = 0, j = s.IndexOf(c, 0, l);
                if (j == -1) // No such substring
                {
                    //    yield return s; // Return original and break
                    //    yield break;
                }

                while (j != -1)
                {
                    if (j - i > 0) // Non empty? 
                    {
                        s.Substring(i, j - i); // Return non-empty match
                    }
                    i = j + 1;
                    j = s.IndexOf(c, i, l - i);
                }

                if (i < l) // Has remainder?
                {
                    // yield return s.Substring(i, l - i); // Return remaining trail
                }
            }

            public void Write(Dictionary<string, For_Box> Data)
            {
                StreamWrite.WriteLine("");
            }
            public void Write(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
            public void Write(string Parameter, string data)
            {
                StreamWrite.WriteLine(Parameter + "\t" + data);
            }
            public void Write(string Parameter, object[] id, object[] data)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }

            }
            public void Write_For_Result(string Parameter)
            {
                StreamWrite.Write(Parameter);
            }
    
            public void Write(string Parameter, object[] id, object[] data, object[] Lot, string Variation)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }
            }
            public void Write(string Parameter, string Data1, string Data2)
            {
                StreamWrite.WriteLine(Parameter + "\t" + Data1 + "\t" + Data2);
            }
            public void Write(string String, string Key, int dummy)
            {
                //int count = 0;
                //bool falg = false;

                //foreach (Dictionary<string, double[]>[] item in Data)
                //{

                //    foreach (Dictionary<string, double[]> items in item)
                //    {
                //        int j = 0;
                //        falg = false;
                //        foreach (KeyValuePair<string, double[]> o in items)
                //        {
                //            StreamWrite.Write("Parameter" + ',');
                //            for (int i = 0; i < o.Value.Length - 1; i++)
                //            {
                //                StreamWrite.Write((id[i]).ToString() + ',');
                //            }
                //            StreamWrite.WriteLine(id[id.Length - 1].ToString());
                //            falg = true;
                //            break;
                //        }
                //        if (falg) break;
                //    }
                //    if (falg) break;
                //}


                //foreach (Dictionary<string, double[]>[] item in Data)
                //{
                //    foreach (Dictionary<string, double[]> items in item)
                //    {
                //        int j = 0;
                //        foreach (KeyValuePair<string, double[]> o in items)
                //        {
                //            StreamWrite.Write(o.Key.ToString() + ',');
                //            for (int i = 0; i < o.Value.Length - 1; i++)
                //            {
                //                StreamWrite.Write(o.Value[i].ToString() + ',');
                //            }
                //            StreamWrite.WriteLine(o.Value[o.Value.Length - 1].ToString());
                //        }
                //    }

                //}
            }

            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, string Option)
            {

            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, KeyValuePair<int, Dictionary<int, string>> OrderbySequence)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(id[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(split[1].ToString() + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j] + ',');
                    }

                    i++;

                }




            }

            public void WriteScript(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
        }

        public class MERGE_S4PD : INT
        {
            public FileStream OepnSepcFS { get; set; }

            public FileStream OepnSepcFS2 { get; set; }
            public StreamReader StreamReader { get; set; }
            public StreamReader StreamReader2 { get; set; }
            public StreamWriter StreamWrite { get; set; }
            public string[] Get_String { get; set; }
            public int TheFirst_trashes { get; set; }
            public int TheEnd_trashes { get; set; }
            public void Write_Open(string FilePath)
            {
                StreamWrite = new StreamWriter(FilePath, false, Encoding.Default);
            }
            public void Read_Open(string Filename)
            {
                OepnSepcFS = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader = new StreamReader(OepnSepcFS);
            }
            public void Read_Open2(string Filename)
            {
                OepnSepcFS2 = new FileStream(Filename, FileMode.Open, FileAccess.Read);
                StreamReader2 = new StreamReader(OepnSepcFS2);
            }
            public void Write_Close()
            {
                StreamWrite.Close();
            }
            public void Read_Close()
            {
                StreamReader.Close();
            }
            public void Read2_Close()
            {
                StreamReader2.Close();
            }
            public string[] Read()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split('\t');

            }
            public string[] Read2()
            {
                return Get_String = StreamReader2.ReadLine().Trim().Split(',');

            }

            public string Read_Cloth_Spec()
            {
                string Get_Strings = "";

                for (int i = 0; i < Get_String.Length; i++)
                {
                    if (i < Get_String.Length - 1)
                    {
                        Get_Strings += Get_String[i] + ",";
                    }
                    else
                    {
                        Get_Strings += Get_String[i];
                    }

                }
                return Get_Strings;
            }
            public string[] Read_Test()
            {
                return Get_String = StreamReader.ReadLine().Trim().Split('\t');
            }
            public void GetSplit(string s, char c)
            {
                int l = s.Length;
                int i = 0, j = s.IndexOf(c, 0, l);
                if (j == -1) // No such substring
                {
                    //    yield return s; // Return original and break
                    //    yield break;
                }

                while (j != -1)
                {
                    if (j - i > 0) // Non empty? 
                    {
                        s.Substring(i, j - i); // Return non-empty match
                    }
                    i = j + 1;
                    j = s.IndexOf(c, i, l - i);
                }

                if (i < l) // Has remainder?
                {
                    // yield return s.Substring(i, l - i); // Return remaining trail
                }
            }

            public void Write(Dictionary<string, For_Box> Data)
            {
                StreamWrite.WriteLine("");
            }
            public void Write(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
            public void Write(string Parameter, string data)
            {
                StreamWrite.WriteLine(Parameter + "\t" + data);
            }
            public void Write(string Parameter, object[] id, object[] data)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }

            }
            public void Write_For_Result(string Parameter)
            {
                StreamWrite.Write(Parameter);
            }

            public void Write(string Parameter, object[] id, object[] data, object[] Lot, string Variation)
            {
                for (int i = 0; i < data.Length + 1; i++)
                {
                    if (i == 0)
                    {
                        StreamWrite.Write("Label" + ',');
                        StreamWrite.WriteLine(Parameter);
                    }
                    else
                    {
                        StreamWrite.Write(id[i - 1].ToString() + ',');
                        StreamWrite.WriteLine(data[i - 1].ToString());
                    }
                }
            }
            public void Write(string Parameter, string Data1, string Data2)
            {
                StreamWrite.WriteLine(Parameter + "\t" + Data1 + "\t" + Data2);
            }
            public void Write(string String, string Key, int dummy)
            {
                //int count = 0;
                //bool falg = false;

                //foreach (Dictionary<string, double[]>[] item in Data)
                //{

                //    foreach (Dictionary<string, double[]> items in item)
                //    {
                //        int j = 0;
                //        falg = false;
                //        foreach (KeyValuePair<string, double[]> o in items)
                //        {
                //            StreamWrite.Write("Parameter" + ',');
                //            for (int i = 0; i < o.Value.Length - 1; i++)
                //            {
                //                StreamWrite.Write((id[i]).ToString() + ',');
                //            }
                //            StreamWrite.WriteLine(id[id.Length - 1].ToString());
                //            falg = true;
                //            break;
                //        }
                //        if (falg) break;
                //    }
                //    if (falg) break;
                //}


                //foreach (Dictionary<string, double[]>[] item in Data)
                //{
                //    foreach (Dictionary<string, double[]> items in item)
                //    {
                //        int j = 0;
                //        foreach (KeyValuePair<string, double[]> o in items)
                //        {
                //            StreamWrite.Write(o.Key.ToString() + ',');
                //            for (int i = 0; i < o.Value.Length - 1; i++)
                //            {
                //                StreamWrite.Write(o.Value[i].ToString() + ',');
                //            }
                //            StreamWrite.WriteLine(o.Value[o.Value.Length - 1].ToString());
                //        }
                //    }

                //}
            }

            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, string Option)
            {

            }
            public void ForBoxplotWrite(string Parameter, object[] id, Dictionary<string, For_Box> Data, KeyValuePair<int, Dictionary<int, string>> OrderbySequence)
            {




                string[] Split = new string[1];

                foreach (string Key in Data.Keys)
                {
                    Split = Key.Split('_');
                    break;
                }

                int i = 0;
                var list = Data.Keys.ToList();
                list.Sort();

                StreamWrite.Write("Label" + ',');
                StreamWrite.Write("Identifier" + ',');
                StreamWrite.Write("Parameter" + ',');
                StreamWrite.Write("Measuer" + ',');
                StreamWrite.Write("Band" + ',');
                StreamWrite.Write("Pmode" + ',');
                StreamWrite.Write("Modulation" + ',');
                StreamWrite.Write("Waveform" + ',');
                StreamWrite.Write("Power_Identifier" + ',');
                StreamWrite.Write("Pout" + ',');
                StreamWrite.Write("Frequency" + ',');
                StreamWrite.Write("Vcc" + ',');
                StreamWrite.Write("Vdd" + ',');
                StreamWrite.Write("DAC1" + ',');
                StreamWrite.Write("DAC2" + ',');
                StreamWrite.Write("TX" + ',');
                StreamWrite.Write("ANT" + ',');
                StreamWrite.Write("RX" + ',');
                StreamWrite.Write("Extra" + ',');
                StreamWrite.Write("Note1" + ',');
                StreamWrite.Write("SpecNumber" + ',');

                StreamWrite.Write("Site" + ',');
                StreamWrite.Write("Lot" + ',');
                StreamWrite.Write("Wafer" + ',');


                StreamWrite.WriteLine(Split[1]);

                foreach (string key in list)
                {
                    string[] split = key.Split('_');
                    For_Box Test_Data = Data[key];



                    for (int j = 0; j < Test_Data.data.Length; j++)
                    {
                        StreamWrite.Write(id[j].ToString() + ',');
                        StreamWrite.Write(split[0].ToString() + ',');
                        StreamWrite.Write(split[1].ToString() + ',');
                        StreamWrite.Write(split[2].ToString() + ',');
                        StreamWrite.Write(split[3].ToString() + ',');
                        StreamWrite.Write(split[4].ToString() + ',');
                        StreamWrite.Write(split[5].ToString() + ',');
                        StreamWrite.Write(split[6].ToString() + ',');
                        StreamWrite.Write(split[7].ToString() + ',');
                        StreamWrite.Write(split[8].ToString() + ',');
                        StreamWrite.Write(split[9].ToString() + ',');
                        StreamWrite.Write(split[10].ToString() + ',');
                        StreamWrite.Write(split[11].ToString() + ',');
                        StreamWrite.Write(split[12].ToString() + ',');
                        StreamWrite.Write(split[13].ToString() + ',');
                        StreamWrite.Write(split[14].ToString() + ',');
                        StreamWrite.Write(split[15].ToString() + ',');
                        StreamWrite.Write(split[16].ToString() + ',');
                        StreamWrite.Write(split[17].ToString() + ',');
                        StreamWrite.Write(split[18].ToString() + ',');
                        StreamWrite.Write(split[19].ToString() + ',');

                        StreamWrite.Write(Test_Data.SITE_ID[j] + ',');
                        StreamWrite.Write(Test_Data.LOT_ID[j] + ',');
                        StreamWrite.Write(Test_Data.WAFER_ID[j] + ',');

                        StreamWrite.WriteLine(Test_Data.data[j] + ',');
                    }

                    i++;

                }




            }

            public void WriteScript(string Parameter)
            {
                StreamWrite.WriteLine(Parameter);
            }
        }

        public interface INT
        {
            FileStream OepnSepcFS { get; set; }
            FileStream OepnSepcFS2 { get; set; }
            StreamReader StreamReader { get; set; }
            StreamReader StreamReader2 { get; set; }
            StreamWriter StreamWrite { get; set; }

            int TheFirst_trashes { get; set; }
            int TheEnd_trashes { get; set; }
            string[] Get_String { get; set; }
            void Write_Open(string FilePath);
            void Read_Open(string FIleName);
            void Read_Open2(string FIleName);
            void Write_Close();
            void Read_Close();

            void Read2_Close();
            string[] Read();
            string[] Read2();

            string Read_Cloth_Spec();
            string[] Read_Test();

            void Write(Dictionary<string, For_Box> Data);
            void Write(string Parameter);
            void Write(string Parameter, string Data);
            void Write(string Parameter, object[] ID, object[] Data);
            void Write_For_Result(string Parameter);
            void Write(string Parameter, object[] ID, object[] Data, object[] Lot, string Variation);
            void Write(string Parameter, string Data1, string Data2);
            void Write(string String, string Key, int dummy);
            void ForBoxplotWrite(string Parameter, object[] ID, Dictionary<string, For_Box> Data, string Option);
            void ForBoxplotWrite(string Parameter, object[] ID, Dictionary<string, For_Box> Data, KeyValuePair<int, Dictionary<int, string>> OrderbySequence);
            void WriteScript(string Parameter);
        }

        public INT Open(string Key)
        {
            INT Int = null;

            switch (Key.ToUpper().Trim())
            {
                case "YIELD":
                    Int = new Yield_CSV();
                    break;
                case "BOXPLOT":
                    Int = new BOXPLOT();
                    break;
                case "FCM":
                    Int = new FCM_Automation_CSV();
                    break;
                case "GETSPEC":
                    Int = new GETSPEC();
                    break;
                case "MERGE":
                    Int = new MERGE();
                    break;
                case "MERGE_S4PD":
                    Int = new MERGE_S4PD();
                    break;
            }
            return Int;
        }


    }

    public class For_Box
    {
        public string Parameter;
        public double[] data;
        public object[] data_object;
        public string[] ID;
        public string[] WAFER_ID;
        public string[] SITE_ID;
        public string[] LOT_ID;


        public double Min;
        public double Max;
        public string STD;
        public string Median;
        public string Yeild;
        public string Apple_Spec_Min;
        public string Apple_Spec_Max;
        public string Broadcom_Spec_Min;
        public string Broadcom_Spec_Max;

        public For_Box(string Parameter, double[] data, string[] ID, string[] WAFER_ID, string[] SITE_ID, string[] LOT_ID, double Min, double Max, string STD, string Median, string Yeild, string Apple_Spec_Min, string Apple_Spec_Max, string Broadcom_Spec_Min, string Broadcom_Spec_Max)
        {
            this.Parameter = Parameter;
            this.data = data;
            this.ID = ID;
            this.WAFER_ID = WAFER_ID;
            this.SITE_ID = SITE_ID;
            this.LOT_ID = LOT_ID;
            this.Min = Min;
            this.Max = Max;
            this.STD = STD;
            this.Median = Median;
            this.Yeild = Yeild;
            this.Apple_Spec_Min = Apple_Spec_Min;
            this.Apple_Spec_Max = Apple_Spec_Max;
            this.Broadcom_Spec_Min = Broadcom_Spec_Min;
            this.Broadcom_Spec_Max = Broadcom_Spec_Max;
        }

        public For_Box(string Parameter, object[] data_object, string[] ID, string[] WAFER_ID, string[] SITE_ID, string[] LOT_ID, double Min, double Max, string STD, string Median, string Yeild, string Apple_Spec_Min, string Apple_Spec_Max, string Broadcom_Spec_Min, string Broadcom_Spec_Max)
        {
            this.Parameter = Parameter;
            this.data_object = data_object;
            this.ID = ID;
            this.WAFER_ID = WAFER_ID;
            this.SITE_ID = SITE_ID;
            this.LOT_ID = LOT_ID;
            this.Min = Min;
            this.Max = Max;
            this.STD = STD;
            this.Median = Median;
            this.Yeild = Yeild;
            this.Apple_Spec_Min = Apple_Spec_Min;
            this.Apple_Spec_Max = Apple_Spec_Max;
            this.Broadcom_Spec_Min = Broadcom_Spec_Min;
            this.Broadcom_Spec_Max = Broadcom_Spec_Max;
        }
    }

}
