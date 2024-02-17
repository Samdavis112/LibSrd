using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using System.Windows;

namespace LibSrd_NETCore
{
    public static class Conversion
    {
        #region CSV
        /// <summary>
        /// Returns a datatable from a CSV file.
        /// </summary>
        /// <returns></returns>
        public static DataTable CSVToDataTable(string FilePath, out string ErrMsg)
        {
            DataTable Dt = new DataTable("Dt");
            ErrMsg = null;

            if (!File.Exists(FilePath))
            {
                ErrMsg = "File not found";
                return null;
            }

            string[] CSVLines = File.ReadAllLines(FilePath);

            //Returning empty datatable if file empty
            if (CSVLines == null || CSVLines.Count() < 1)
            {
                return Dt;
            }

            //Headers
            string[] Headers = CSVLines[0].Split(',');
            for (int i = 0; i < Headers.Count(); i++)
            {
                if (Headers[i] == null || Headers[i] == "")//Missing Header
                {
                    ErrMsg = "Invalid header in CSV file.";
                    return null;
                }
                Headers[i] = Headers[i].Trim();
                Headers[i] = Headers[i].Trim('"');
                Dt.Columns.Add(Headers[i]);
            }

            //Content
            if (CSVLines.Length == 1) return Dt; //Just Headers

            for (int i = 1; i < CSVLines.Count(); i++)
            {
                if (CSVLines[i] == null || CSVLines[i] == "") //Empty row
                    continue;
                string[] Entry = CSVLines[i].Split(',');
                for (int j = 0; j < Entry.Count(); j++)
                {
                    if (Entry[j] == null || Entry[j] == "") //Blank entry
                        continue;

                    Entry[j] = Entry[j].Trim();
                    Entry[j] = Entry[j].Trim('"');
                }
                Dt.Rows.Add(Entry);
            }

            return Dt;
        }

        /// <summary>
        /// Returns a datatable from this string array.
        /// </summary>
        /// <returns></returns>
        public static DataTable CSVToDataTable(this string[] CSVLines, out string ErrMsg)
        {
            DataTable Dt = new DataTable("Dt");
            ErrMsg = null;

            //Returning empty datatable if file empty
            if (CSVLines == null || CSVLines.Count() < 1)
            {
                return Dt;
            }

            //Headers
            string[] Headers = CSVLines[0].Split(',');
            for (int i = 0; i < Headers.Count(); i++)
            {
                if (Headers[i] == null || Headers[i] == "")//Missing Header
                {
                    ErrMsg = "Invalid header in CSV file.";
                    return null;
                }
                Headers[i] = Headers[i].Trim();
                Headers[i] = Headers[i].Trim('"');
                Dt.Columns.Add(Headers[i]);
            }

            //Content
            if (CSVLines.Length == 1) return Dt; //Just Headers

            for (int i = 1; i < CSVLines.Count(); i++)
            {
                if (CSVLines[i] == null || CSVLines[i] == "") //Empty row
                    continue;
                string[] Entry = CSVLines[i].Split(',');
                for (int j = 0; j < Entry.Count(); j++)
                {
                    if (Entry[j] == null || Entry[j] == "") //Blank entry
                        continue;

                    Entry[j] = Entry[j].Trim();
                    Entry[j] = Entry[j].Trim('"');
                }
                Dt.Rows.Add(Entry);
            }

            return Dt;
        }

        /// <summary>
        /// Writes a Datatable to a CSV file.
        /// </summary>
        /// <returns></returns>
        public static void ToCSV(this DataTable DT, string FilePath, out string ErrMsg)
        {
            ErrMsg = null;

            if (DT == null || DT.Columns.Count < 1)
            {
                return;
            }

            //Headers
            List<string> Headers = new List<string>();
            foreach (DataColumn dc in DT.Columns)
            {
                Headers.Add(dc.ColumnName);
            }
            string HeadersCSV = null;

            for (int i = 0; i < Headers.Count() - 1; i++)
            {
                HeadersCSV += $"\"{Headers[i]}\",";
            }
            HeadersCSV += $"\"{Headers[Headers.Count() - 1]}\"";

            //Data
            List<string> Rows = new List<string>();
            if (DT.Rows.Count > 0)
            {
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    Rows.Add(string.Empty);
                    string[] Entrys = new string[DT.Columns.Count];
                    for (int j = 0; j < DT.Columns.Count; j++)
                    {
                        if (DT.Rows[i][j] != null || DT.Rows[i][j].ToString() != " ")
                            Entrys[j] = $"\"{DT.Rows[i][j]}\"";
                        else
                            Entrys[j] = " ";
                    }

                    for (int c = 0; c < Entrys.Count() - 1; c++)
                    {
                        Rows[i] += Entrys[c] + ",";
                    }
                    Rows[i] += Entrys[Entrys.Count() - 1];
                }
            }

            //Writing
            List<string> Lines = new List<string>();
            Lines.Add(HeadersCSV);

            if (Rows != null || Rows.Count() > 0)
            {
                foreach (string item in Rows)
                {
                    Lines.Add(item);
                }
            }

            File.WriteAllLines(FilePath, Lines.ToArray());
        }

        /// <summary>
        /// Returns a string array of the lines of a CSV file from a datatable
        /// </summary>
        /// <returns></returns>
        public static string[] ToCSV(this DataTable DT, out string ErrMsg)
        {
            ErrMsg = null;

            if (DT == null || DT.Columns.Count < 1)
            {
                return null;
            }

            //Headers
            List<string> Headers = new List<string>();
            foreach (DataColumn dc in DT.Columns)
            {
                Headers.Add(dc.ColumnName);
            }
            string HeadersCSV = null;

            for (int i = 0; i < Headers.Count() - 1; i++)
            {
                HeadersCSV += $"\"{Headers[i]}\",";
            }
            HeadersCSV += $"\"{Headers[Headers.Count() - 1]}\"";

            //Data
            List<string> Rows = new List<string>();
            if (DT.Rows.Count > 0)
            {
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    Rows.Add(string.Empty);
                    string[] Entrys = new string[DT.Columns.Count];
                    for (int j = 0; j < DT.Columns.Count; j++)
                    {
                        if (DT.Rows[i][j] != null || DT.Rows[i][j].ToString() != " ")
                            Entrys[j] = $"\"{DT.Rows[i][j]}\"";
                        else
                            Entrys[j] = " ";
                    }

                    for (int c = 0; c < Entrys.Count() - 1; c++)
                    {
                        Rows[i] += Entrys[c] + ",";
                    }
                    Rows[i] += Entrys[Entrys.Count() - 1];
                }
            }

            //Writing
            List<string> Lines = new List<string>();
            Lines.Add(HeadersCSV);

            if (Rows != null || Rows.Count() > 0)
            {
                foreach (string item in Rows)
                {
                    Lines.Add(item);
                }
            }

            return Lines.ToArray();
        }
        #endregion

        #region XML
        /// <summary>
        /// Returns a datatable from a XML filepath.
        /// </summary>
        /// <param name="Filepath">A string contaning the filepath to the XML file</param>
        /// <param name="ErrMsg">Will contain an error message if error occours</param>
        public static DataTable XMLToDataTable(string Filepath, out string ErrMsg)
        {
            ErrMsg = null;
            if (!File.Exists(Filepath))
            {
                ErrMsg = "File does not exist, or is innaccessable.";
                return null;
            }

            try
            {
                DataSet ds = new DataSet();
                ds.ReadXml(Filepath);

                DataTable Dt = ds.Tables[0];

                return Dt;
            }
            catch (Exception e)
            {
                ErrMsg = e.Message;
                return null;
            }
        }
        /// <summary>
        /// Writes an xml file from the given datatable. [UNTESTED]
        /// </summary>
        public static void ToXML(this DataTable dt, string Filepath, out string ErrMsg)
        {
            ErrMsg = null;

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);

            try
            {
                ds.WriteXml(Filepath);
            }
            catch (Exception ex) { ErrMsg = ex.Message; }
        }
        #endregion

        #region Json
        public static object JsonToObject(string Json, Type type)
        {
            return JsonConvert.DeserializeObject(Json, type);
        }
        public static string ToJson(this object obj)
        {
            return JsonConvert.SerializeObject(obj, Formatting.Indented);
        }
        #endregion
    }
}