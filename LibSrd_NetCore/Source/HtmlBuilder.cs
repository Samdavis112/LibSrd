using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace LibSrd_NETCore
{
    public class HtmlBuilder
    {
        private StringBuilder htmlDocument = new StringBuilder();

        /// <summary>
        /// Open the html file and initalises the object.
        /// </summary>
        /// <param name="Title">Form Title</param>
        /// <param name="useDefaultStyles">Should the default styles be used?</param>
        /// <param name="StyleFilepath">path of CSS file</param>
        /// <param name="DefaultImageRoot">Base path of default images in the document</param>
        public HtmlBuilder(string Title, bool useDefaultStyles = true, string StyleFilepath = null, string DefaultImageRoot = null)
        {
            htmlDocument.AppendLine("<html>");
            htmlDocument.AppendLine("<head>");

            if (useDefaultStyles)
            {
                htmlDocument.AppendLine("<style>");
                // htmlDocument.AppendLine(@"@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@100;200;300;400&display=swap');");
                htmlDocument.AppendLine("* { font-family: Verdana, arial, Helvetica, sans-serif; }");
                htmlDocument.AppendLine("body { font-size: 10.5pt; }");
                htmlDocument.AppendLine("H1 { font-size: 17pt;}");
                htmlDocument.AppendLine("H2 { font-size: 13pt;}");
                htmlDocument.AppendLine("H3 { font-size: 11pt;}");
                htmlDocument.AppendLine("H4,H5,TH,TD { font-size: 10pt;}");
                htmlDocument.AppendLine("table { border-collapse: collapse; }");
                htmlDocument.AppendLine("table, th { border: solid #888888 1px; padding: 6px; }");
                htmlDocument.AppendLine("table, td { border: solid #888888 1px; padding: 4px; }");
                htmlDocument.AppendLine("</style>");
            }

            htmlDocument.AppendLine("<title>" + Title + "</title>");
            if (StyleFilepath != null && StyleFilepath != "")
            {
                htmlDocument.AppendLine("<link rel=\"stylesheet\" type=\"text/css\" href=\"" + StyleFilepath + "\" title=\"DevEng Style\" />");
            }
            if (DefaultImageRoot != null && DefaultImageRoot != "")
            {
                //<base href="http://www.w3schools.com/images/"
                htmlDocument.AppendLine("<base href=\"" + DefaultImageRoot + "\">");
            }
            htmlDocument.AppendLine("</head>");
            htmlDocument.AppendLine("<body>");
        }

        #region Comments/RawText/HTML
        /// <summary>
        /// Adds raw text. Html formatting must be manually applied.
        /// See also InsertSnippet(...) 
        /// </summary>
        /// <param name="Text"></param>
        public void AppendRawText(string Text)
        {
            htmlDocument.Append(Text);
            htmlDocument.AppendLine();
        }

        /// <summary>
        /// Inserts a raw HTML snippet. Optionally the HTML snippet may have placeholders that are replaced with values from a dictionary.
        /// This is useful for such as complex tables that dont map well to a simple DataTable.
        /// </summary>
        /// <param name="snippet">A block of html code, optionally with 'placeholders' that are substituted with values from Placeholder_ValuesDict</param>
        /// <param name="Placeholder_InsertionsDict">Dictionary of Placeholder to values. e.g. "{title}" to "Summary Table". Can be null.</param>
        public void InsertSnippet(StringBuilder snippet, Dictionary<string, string> Placeholder_ValuesDict = null)
        {
            if (Placeholder_ValuesDict != null)
            {
                foreach (var item in Placeholder_ValuesDict.Keys)
                {
                    snippet.Replace(item, Placeholder_ValuesDict[item]);
                }
            }

            htmlDocument.Append(snippet);
            return;
        }

        /// <summary>
        /// Inserts a raw HTML snippet. Optionally the HTML snippet may have placeholders that are replaced with values from a dictionary.
        /// This is useful for such as complex tables that dont map well to a simple DataTable.
        /// </summary>
        /// <param name="snippet">A block of html code, optionally with 'placeholders' that are substituted with values from Placeholder_ValuesDict</param>
        /// <param name="Placeholder_ValuesDict">Dictionary of Placeholder to values. e.g. "{title}" to "Summary Table". Can be null.</param>
        public void InsertSnippet(string snippet, Dictionary<string, string> Placeholder_ValuesDict = null)
        {
            if (Placeholder_ValuesDict != null)
            {
                foreach (var item in Placeholder_ValuesDict.Keys)
                {
                    snippet.Replace(item, Placeholder_ValuesDict[item]);
                }
            }

            htmlDocument.Append(snippet);
            return;
        }

        public void Comment(string Comment)
        {
            htmlDocument.AppendLine("<!--" + Comment + "-->");
        }
        #endregion

        #region Paragraphs/Breaks/Headings
        public void Heading(string heading, int level, string Class = null)
        {
            Heading(heading, level, null, Class);
        }

        public void Heading(string heading, int level, string label, string Class = null)
        {
            string lvl = level.ToString();
            if (label == null || label.Length == 0)
                htmlDocument.AppendLine("<h" + lvl + GetClass(Class) + ">" + heading + "</h" + lvl + ">");
            else
                htmlDocument.AppendLine("<h" + lvl + GetClass(Class) + "><a id=\"" + label + "\">" + heading + "</a></h" + lvl + ">");
        }

        public void Para(string text, string Class = null)
        {
            htmlDocument.AppendLine("<p" + GetClass(Class) + ">" + text + "</p>");
        }

        public void Para(StringBuilder text, string Class = null)
        {
            htmlDocument.AppendLine("<p" + GetClass(Class) + ">" + text.ToString() + "</p>");
        }

        public void Line()
        {
            htmlDocument.AppendLine("<hr>");
        }

        public void Break()
        {
            htmlDocument.AppendLine("<br>");
        }

        public void ContactInfo(IEnumerable<string> Lines, string Class = null)
        {
            /* <address>
                    Written by W3Schools.com<br>
                    <a href="mailto:us@example.org">Email us</a><br>
                    Address: Box 564, Disneyland<br>
                    Phone: +12 34 56 78
                </address> */
            htmlDocument.AppendLine("<address" + GetClass(Class) + ">");
            foreach (string line in Lines)
            {
                htmlDocument.AppendLine(line + "<br>");
            }
            htmlDocument.AppendLine("</address>");
        }

        public void ContactInfo(string Author, string Email, string Class = null)
        {
            /* <address>
                    Written by W3Schools.com<br>
                    <a href="mailto:us@example.org">Email us</a><br>
                    Address: Box 564, Disneyland<br>
                    Phone: +12 34 56 78
                </address> */
            htmlDocument.AppendLine("<address" + GetClass(Class) + ">");
            htmlDocument.AppendLine("Author: " + Author + "<br>");
            htmlDocument.AppendLine("Email: " + FormatEmailLink(Email, Email) + "<br>");
            htmlDocument.AppendLine("</address>");
        }
        #endregion

        #region Misc
        /// <summary>
        /// Writes the html document to the filepath provided.
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public bool WriteHtml(string filepath)
        {
            htmlDocument.AppendLine("</body>");
            htmlDocument.AppendLine("</html>");
            try
            {
                File.WriteAllText(filepath, htmlDocument.ToString());
                htmlDocument.Replace("</body>", "");
                htmlDocument.Replace("</html>", "");
            }
            catch
            {
                htmlDocument.Replace("</body>", "");
                htmlDocument.Replace("</html>", "");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Will return the html document in a string.
        /// </summary>
        /// <returns></returns>
        public string GetHtml()
        {
            // Add finishing tags
            htmlDocument.AppendLine("</body>");
            htmlDocument.AppendLine("</html>");
            string temp = htmlDocument.ToString();

            // Remove finishing tags for next time.
            htmlDocument.Replace("</body>", "");
            htmlDocument.Replace("</html>", "");

            return temp;
        }

        /// <summary>
        /// Gets the contents of the StringBuilder as is.
        /// It is normal to use CloseGetHtml() which closes the html properly first.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return htmlDocument.ToString();
        }
        #endregion

        #region Lists
        public void ListGeneral(ListObj listObj)
        {
            htmlDocument.AppendLine(listObj.GenerateHtml());
        }

        public void ListPlain(IEnumerable<string> items, string Class = null)
        {
            if (items == null) return;
            htmlDocument.AppendLine("<para" + GetClass(Class) + ">");
            foreach (string item in items)
            {
                htmlDocument.AppendLine(item + "<br/>");
            }
            htmlDocument.AppendLine("</para>");
        }

        public void ListBullet(IEnumerable<string> items, string Class = null)
        {
            if (items == null) return;
            htmlDocument.AppendLine("<ul" + GetClass(Class) + ">");
            foreach (string item in items)
            {
                htmlDocument.AppendLine("<li>" + item + "</li>");
            }
            htmlDocument.AppendLine("</ul>");
        }

        public void ListNumbered(IEnumerable<string> items, string Class = null)
        {
            if (items == null) return;
            htmlDocument.AppendLine("<ol" + GetClass(Class) + ">");
            foreach (string item in items)
            {
                htmlDocument.AppendLine("<li>" + item + "</li>");
            }
            htmlDocument.AppendLine("</ol>");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Items1">First level List items</param>
        /// <param name="Items2">2nd level list item</param>
        public void ListNumbered(string[] Items1, string[][] Items2, string Class = null)   // NEEDS VERIFYING
        {
            if (Items1 == null) return;
            htmlDocument.AppendLine("<ol" + GetClass(Class) + ">");
            for (int i = 0; i < Items1.Length; i++)
            {
                htmlDocument.AppendLine("<li>" + Items1[i] + "</li>");
                string[] sublist = Items2[i];
                if (sublist != null && sublist.Length != 0)
                {
                    htmlDocument.AppendLine("<ol>");
                    foreach (string subitem in sublist)
                    {
                        htmlDocument.AppendLine("<li>" + subitem + "</li>");
                    }
                    htmlDocument.AppendLine("</ol>");
                }
            }
            htmlDocument.AppendLine("</ol>");
        }
        #endregion

        #region Tables
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ColHeaders"></param>
        /// <param name="ColAlignments"></param>
        /// <param name="TableData"></param>
        /// <param name="BorderFlag"></param>
        public void Table(string[] ColHeaders, string[] ColAlignments, string[,] TableData, bool BorderFlag, bool CenterFlag, string Class = null)
        {
            string[][] tab = new string[TableData.GetLength(0)][];
            for (int i = 0; i < TableData.GetLength(0); i++)
            {
                tab[i] = new string[TableData.GetLength(1)];
                for (int j = 0; j < tab[i].Length; j++)
                {
                    tab[i][j] = TableData[i, j];
                }
            }
            Table(ColHeaders, ColAlignments, tab, null, true, BorderFlag, CenterFlag, Class);
        }

        public void Table(IList<string> ColHeaders, IList<string> ColAlignments, IList<IList<string>> TableData, bool RowsOfColumnsFlag, bool BorderFlag, bool CenterFlag, string Class = null)
        {
            IList<IList<string>> ColorData = null;
            Table(ColHeaders, ColAlignments, TableData, ColorData, RowsOfColumnsFlag, BorderFlag, CenterFlag, Class);
        }

        public void Table(IList<string> ColHeaders, IList<string> ColAlignments, IList<IList<string>> TableData, IList<IList<string>> ColorData, bool RowsOfColumnsFlag, bool BorderFlag, bool CenterFlag, string Class = null)
        {
            if (TableData == null || TableData.Count == 0) return;

            if (CenterFlag) htmlDocument.Append("<center>");
            htmlDocument.Append("<table" + GetClass(Class)); if (BorderFlag) htmlDocument.Append(" border=\"1\""); htmlDocument.AppendLine(">");   // <table border="1">

            int Ncols = TableData[0].Count;   // TableData[0] is first row in RowsOfColumns mode
            if (!RowsOfColumnsFlag) Ncols = TableData.Count;

            // preprocess Column alignments to ensure legal
            if (ColAlignments == null || ColAlignments.Count < Ncols) ColAlignments = new string[Ncols];
            for (int j = 0; j < ColAlignments.Count; j++)
            {
                if (ColAlignments[j] == null || ColAlignments[j].Length == 0) ColAlignments[j] = "center";
                else if (ColAlignments[j].Substring(0, 1).ToUpper() == "L") ColAlignments[j] = "left";
                else if (ColAlignments[j].Substring(0, 1).ToUpper() == "R") ColAlignments[j] = "right";
                else ColAlignments[j] = "center";

                ColAlignments[j] = ColAlignments[j].ToLower().Replace("centre", "center");
            }

            if (ColHeaders != null)     // <tr> <th align="left">Extension</th>  <th align="left">Format</th> </tr>
            {

                htmlDocument.Append("<tr> ");
                for (int j = 0; j < ColHeaders.Count; j++)
                {
                    //  htmlDocument.Append("<th align=\""+ColAlignments[j]+"\">"+ColHeaders[j]+"</th>");
                    htmlDocument.Append("<th align=\"center\">" + ColHeaders[j] + "</th>");  // hard-wire header alignment to centered
                }
                htmlDocument.AppendLine(" </tr>");
            }

            if (RowsOfColumnsFlag)
            {
                for (int i = 0; i < TableData.Count; i++)    //  <tr> <td>No ext.</td>      <td>RFMD assembly format</td> </tr>
                {
                    htmlDocument.Append("<tr> ");
                    for (int j = 0; j < TableData[i].Count; j++)
                    {
                        htmlDocument.Append("<td");
                        if (ColAlignments != null && ColAlignments[j] != null && ColAlignments[j] != null && ColAlignments[j] != "") htmlDocument.Append(" align=\"" + ColAlignments[j] + "\"");
                        if (ColorData != null && ColorData[i] != null && ColorData[i][j] != null && ColorData[i][j] != null && ColorData[i][j] != "") htmlDocument.Append(" bgcolor=\"" + ColorData[i][j] + "\"");
                        htmlDocument.Append(">");
                        htmlDocument.Append(TableData[i][j]);
                        htmlDocument.Append("</td>");
                    }
                    htmlDocument.AppendLine(" </tr>");
                }
            }
            else
            {
                for (int i = 0; i < TableData[0].Count; i++)    //  <tr> <td>No ext.</td>      <td>RFMD assembly format</td> </tr>
                {
                    htmlDocument.Append("<tr> ");
                    for (int j = 0; j < TableData.Count; j++)
                    {
                        htmlDocument.Append("<td");
                        if (ColAlignments != null && ColAlignments[j] != null && ColAlignments[j] != null && ColAlignments[j] != "") htmlDocument.Append(" align=\"" + ColAlignments[j] + "\"");
                        if (ColorData != null && ColorData[j] != null && ColorData[j][i] != null && ColorData[j][i] != "") htmlDocument.Append(" bgcolor=\"" + ColorData[j][i] + "\"");
                        htmlDocument.Append(">");
                        htmlDocument.Append(TableData[j][i]);
                        htmlDocument.Append("</td>");
                    }
                    htmlDocument.AppendLine(" </tr>");
                }
            }
            htmlDocument.AppendLine("</table>");
            if (CenterFlag) htmlDocument.Append("</center>");
        }

        /// <summary>
        /// The best approach for tables
        /// </summary>
        /// <param name="Main">DataTable of strings</param>
        /// <param name="ColAlignments">array of L, R, C, left. right, center, centre codes. If length is lesst than no of columns in the datatable then the
        ///                             alignments get topped up with center's</param>
        /// <param name="ColorData">DataTable of strings describing the shading color. No entry for a cell means not shaded</param>
        /// <param name="BorderFlag"></param>
        /// <param name="CenterFlag"></param>
        /// <param name="defaultFormat">The string format to use in tables if values are real</param>
        public void Table(DataTable Main, IList<string> ColAlignments, DataTable ColorData, bool BorderFlag, bool CenterFlag, string defaultFormat = "0.###", string Class = null)
        {
            if (Main == null || Main.Rows.Count == 0) return;

            if (CenterFlag) htmlDocument.Append("<center>");
            htmlDocument.Append("<table" + GetClass(Class)); if (BorderFlag) htmlDocument.Append(" border=\"1\""); htmlDocument.AppendLine(">");   // <table border="1">

            //int Ncols = TableData[0].Length;   // TableData[0] is first row in RowsOfColumns mode
            //if (!RowsOfColumnsFlag) Ncols = TableData.Length;

            // preprocess Column alignments to ensure legal
            if (ColAlignments == null) ColAlignments = new string[Main.Columns.Count];

            if (ColAlignments.Count < Main.Columns.Count) // extend partially supplied alignments to fill with C's
            {
                var _ColAlignments = new string[Main.Columns.Count];
                for (int i = 0; i < _ColAlignments.Length; i++)
                {
                    if (i < ColAlignments.Count) _ColAlignments[i] = ColAlignments[i]; else _ColAlignments[i] = "C";
                }
                ColAlignments = _ColAlignments;
            }
            //if (ColAlignments == null || ColAlignments.Count < Main.Columns.Count) ColAlignments = new string[Main.Columns.Count];
            for (int j = 0; j < ColAlignments.Count; j++)
            {
                if (ColAlignments[j] == null || ColAlignments[j].Length == 0) ColAlignments[j] = "center";
                else if (ColAlignments[j].Substring(0, 1).ToUpper() == "L") ColAlignments[j] = "left";
                else if (ColAlignments[j].Substring(0, 1).ToUpper() == "R") ColAlignments[j] = "right";
                else ColAlignments[j] = "center";

                ColAlignments[j] = ColAlignments[j].ToLower().Replace("centre", "center");
            }

            //ColHeaders: <tr> <th align="left">Extension</th>  <th align="left">Format</th> </tr>
            htmlDocument.Append("<tr> ");
            for (int j = 0; j < Main.Columns.Count; j++)
            {
                //  htmlDocument.Append("<th align=\""+ColAlignments[j]+"\">"+ColHeaders[j]+"</th>");
                htmlDocument.Append("<th align=\"center\">" + Main.Columns[j].ColumnName + "</th>");  // hard-wire header alignment to centered
            }
            htmlDocument.AppendLine(" </tr>");


            // fill in the table
            for (int i = 0; i < Main.Rows.Count; i++)    //  <tr> <td>No ext.</td>      <td>RFMD assembly format</td> </tr>
            {
                htmlDocument.Append("<tr> ");
                for (int j = 0; j < Main.Columns.Count; j++)
                {
                    htmlDocument.Append("<td");
                    if (ColAlignments != null && ColAlignments[j] != null && ColAlignments[j] != null && ColAlignments[j] != "") htmlDocument.Append(" align=\"" + ColAlignments[j] + "\"");
                    if (ColorData != null && ColorData.Rows[i][j] != DBNull.Value && (string)ColorData.Rows[i][j] != "") htmlDocument.Append(" bgcolor=\"" + (string)ColorData.Rows[i][j] + "\"");
                    htmlDocument.Append(">");
                    if (Main.Rows[i][j] != DBNull.Value && Main.Columns[j].DataType == typeof(double))
                        htmlDocument.Append(((double)Main.Rows[i][j]).ToString(defaultFormat));
                    else if (Main.Rows[i][j] != DBNull.Value && Main.Columns[j].DataType == typeof(float))
                        htmlDocument.Append(((float)Main.Rows[i][j]).ToString(defaultFormat));
                    else
                        htmlDocument.Append(Main.Rows[i][j]);
                    htmlDocument.Append("</td>");
                }
                htmlDocument.AppendLine(" </tr>");
            }

            htmlDocument.AppendLine("</table>");
            if (CenterFlag) htmlDocument.Append("</center>");
        }

        /// <summary>
        /// UNTESTED
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="Format">Format string (for non string data elements)</param>
        /// <param name="ColAlignments"></param>
        /// <param name="ColorData">string[j][i] ColorData</param>
        /// <param name="BorderFlag"></param>
        /// <param name="CenterFlag"></param>
        public void Table(DataTable dataTable, IList<string> Format, IList<string> ColAlignments, IList<IList<string>> ColorData, bool BorderFlag, bool CenterFlag, string Class = null)
        {
            if (dataTable == null || dataTable.Columns.Count == 0) return;

            if (CenterFlag) htmlDocument.Append("<center>");
            htmlDocument.Append("<table" + GetClass(Class)); if (BorderFlag) htmlDocument.Append(" border=\"1\""); htmlDocument.AppendLine(">");   // <table border="1">

            int Ncols = dataTable.Columns.Count;

            // preprocess Column alignments to ensure legal
            if (ColAlignments == null || ColAlignments.Count < Ncols) ColAlignments = new string[Ncols];
            for (int j = 0; j < ColAlignments.Count; j++)
            {
                if (ColAlignments[j] == null || ColAlignments[j].Length == 0) ColAlignments[j] = "center";
                else if (ColAlignments[j].Substring(0, 1).ToUpper() == "L") ColAlignments[j] = "left";
                else if (ColAlignments[j].Substring(0, 1).ToUpper() == "R") ColAlignments[j] = "right";
                else ColAlignments[j] = "center";

                ColAlignments[j] = ColAlignments[j].ToLower().Replace("centre", "center");
            }


            htmlDocument.Append("<tr> ");
            for (int j = 0; j < Ncols; j++)
            {
                //  htmlDocument.Append("<th align=\""+ColAlignments[j]+"\">"+ColHeaders[j]+"</th>");
                htmlDocument.Append("<th align=\"center\">" + dataTable.Columns[j].ToString() + "</th>");  // hard-wire header alignment to centered
            }
            htmlDocument.AppendLine(" </tr>");

            for (int i = 0; i < dataTable.Rows.Count; i++)    //  <tr> <td>No ext.</td>      <td>RFMD assembly format</td> </tr>
            {
                htmlDocument.Append("<tr> ");
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    htmlDocument.Append("<td");
                    if (ColAlignments != null && ColAlignments[j] != null && ColAlignments[j] != null && ColAlignments[j] != "") htmlDocument.Append(" align=\"" + ColAlignments[j] + "\"");
                    if (ColorData != null && ColorData[j] != null && ColorData[j][i] != null && ColorData[j][i] != "") htmlDocument.Append(" bgcolor=\"" + ColorData[j][i] + "\"");
                    //     if (ColorData != null && ColorData[i] != null && ColorData[i][j] != null && ColorData[i][j] != null && ColorData[i][j] != "") htmlDocument.Append(" bgcolor=\"" + ColorData[i][j] + "\"");
                    htmlDocument.Append(">");

                    object element = dataTable.Rows[i][j];
                    Type typ = dataTable.Columns[j].DataType;
                    string str = "";
                    if (element != DBNull.Value)
                    {
                        if (typ == typeof(string))
                        {
                            str = (string)element;
                        }
                        else if (typ == typeof(double))
                        {
                            if (Format != null && j < Format.Count && Format[j] != null && Format[j].Length > 0)
                            {
                                str = ((double)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((double)element).ToString();
                            }
                        }
                        else if (typ == typeof(int))
                        {
                            if (Format != null && j < Format.Count && Format[j] != null && Format[j].Length > 0)
                            {
                                str = ((int)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((int)element).ToString();
                            }
                        }
                        else if (typ == typeof(float))
                        {
                            if (Format != null && j < Format.Count && Format[j] != null && Format[j].Length > 0)
                            {
                                str = ((float)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((float)element).ToString();
                            }
                        }
                        else if (typ == typeof(Int16))
                        {
                            if (Format != null && j < Format.Count && Format[j] != null && Format[j].Length > 0)
                            {
                                str = ((Int16)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((Int16)element).ToString();
                            }
                        }
                        else if (typ == typeof(decimal))
                        {
                            if (Format != null && j < Format.Count && Format[j] != null && Format[j].Length > 0)
                            {
                                str = ((decimal)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((decimal)element).ToString();
                            }
                        }
                        else if (typ == typeof(bool))
                        {
                            str = ((bool)element).ToString();
                        }
                    }
                    else  // object was empty
                    {
                        str = "";
                    }
                    htmlDocument.Append(str + "</td>");
                }
                htmlDocument.AppendLine(" </tr>");
            }

            htmlDocument.AppendLine("</table>");
            if (CenterFlag) htmlDocument.Append("</center>");
        }

        /// <summary>
        /// UNTESTED
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="Format"></param>
        /// <param name="ColAlignments"></param>
        /// <param name="BorderFlag"></param>
        /// <param name="CenterFlag"></param>
        public void Table(DataTable dataTable, IList<string> Format, IList<string> ColAlignments, bool BorderFlag, bool CenterFlag, string Class = null)
        {
            Table(dataTable, Format, ColAlignments, null, BorderFlag, CenterFlag, Class);
        }

        // this is to be the main datatable series
        public void Table(DataTable dataTable, Dictionary<string, string> FormatDict, Dictionary<string, string> ColAlignDict, bool BorderFlag, bool CenterFlag, string Class = null)
        {
            Table(dataTable, FormatDict, ColAlignDict, null, BorderFlag, CenterFlag, Class);

        }

        // this is to be the main datatable series
        public void Table(DataTable dataTable, Dictionary<string, string> FormatDict, Dictionary<string, string> ColAlignDict, DataTable ColorData, bool BorderFlag, bool CenterFlag, string Class = null)
        {
            if (dataTable == null || dataTable.Columns.Count == 0) return;

            if (CenterFlag) htmlDocument.Append("<center>");
            htmlDocument.Append("<table" + GetClass(Class)); if (BorderFlag) htmlDocument.Append(" border=\"1\""); htmlDocument.AppendLine(">");   // <table border="1">

            int Ncols = dataTable.Columns.Count;

            // preprocess Column alignments
            string[] ColAlignments = new string[Ncols];
            for (int j = 0; j < Ncols; j++)
            {
                ColAlignments[j] = "center";
                string colnam = dataTable.Columns[j].ColumnName;
                if (ColAlignDict != null && ColAlignDict.ContainsKey(colnam))
                {
                    string align = ColAlignDict[colnam];
                    if (align.Length > 0)
                    {
                        if (align.Substring(0, 1).ToUpper() == "L") align = "left";
                        else if (align.Substring(0, 1).ToUpper() == "R") align = "right";
                        else align = "center";
                    }
                    ColAlignments[j] = align;
                }
            }
            // preprocess format
            string[] Format = new string[Ncols];
            for (int j = 0; j < Ncols; j++)
            {
                Format[j] = null;
                string colnam = dataTable.Columns[j].ColumnName;
                if (FormatDict != null && FormatDict.ContainsKey(colnam))
                {
                    string fmt = FormatDict[colnam];
                    if (fmt.Length > 0) Format[j] = fmt;
                }
            }


            htmlDocument.Append("<tr> ");
            for (int j = 0; j < Ncols; j++)
            {
                //  htmlDocument.Append("<th align=\""+ColAlignments[j]+"\">"+ColHeaders[j]+"</th>");
                htmlDocument.Append("<th align=\"center\">" + dataTable.Columns[j].ToString() + "</th>");  // hard-wire header alignment to centered
            }
            htmlDocument.AppendLine(" </tr>");

            for (int i = 0; i < dataTable.Rows.Count; i++)    //  <tr> <td>No ext.</td>      <td>RFMD assembly format</td> </tr>
            {
                htmlDocument.Append("<tr> ");
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    string colnam = dataTable.Columns[j].ColumnName;
                    htmlDocument.Append("<td");
                    if (ColAlignments != null && ColAlignments[j] != null && ColAlignments[j] != null && ColAlignments[j] != "") htmlDocument.Append(" align=\"" + ColAlignments[j] + "\"");
                    if (ColorData != null && ColorData.Columns.Contains(colnam) && ColorData.Rows[i][colnam] != DBNull.Value && (string)ColorData.Rows[i][colnam] != "") htmlDocument.Append(" bgcolor=\"" + (string)ColorData.Rows[i][colnam] + "\"");
                    //     if (ColorData != null && ColorData[i] != null && ColorData[i][j] != null && ColorData[i][j] != null && ColorData[i][j] != "") htmlDocument.Append(" bgcolor=\"" + ColorData[i][j] + "\"");
                    htmlDocument.Append(">");

                    object element = dataTable.Rows[i][j];
                    Type typ = dataTable.Columns[j].DataType;
                    string str = "";
                    if (element != DBNull.Value)
                    {
                        if (typ == typeof(string))
                        {
                            str = (string)element;
                        }
                        else if (typ == typeof(double))
                        {
                            if (Format[j] == null || Format[j].Trim().ToUpper().StartsWith("N"))  // My "Nice" format e.g "N3" or ""
                            {
                                int digits = 3;                                                     // default is 3 dps
                                if (Format[j] != null && Format[j].Trim().Length > 1)
                                {
                                    digits = Convert.ToInt32(Format[j].Trim().Substring(1));
                                }
                                str = NiceFormat((double)element, digits);
                            }
                            else if (Format[j] != null)                                             // eg "F3" or "0.00"
                            {
                                str = ((double)element).ToString(Format[j]);
                            }
                            else                                                                    // default
                            {
                                str = ((double)element).ToString();
                            }
                        }
                        else if (typ == typeof(int))
                        {
                            if (Format[j] != null)
                            {
                                str = ((int)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((int)element).ToString();
                            }
                        }
                        else if (typ == typeof(float))
                        {
                            if (Format[j] != null)
                            {
                                str = ((float)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((float)element).ToString();
                            }
                        }
                        else if (typ == typeof(Int16))
                        {
                            if (Format[j] != null)
                            {
                                str = ((Int16)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((Int16)element).ToString();
                            }
                        }
                        else if (typ == typeof(decimal))
                        {
                            if (Format[j] != null)
                            {
                                str = ((decimal)element).ToString(Format[j]);
                            }
                            else
                            {
                                str = ((decimal)element).ToString();
                            }
                        }
                        else if (typ == typeof(bool))
                        {
                            str = ((bool)element).ToString();
                        }
                    }
                    else  // object was empty
                    {
                        str = "";
                    }
                    htmlDocument.Append(str + "</td>");
                }
                htmlDocument.AppendLine(" </tr>");
            }

            htmlDocument.AppendLine("</table>");
            if (CenterFlag) htmlDocument.Append("</center>");
        }

        /// <summary>
        /// OOPS - FILTERS OUT THE WAFERNO!
        /// Method for filtering HTML text created by HtmlBuilder class to keep only the lines in tables that are colour highlighted.
        /// Lines outside of tables are not filtered.
        /// For example to post-process html parameter tables to show only OOC table entries.
        /// Example of usage:  
        /// webBrowser.DocumentText = HtmlTableFilter(HtmlText, RemoveNonOocLinesFlag);
        /// </summary>
        /// <param name="HtmlText">The string containing the HTML</param>
        /// <param name="RemoveNonColouredLines">Set true to keep only highlighted </param>
        /// <param name="ColoursToRetain">defaults to Orange and Red</param>
        /// <param name="ReplacementMessages">Optional list of messges to replace any tables that have all rows
        ///                          filtered out. The index of the message list corresponds to the index of the
        ///                          tables in the Html document.</param>
        /// <param name="ReplacementMessagesFontSize">"1" to "7". e.g. "3" or "+1"</param>
        /// <returns></returns>
        public static string HtmlTableFilter(string HtmlText, bool RemoveNonColouredLines = true, IList<string> ColoursToRetain = null, IList<string> ReplacementMessages = null, string ReplacementMessagesFontSize = "3", bool ForceFirstColumnPopulated = true)
        {
            if (!RemoveNonColouredLines) return HtmlText;

            if (ColoursToRetain == null) ColoursToRetain = new List<string> { "Orange", "Red" };
            for (int i = 0; i < ColoursToRetain.Count; i++) { ColoursToRetain[i] = ColoursToRetain[i].ToLower(); }  // so not case sensitive

            if (ReplacementMessages == null) ReplacementMessages = new List<string>(0);

            StringBuilder sb = new StringBuilder();

            using (var sr = new StringReader(HtmlText))
            {
                string line;
                bool InTable = false;
                bool DidTableHaveRows = false;
                int TableCount = 0;
                string FirstColumnValue = null;  // For when FirstColumn is the WaferNo and is filtered out
                while ((line = sr.ReadLine()) != null)
                {
                    bool keepLine = true;

                    string lineLC = line.ToLower();
                    if (lineLC.Contains("<table")) { InTable = true; DidTableHaveRows = false; TableCount++; FirstColumnValue = null; continue; }
                    if (lineLC.Contains("</table"))
                    {
                        InTable = false;

                        // table had nothing left so substitue a message if provided
                        if (!DidTableHaveRows && TableCount <= ReplacementMessages.Count)
                        {//<font size="5" color="#ff0000">font size and color</font>
                            sb.AppendLine("<p><font size=\"" + ReplacementMessagesFontSize.Trim('"') + "\">" + ReplacementMessages[TableCount - 1] + "</font></p>");
                        }
                        continue;
                    }

                    if (InTable)
                    {
                        if (ForceFirstColumnPopulated && String.IsNullOrEmpty(FirstColumnValue))  // remember the first non-blank item of the table
                        {
                            string firstColValue = "";
                            if (lineLC.Trim().StartsWith("<tr")) firstColValue = HtmlTableRowFirstElement(lineLC);
                            if (!String.IsNullOrEmpty(firstColValue)) FirstColumnValue = firstColValue;
                        }

                        keepLine = false;
                        //if (lineLC.Contains("<td") && lineLC.Contains("bgcolor") && lineLC.Contains(ColourToRetain.ToLower())) keepLine = true;
                        if (lineLC.Contains("<td") && lineLC.Contains("bgcolor"))
                        {
                            foreach (var colour in ColoursToRetain)
                            {
                                if (lineLC.Contains(colour)) { keepLine = true; DidTableHaveRows = true; }
                            }
                        }

                        // make sure first column is populated - PROBABLY WANT TO MAKE JUST THE FIRST ONE OF THE SET??
                        // OR MAKE THE Html TABLE ON THE FLY FROM THE SUMMARY TABLE??
                        if (ForceFirstColumnPopulated && keepLine && HtmlTableRowFirstElement(lineLC).Trim().Length == 0)
                        {
                            line = HtmlTableRowInsertAtFirstElement(line, FirstColumnValue);
                        }
                    }

                    if (keepLine) sb.AppendLine(line);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// returns an array of elements parsed from an html [tr] row</tr>
        /// </summary>
        /// <param name="tableLine"></param>
        /// <returns></returns>
        static string[] HtmlTableRowSplit(string tableLine)
        {
            tableLine = tableLine.Replace("<tr>", "").Replace("</tr>", "").Replace("<th align=\"center\">", "|").Replace("</th>", "|");
            string[] parts = tableLine.Split('|');
            return parts;
        }

        /// <summary>
        /// returns the first element parse from an html [tr] row
        /// </summary>
        /// <param name="tableLine"></param>
        /// <returns>typically the WaferNo or blank</returns>
        static string HtmlTableRowFirstElement(string tableLine)
        {
            string[] parts = HtmlTableRowSplit(tableLine);
            if (parts.Length == 0) return "";
            return parts[0].Trim();
        }

        /// <summary>
        /// add the ItemToInsert text inb ta the end of the first item of the HTML table row
        /// </summary>
        /// <param name="tableLine"></param>
        /// <param name="ItemToInsert"></param>
        /// <returns></returns>
        static string HtmlTableRowInsertAtFirstElement(string tableLine, string ItemToInsert)
        {
            string[] parts = tableLine.Split(new string[] { "</th>" }, StringSplitOptions.None);
            if (parts.Length > 0)
            {
                parts[0] += ItemToInsert;
            }
            return String.Join("</th>", parts);
        }
        #endregion

        #region Images
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="Caption"></param>
        /// <param name="border"></param>
        /// <returns></returns>
        public bool Image(string filepath, int width, int height, string Caption, bool border, string Class = null)
        {
            int borderWid = 0; if (border) borderWid = 1;

            bool OK = false;
            if (File.Exists(filepath)) OK = true;
            /*
            <h2>Norwegian Mountain Trip</h2>
            <img border="0" src="/images/pulpit.jpg" alt="Pulpit rock" width="304" height="228"> */

            htmlDocument.AppendLine("<img" + GetClass(Class) + " border=\"" + borderWid.ToString() + "\" src=\"" + filepath + "\" alt=\"" + Caption + "\" width=\"" + width.ToString() +
                "\" height=\"" + height.ToString() + "\">");
            htmlDocument.AppendLine("<br>" + Caption + "</br>");

            return OK;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="width"></param>
        /// <param name="Caption"></param>
        /// <param name="border"></param>
        /// <returns></returns>
        public bool Image(string filepath, int width, string Caption, bool border, string Class = null)
        {
            int borderWid = 0; if (border) borderWid = 1;

            bool OK = false;
            if (File.Exists(filepath)) OK = true;
            /*
            <h2>Norwegian Mountain Trip</h2>
            <img border="0" src="/images/pulpit.jpg" alt="Pulpit rock" width="304" height="228"> */

            htmlDocument.AppendLine("<img" + GetClass(Class) + " border=\"" + borderWid.ToString() + "\" src=\"" + filepath + "\" alt=\"" + Caption + "\" width=\"" + width.ToString() +
                "\">");
            htmlDocument.AppendLine("<br>" + Caption + "</br>");

            return OK;
        }
        #endregion

        #region structures
        /// <summary>
        /// Create your very own div! Right here!
        /// </summary>
        /// <param name="contents"></param>
        public void Div(StringBuilder contents, string Class = null)
        {
            if (contents == null)
                htmlDocument.Append("<div" + GetClass(Class) + ">" + "</div>");
            else if (contents.ToString().ToLower().Contains("<div") && contents.ToString().ToLower().Contains("</div>"))
                htmlDocument.Append(contents);
            else
                htmlDocument.Append("<div" + GetClass(Class) + ">" + contents.ToString() + "</div>");
        }
        #endregion

        #region Helper_utilities/Formatting
        /// <summary>
        /// General list/listItem object
        /// </summary>
        public class ListObj
        {
            public string Text = null;
            public bool Numbered = false;
            public List<ListObj> Sublist = new List<ListObj>(4);

            public ListObj() { }

            public ListObj(string Text)
            {
                this.Text = Text;
            }

            public ListObj(string Text, List<ListObj> Sublist)
            {
                this.Text = Text;
                this.Sublist = Sublist;
            }

            public string GenerateHtml()   // recursive expansion of the list. NOT WORKING
            {
                string beg = "<ul>\r\n", end = "</ul>\r\n";
                if (Numbered) beg = "<ol>\r\n"; end = "</ol>\r\n";

                string ret = beg + Text + "\r\n";
                if (Sublist != null && Sublist.Count != 0)
                {
                    foreach (ListObj lit in Sublist)
                    {
                        ret += lit.GenerateHtml();
                    }
                }
                ret += end;
                return ret;
            }

        }

        public static ListObj MakeSubList(string[] Items, bool NumberedList)
        {
            ListObj lo = new ListObj();
            lo.Numbered = NumberedList;

            foreach (string itemstr in Items)
            {
                ListObj li = new ListObj();
                li.Text = itemstr;
                lo.Sublist.Add(li);
            }
            return lo;
        }

        public static ListObj MakeSublist(ListObj[] Items, bool NumberedList)
        {
            ListObj lo = new ListObj();
            lo.Numbered = NumberedList;
            foreach (ListObj item in Items)
            {
                lo.Sublist.Add(item);
            }
            return lo;
        }

        /// <summary>
        /// Displays HoverText when mouse hovers over the Text
        /// </summary>
        /// <param name="Text"></param>
        /// <param name="HoverText"></param>
        /// <returns></returns>
        public static string HoverText(string Text, string HoverText) //<abbr title="as soon as possible">ASAP</abbr>
        {
            return "<abbr title=\"" + HoverText + "\">" + Text + "</abbr>";
        }

        /// <summary>
        /// Returns a formatted hyperlink
        /// </summary>
        /// <param name="LinkText"></param>
        /// <param name="LinkUrlTarget"></param>
        /// <param name="NewWindow">set true to open the link in a new browser window</param>
        /// <returns></returns>
        public static string FormatHyperLink(string LinkText, string LinkUrlTarget, bool NewWindow) // <a href="url">Link text</a>  or <a href="url" target="_blank">Link Text</a> 
        {
            string result;
            if (!NewWindow)
                result = "<a href=\"" + LinkUrlTarget + "\">" + LinkText + "</a>";
            else
                result = "<a href=\"" + LinkUrlTarget + "\" target=\"_blank\">" + LinkText + "</a>";
            return result;
        }


        public static string FormatLocalLink(string LinkText, string LocalLabel, bool NewWindow) // <a href="#label">Link text</a>  or <a href="#label" target="_blank">Link Text</a> 
        {
            string result;
            if (!NewWindow)
                result = "<a href=\"#" + LocalLabel + "\">" + LinkText + "</a>";
            else
                result = "<a href=\"#" + LocalLabel + "\" target=\"_blank\">" + LinkText + "</a>";
            return result;
        }

        /// <summary>
        /// Examples: emailstr=FormatEmailLink("rob.davis@company.com"),  emailstr=FormatEmailLink("rob", "rob.davis@company.com", "subject string", "body string")
        /// 
        /// </summary>
        /// <param name="LinkText">set null for LinkText to inherit Email Address</param>
        /// <param name="EmailAddress">Set null for Email address to inherit LinkText [Default]</param>
        /// <param name="subject">set blank for no subject [Default]</param>
        /// <param name="body">set blank for no body [Default]</param>
        /// <returns></returns>
        public static string FormatEmailLink(string LinkText, string EmailAddress = null, string subject = "", string body = "") // <a href="url">Link text</a>  or <a href="url" target="_blank">Link Text</a> 
        {
            if (subject == null) subject = "";
            if (subject != "") subject = "?subject=" + subject;

            if (body == null) body = "";
            if (body != "") body = "?body=" + body;

            if (EmailAddress == null) EmailAddress = LinkText;
            if (LinkText == null) LinkText = EmailAddress;

            string result;
            result = "<a href=\"mailto:" + EmailAddress + subject + body + "\">" + LinkText + "</a>";
            return result;
        }

        /// <summary>
        /// replaces all instances of a placeholder keys in the document with the specified text
        /// </summary>
        /// <param name="LPlaceholderSubstitutePairs">List of string[] {placeholder, subsitute} keys</param>
        public void PlaceholderReplace(List<string[]> LPlaceholderSubstitutePairs)
        {
            foreach (string[] pair in LPlaceholderSubstitutePairs)
            {
                if (pair != null && pair.Length > 1)
                {
                    htmlDocument.Replace(pair[0], pair[1]);
                }
            }
        }


        public static string it(string Text)
        {
            return "<it>" + Text + "</it>";
        }
        public static string Italics(string Text)
        {
            return "<it>" + Text + "</it>";
        }

        public static string b(string Text)
        {
            return "<b>" + Text + "</b>";
        }
        public static string Bold(string Text)
        {
            return "<b>" + Text + "</b>";
        }
        public static string Sub(string Text)
        {
            return "<sub>" + Text + "</sub>";
        }
        public static string Sup(string Text)
        {
            return "<sup>" + Text + "</sup>";
        }
        public static string Small(string Text)
        {
            return "<small>" + Text + "</small>";
        }

        static string NiceFormat(double val, int ndps)
        {
            double MinDpVal = Math.Pow(10, -(ndps - 1));

            string fmt = "";
            if (Math.Abs(val) > MinDpVal && Math.Abs(val) < 1e4)
            {
                fmt = "f" + ndps.ToString();
                string sval = val.ToString(fmt);

                sval = sval.TrimEnd(new char[] { '0' });
                sval = sval.TrimEnd(new char[] { '.' });
                if (sval.Length == 0) sval = "0";
                return sval;
            }
            else if (Math.Abs(val) < 10 * Double.Epsilon)
                return "0";
            else
            {
                fmt = "0.";
                for (int i = 0; i < ndps; i++)
                {
                    fmt += "0";
                }
                fmt += "e00";
                string sval = val.ToString(fmt);

                string[] parts = sval.Split(new char[] { 'e' });
                parts[0] = parts[0].TrimEnd(new char[] { '0' });
                parts[0] = parts[0].TrimEnd(new char[] { '.' });
                if (parts.Length > 1)
                    return parts[0] + "e" + parts[1];
                else
                    return parts[0];    // can happen if value is NaN
            }

        }

        public void Preformatted(string text)
        {
            if (text == null) return;
            htmlDocument.AppendLine("<pre>");
            htmlDocument.AppendLine(text);
            htmlDocument.AppendLine("</pre>");
        }
        #endregion utilities

        #region Scripting and Styling
        /// <summary>
        /// Add a custom js script to the webpage.
        /// </summary>
        /// <param name="script"></param>
        public void Script(StringBuilder script)
        {
            if (script.ToString().ToLower().Contains("<script>") && script.ToString().ToLower().Contains("</script>"))
                htmlDocument.Append(script);
            else
                htmlDocument.Append("<script>" + script.ToString() + "</script>");
        }

        public void Script(string script)
        {
            if (script.ToLower().Contains("<script>") && script.ToLower().Contains("</script>"))
                htmlDocument.Append(script);
            else
                htmlDocument.Append("<script>" + script + "</script>");
        }

        /// <summary>
        /// Add a custom css style to the webpage.
        /// </summary>
        /// <param name="style"></param>
        public void Style(StringBuilder style)
        {
            if (style.ToString().ToLower().Contains("<style>") && style.ToString().ToLower().Contains("</style>"))
                htmlDocument.Append(style);
            else
                htmlDocument.Append("<style>" + style.ToString() + "</style>");
        }

        public void Style(string style)
        {
            if (style.ToLower().Contains("<style>") && style.ToLower().Contains("</style>"))
                htmlDocument.Append(style);
            else
                htmlDocument.Append("<style>" + style + "</style>");
        }

        /// <summary>
        /// Get the class to add to each tag, or nothing if null
        /// </summary>
        /// <param name="Class"></param>
        /// <returns></returns>
        public string GetClass(string Class)
        {
            if (Class == null) return "";
            else
                return " class=\"" + Class + "\"";
        }
        #endregion
    }
}