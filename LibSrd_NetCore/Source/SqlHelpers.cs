using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace LibSrd_NETCore
{
    public class SqlHelpers
    {
        #region Delete

        /// <summary>
        /// Deletes rows according to dictionary of WHERE conditions 
        /// </summary>
        /// <param name="sqlTableName"></param>
        /// <param name="ColName_ColValueEqual"></param>
        /// <param name="ColName_ColValueIn"></param>
        /// <param name="sqlConn"></param>
        /// <param name="sqlTimeout_s"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static int sqlDeleteRows(string sqlTableName, Dictionary<string, string> ColName_ColValueEqual, Dictionary<string, string> ColName_ColValueIn, SqlConnection sqlConn, int sqlTimeout_s, out string ErrMsg) //Delete matching entries in Sql (Metadata) DB
        {
            ErrMsg = null;
            if (ColName_ColValueEqual == null && ColName_ColValueIn == null) return 0;

            // string sqlStr = "DELETE FROM " + TableName + " WHERE (LotNo = @LotNo AND WaferNo IN ( @WaferList ) AND Route = @Route AND TestStage = @TestStage AND PartNo = @PartNo)";
            List<string> whereStrList = new List<string>();
            if (ColName_ColValueEqual != null)
            {
                foreach (var key in ColName_ColValueEqual.Keys)
                {
                    whereStrList.Add(key + "= " + ColName_ColValueEqual[key]);
                }
            }
            if (ColName_ColValueIn != null)
            {
                foreach (var key in ColName_ColValueIn.Keys)
                {
                    whereStrList.Add(key + " in (" + ColName_ColValueIn[key] + ")");
                }
            }
            string whereStr = " WHERE " + String.Join(" AND ", whereStrList);

            string sqlStr = "DELETE FROM " + sqlTableName + whereStr;

            int RowsDeleted = 0;
            try
            {
                SqlCommand sqcmd = new SqlCommand(sqlStr, sqlConn);

                sqcmd.CommandTimeout = sqlTimeout_s; // 300; // 120;  // default seems to be 30 sec, 120 was failing for 58040 deletes 

                RowsDeleted = sqcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                return -1;
            }
            return RowsDeleted;
        }

        /// <summary>
        /// Method to delete rows from a Sql table. Rows that match the criteria in the MatchingRow DataTable are deleted.
        /// </summary>
        /// <param name="sqlTableName"></param>
        /// <param name="MatchingRowTable">Any SQL rows that match any of the rows in MatchingRowTable are deleted.
        ///                               (A match occurs when all the values *in the provided* columns are equal). 
        ///                                A convenient way to generate such a table could be something like 
        ///                                DataTable dtDistinct = dataView.ToTable(true,"PFRNo","TestStage"); 
        /// </param>
        /// <param name="sqlConn"></param>
        /// <param name="sqlTimeout_s"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static int sqlDeleteRows(string sqlTableName, DataTable MatchingRowTable, SqlConnection sqlConn, int sqlTimeout_s, out string ErrMsg) //Delete matching entries in Sql (Metadata) DB
        {
            ErrMsg = null;
            if (MatchingRowTable == null || MatchingRowTable.Rows.Count == 0) return 0;

            // string sqlStr = "DELETE FROM " + TableName + " WHERE (LotNo = @LotNo AND WaferNo IN ( @WaferList ) AND Route = @Route AND TestStage = @TestStage AND PartNo = @PartNo)";

            List<List<string>> whereOrList = new List<List<string>>();
            foreach (DataRow row in MatchingRowTable.Rows)
            {
                List<string> whereAndList = new List<string>();
                foreach (DataColumn col in MatchingRowTable.Columns)
                {
                    if (row[col] != DBNull.Value)
                    {
                        string name = col.ColumnName;
                        object obj = Convert.ChangeType(row[col], col.DataType);

                        string value;
                        if (col.DataType != typeof(DateTime))
                            value = Convert.ChangeType(row[col], col.DataType).ToString();
                        else
                            value = DateTime2SqlDateTime((DateTime)row[col]);

                        if (col.DataType == typeof(string) || col.DataType == typeof(DateTime) || col.DataType == typeof(TimeSpan)) value = "'" + value + "'";
                        whereAndList.Add(name + "=" + value);
                    }
                }
                whereOrList.Add(whereAndList);
            }


            List<string> whereList = new List<string>();
            foreach (var andList in whereOrList)
            {
                if (andList.Count > 0)
                    whereList.Add("(" + String.Join(" AND ", andList) + ")");
            }

            string whereStr = " WHERE " + String.Join(" OR ", whereList);


            string sqlStr = "DELETE FROM " + sqlTableName + whereStr;

            int RowsDeleted = 0;
            try
            {
                SqlCommand sqcmd = new SqlCommand(sqlStr, sqlConn);

                sqcmd.CommandTimeout = sqlTimeout_s; // 300; // 120;  // default seems to be 30 sec, 120 was failing for 58040 deletes 

                RowsDeleted = sqcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                return -1;
            }
            return RowsDeleted;
        }

        /// <summary>
        /// Converts a DateTime to SQL format
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static string DateTime2SqlDateTime(DateTime dateTime, bool Includefff = false)
        {
            if (!Includefff)
                return dateTime.ToString("yyyy-MM-ddTHH:mm:ss");  // HH for 24 hr
            else
                return dateTime.ToString("yyyy-MM-ddTHH:mm:ss:fff");  // HH for 24 hr
        }

        static int sqlDeleteRowsNOTWORKING(string TableName, Dictionary<string, string> ColName_ColValueEqual, Dictionary<string, string> ColName_ColValueIn, SqlConnection sqlConn, int sqlTimeout_s, out string ErrMsg) //Delete matching entries in Sql (Metadata) DB
        {
            ErrMsg = null;
            if (ColName_ColValueEqual == null && ColName_ColValueIn == null) return 0;

            // string sqlStr = "DELETE FROM " + TableName + " WHERE (LotNo = @LotNo AND WaferNo IN ( @WaferList ) AND Route = @Route AND TestStage = @TestStage AND PartNo = @PartNo)";
            List<string> whereStrList = new List<string>();
            if (ColName_ColValueEqual != null)
            {
                foreach (var key in ColName_ColValueEqual.Keys)
                {
                    whereStrList.Add(key + "= @" + key);
                }
            }
            if (ColName_ColValueIn != null)
            {
                foreach (var key in ColName_ColValueIn.Keys)
                {
                    whereStrList.Add(key + " in (@" + key + ")");
                }
            }
            string whereStr = " WHERE " + String.Join(" AND ", whereStrList);

            string sqlStr = "DELETE FROM " + TableName + whereStr;

            int RowsDeleted = 0;
            try
            {
                SqlCommand sqcmd = new SqlCommand(sqlStr, sqlConn);

                sqcmd.CommandTimeout = sqlTimeout_s; // 300; // 120;  // default seems to be 30 sec, 120 was failing for 58040 deletes 

                if (ColName_ColValueEqual != null)
                {
                    foreach (var key in ColName_ColValueEqual.Keys)
                    {
                        sqcmd.Parameters.AddWithValue("@" + key, ColName_ColValueEqual[key]);
                    }
                }
                if (ColName_ColValueIn != null)
                {
                    foreach (var key in ColName_ColValueIn.Keys)
                    {
                        sqcmd.Parameters.AddWithValue("@" + key, ColName_ColValueIn[key]);
                    }
                }

                sqcmd.CommandType = CommandType.Text;
                RowsDeleted = sqcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                return -1;
            }
            return RowsDeleted;
        }

        #endregion Delete

        #region Create Helpers
        /// <summary>
        /// Creates a sql Table to match a DataTable
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="TableName">Defaults to dbo.TableName</param>
        /// <param name="DefaultVarcharSize">default VARCHAR size. Fails over to MAX if LE 0</param>
        /// <param name="PredfinedColumnTypesDict">Predefined SQL column types dict override</param>
        /// <param name="ConnString"></param>
        /// <param name="ErrMsg"></param>
        public static void sqlTableCreateMatchingTable(DataTable dt, string TableName, int DefaultVarcharSize, Dictionary<string, string> PredfinedColumnTypesDict, string ConnString, out string ErrMsg)
        {
            ErrMsg = null;
            try
            {
                using (SqlConnection sqlConn = new SqlConnection(ConnString))
                {
                    sqlConn.Open();
                    sqlTableCreateMatchingTable(dt, TableName, DefaultVarcharSize, PredfinedColumnTypesDict, sqlConn, out ErrMsg);
                }
            }
            catch (Exception ex)
            {
                ErrMsg = "Error creating Sql table: " + ex.Message;
            }
        }

        /// <summary>
        /// Creates a sql Table to match a DataTable
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sqlTableName">Defaults to dbo.TableName</param>
        /// <param name="DefaultVarcharSize">default VARCHAR size. Fails over to MAX if LE 0</param>
        /// <param name="PredfinedColumnTypesDict">Predefined SQL column types dict override</param>
        /// <param name="sqlConn"></param>
        /// <param name="ErrMsg"></param>
        public static void sqlTableCreateMatchingTable(DataTable dt, string sqlTableName, int DefaultVarcharSize, Dictionary<string, string> PredfinedColumnTypesDict, SqlConnection sqlConn, out string ErrMsg)
        {
            ErrMsg = null;
            string sqlCmd = sqlTableCreate_GetCmdMatchingTable(dt, sqlTableName, DefaultVarcharSize, PredfinedColumnTypesDict);

            if (String.IsNullOrEmpty(sqlCmd))
            {
                ErrMsg = "Error, cannot make table " + sqlTableName + " to match DataTable as the table is invalid (null or no columns)";
                return;
            }

            try
            {
                using (SqlCommand sqlQuery = new SqlCommand(sqlCmd, sqlConn))
                {
                    SqlDataReader reader = sqlQuery.ExecuteReader();
                    reader.Close();
                }
            }
            catch (Exception ex)
            {
                ErrMsg = "Error creating Sql table: " + ex.Message;
            }
        }

        /// <summary>
        /// Builds a sql command string to make a SQL table that matches a DataTable
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sqlTableName">if not supplied then defaults to dbo.{dt.TableName}</param>
        /// <param name="DefaultVarcharSize">the default VARCHAR size where not in Predefined dict. Fails over to MAX if LE 0</param>
        /// <param name="PredfinedColumnTypesDict"></param>
        /// <returns>Command to make the string, null on error</returns>
        public static string sqlTableCreate_GetCmdMatchingTable(DataTable dt, string sqlTableName, int DefaultVarcharSize, Dictionary<string, string> PredfinedColumnTypesDict)
        {
            if (dt == null || dt.Columns.Count == 0) return null;

            if (String.IsNullOrEmpty(sqlTableName)) sqlTableName = dt.TableName;
            if (!sqlTableName.Contains('.')) sqlTableName = "dbo." + sqlTableName;

            StringBuilder sqCmd = new StringBuilder();

            sqCmd.Append("CREATE TABLE " + sqlTableName + "(");

            for (int j = 0; j < dt.Columns.Count; j++)
            {
                string colName = dt.Columns[j].ColumnName.Trim(new char[] { '[', ']' });
                Type colType = dt.Columns[j].DataType;

                string colTypeString = sqlTableCreate_GetMatchingType(colName, colType, DefaultVarcharSize, PredfinedColumnTypesDict);

                sqCmd.Append("[" + colName + "] " + colTypeString);
                if (j < dt.Columns.Count - 1) sqCmd.Append(", ");
            }
            sqCmd.Append(");");
            return sqCmd.ToString();
        }

        /// <summary>
        /// Gets the sql column type string according to .NET type. If the type is predefined in the Predefinedict then use that one 
        /// </summary>
        /// <param name="ColumnName"></param>
        /// <param name="type"></param>
        /// <param name="DefaultVarcharSize">efault VARCHAR size. Fails over to MAX if LE 0</param>
        /// <param name="PredfinedDict">if null then some "standard" columns are added</param>
        /// <returns></returns>
        static string sqlTableCreate_GetMatchingType(string ColumnName, Type type, int DefaultVarcharSize, Dictionary<string, string> PredfinedDict)
        {
            #region Default SQL Column Types
            if (PredfinedDict == null)
            {
                PredfinedDict = new Dictionary<string, string>();
                PredfinedDict.Add("Comment", "VARCHAR(256)");
                PredfinedDict.Add("LotNo", "VARCHAR(24)");
                PredfinedDict.Add("LotNo2", "VARCHAR(24)");
                PredfinedDict.Add("LotNo3", "VARCHAR(24)");
                PredfinedDict.Add("WaferNo", "SMALLINT");
                PredfinedDict.Add("Shot", "VARCHAR(8)");
                PredfinedDict.Add("BarNo", "SMALLINT");
                PredfinedDict.Add("DieNo", "SMALLINT");
                PredfinedDict.Add("SeriaNo", "INT");
                PredfinedDict.Add("PartNo", "VARCHAR(32)");
                PredfinedDict.Add("TestStage", "VARCHAR(8)");
                PredfinedDict.Add("X", "SMALLINT");
                PredfinedDict.Add("Y", "SMALLINT");
            }
            #endregion Default SQL Column Types

            foreach (string key in PredfinedDict.Keys)
            {
                if (key.Trim().ToLower() == ColumnName.Trim().ToLower())
                {
                    return " " + PredfinedDict[key].Trim();
                }
            }

            return sqlTableCreate_GetMatchingType(type, DefaultVarcharSize);
        }

        /// <summary>
        /// gets the sql column type corresponding to the .NET type. If varchar then uses the size given
        /// </summary>
        /// <param name="type"></param>
        /// <param name="VarcharSize">Default VARCHAR size. Fails over to MAX if LE 0</param>
        /// <returns></returns>
        static string sqlTableCreate_GetMatchingType(Type type, int VarcharSize)
        {
            string varcharStr = "MAX";
            if (VarcharSize > 0) varcharStr = VarcharSize.ToString();

            if (type == typeof(string))
            {
                return " VARCHAR(" + varcharStr + ")";
            }
            else if (type == typeof(double) || type == typeof(float))
            {
                return " REAL";
            }
            else if (type == typeof(int))
            {
                return " INT";
            }
            else if (type == typeof(Int64))
            {
                return " BIGINT";
            }
            else if (type == typeof(Int16))
            {
                return " SMALLINT";
            }
            else if (type == typeof(bool))
            {
                return " BIT";
            }
            else if (type == typeof(DateTime))
            {
                return " DATETIME";
            }
            else if (type == typeof(Guid))
            {
                return " UNIQUEIDENTIFIER";
            }
            else
            {
                return " VARCHAR(" + varcharStr + ")";
            }
            //return "ERROR, UNKNOWN TYPE PROVIDED TO sqlMatchingType";
        }

        #endregion Create Helpers

        #region Execute

        public static int sqlExecuteNonQuery(string SqlCommandStr, int Timeout_s, string ConnectionString, out string ErrMsg)
        {
            int nrows = -1;
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                nrows = sqlExecuteNonQuery(SqlCommandStr, Timeout_s, conn, out ErrMsg);
            }
            return nrows;
        }

        /// <summary>
        /// Executes a non query command
        /// </summary>
        /// <param name="SqlCommandStr">e.g. "DELETE FROM " + tableDict["Meta"] + " WHERE ...";</param>
        /// <param name="Timeout_s">command timeout</param>
        /// <param name="sqlConn"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static int sqlExecuteNonQuery(string SqlCommandStr, int Timeout_s, SqlConnection sqlConn, out string ErrMsg)
        {
            int RowsAffected = 0;
            try
            {
                SqlCommand sqcmd = new SqlCommand(SqlCommandStr, sqlConn);

                sqcmd.CommandTimeout = Timeout_s; // 300; // 120;  // default seems to be 30 sec, 120 was failing for 58040 deletes 
                sqcmd.CommandType = CommandType.Text;
                RowsAffected = sqcmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                return -1;
            }
            ErrMsg = null;
            return RowsAffected;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="SqlCommandStr">e.g SELECT MIN(Colnam) FROM table</param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static object sqlExecuteScalar(string SqlCommandStr, SqlConnection conn)
        {
            SqlCommand SqCmd = new SqlCommand(SqlCommandStr, conn);

            object value = SqCmd.ExecuteScalar();

            return value;
        }

        #endregion Execute

        #region Upload
        public static Result sqlUpload(DataTable dt, string TableName, string ConnectionString, SqlBulkCopyOptions copyOptions, int Timeout_s)
        {
            string ErrorMsg;
            sqlUpload(dt, TableName, ConnectionString, copyOptions, Timeout_s, null, out ErrorMsg);
            if (String.IsNullOrEmpty(ErrorMsg))
                return Result.Ok();
            else
                return Result.Fail(ErrorMsg);
        }

        /// <summary>
        /// Manages bulkcopy of a DataTable. Uses ConnectionString rather than a Sql connection so as to allow the use of
        /// SqlBulkCopyOptions. A possible use of the latter is: SqlBullCopyOptions.TableLock | SqlBullCopyOptions.KeepIdentity
        /// for a big table where the identity is being provided from the datatable rather than being autoincremented.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="TableName">Defaults to dbo.TableName</param>
        /// <param name="ConnectionString"></param>
        /// <param name="copyOptions">e.g. SqlBulkCopyOptions.Default</param>
        /// <param name="Timeout_s"></param>
        /// <param name="ErrorMsg"></param>
        public static void sqlUpload(DataTable dt, string TableName, string ConnectionString, SqlBulkCopyOptions copyOptions, int Timeout_s, out string ErrorMsg)
        {
            sqlUpload(dt, TableName, ConnectionString, copyOptions, Timeout_s, null, out ErrorMsg);
        }

        /// <summary>
        /// Manages bulkcopy of a DataTable. Uses ConnectionString rather than a Sql connection so as to allow the use of
        /// SqlBulkCopyOptions. A possible use of the latter is: SqlBullCopyOptions.TableLock | SqlBullCopyOptions.KeepIdentity
        /// for a big table where the identity is being provided from the datatable rather than being autoincremented.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="TableName">Defaults to dbo.TableName</param>
        /// <param name="ConnectionString"></param>
        /// <param name="copyOptions">e.g. SqlBulkCopyOptions.Default</param>
        /// <param name="Timeout_s"></param>
        /// <param name="VarcharLenDict">Optionally allows trimming of string entries to a specified string length: e.g. VarcharLenDict.Add(colnam,256). 
        ///                              Column name keys are case sensitive</param>
        /// <param name="ErrorMsg"></param>
        public static void sqlUpload(DataTable dt, string TableName, string ConnectionString, SqlBulkCopyOptions copyOptions, int Timeout_s, Dictionary<string, int> VarcharLenDict, out string ErrorMsg)
        {
            ErrorMsg = "";
            if (Timeout_s < 0) Timeout_s = 600;

            if (String.IsNullOrEmpty(TableName)) TableName = dt.TableName;
            if (!TableName.Contains('.')) TableName = "dbo." + TableName;

            // get sql columns
            List<string> SqlColumns = null;
            List<string> SqlColumnsLC = null;
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    conn.Open();
                    SqlColumns = sqlGetTableColumnNames(TableName, conn);
                }
                SqlColumnsLC = SqlColumns.ConvertAll(d => d.ToLower());
            }
            catch (Exception ex)
            {
                ErrorMsg = ex.Message;
                return;
            }

            #region ensure stored data is in single precision range
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                if (dt.Columns[j].DataType == typeof(double))
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i][j] != DBNull.Value) // can be null e.g. PM057.0003_01_0_FMA3010_100313-014435 summary table (Std dev when 
                        {
                            if ((double)dt.Rows[i][j] < -1e38)
                                dt.Rows[i][j] = -1e38;
                            else if ((double)dt.Rows[i][j] > 1e38)
                                dt.Rows[i][j] = 1e38;
                        }

                    }
                }
                // trim the text length to be less than the specified value in VarcharLenDict if required.
                else if (VarcharLenDict != null && dt.Columns[j].DataType == typeof(string) && VarcharLenDict.ContainsKey(dt.Columns[j].ColumnName))
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i][j] != DBNull.Value)
                        {
                            string str = (string)dt.Rows[i][j];
                            if (str.Length > VarcharLenDict[dt.Columns[j].ColumnName]) dt.Rows[i][j] = str.Substring(0, VarcharLenDict[dt.Columns[j].ColumnName]);
                        }
                    }
                }
            }
            #endregion ensure stored data is in single precision range

            using (SqlBulkCopy bc = new SqlBulkCopy(ConnectionString, copyOptions))
            {
                bc.BulkCopyTimeout = Timeout_s; // 600; //120;   // default is 30 sec

                bc.DestinationTableName = TableName;

                // Option to mapping the column names so column order doesnt have to match and can use Sql Identity option for automatic 
                // generation of integer primary key
                // Column names can have case difference
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string nam = dt.Columns[j].ColumnName;
                    //bc.ColumnMappings.Add(new SqlBulkCopyColumnMapping(nam, nam));  // map nam in dt onto nam in SQL table IF THE COLUMN EXISTS IN THE SQL TABLE

                    if (SqlColumnsLC.Contains(nam.ToLower()))
                    {
                        int inamSql = SqlColumnsLC.IndexOf(nam.ToLower());
                        bc.ColumnMappings.Add(new SqlBulkCopyColumnMapping(nam, SqlColumns[inamSql]));  // map nam in dt onto nam in SQL table IF THE COLUMN EXISTS IN THE SQL TABLE
                    }
                }

                try
                {
                    bc.WriteToServer(dt);
                }
                catch (Exception ex)
                {
                    ErrorMsg = ex.Message;
                }
            }
        }
        #endregion Upload

        #region Misc
        /// <summary>
        /// Test if Tablename is listed in information_schema.tables
        /// </summary>
        /// <param name="TableName"></param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static bool sqlGetTableExists(string TableName, SqlConnection conn)
        {
            // strip of any schema
            if (TableName.Contains('.'))
                TableName = TableName.Substring(TableName.IndexOf('.') + 1);

            TableName = TableName.Trim(new char[] { ' ', '[', ']' });

            //     List<string> list = new List<string>();

            string cmd = "select case when exists((select * from information_schema.tables where table_name = '" + TableName + "')) then 1 else 0 end";

            SqlCommand SqCmd = new SqlCommand(cmd, conn);

            bool exists = ((int)SqCmd.ExecuteScalar() == 1);

            return exists;
        }

        /// <summary>
        /// Get the max (int) value in a column
        /// </summary>
        /// <param name="TableName"></param>
        /// <param name="ColumnName"></param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static int sqlGetMaxInt(string TableName, string ColumnName, SqlConnection conn)
        {
            object val = sqlGetFunctionValue("MAX", TableName, ColumnName, conn);
            try
            {
                return (int)val;
            }
            catch
            {
                return -Int32.MaxValue;
            }
        }



        /// <summary>
        /// "SELECT " + FunctionName + "(" + ColumnName + ") FROM " + TableName + ";";
        /// </summary>
        /// <param name="FunctionName"></param>
        /// <param name="TableName"></param>
        /// <param name="ColumnName"></param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static object sqlGetFunctionValue(string FunctionName, string TableName, string ColumnName, SqlConnection conn)
        {

            string cmd = "SELECT " + FunctionName + "(" + ColumnName + ") FROM " + TableName + ";";

            SqlCommand SqCmd = new SqlCommand(cmd, conn);

            object value = SqCmd.ExecuteScalar();

            return value;
        }
        #endregion Misc

        #region Query to DataTable
        /// <summary>
        /// General sql query to Datatable method
        /// </summary>
        /// <param name="dt">Created automatically if null, or uses the structure provided</param>
        /// <param name="query"></param>
        /// <param name="ConnString">Sql Server connection string</param>
        /// <param name="ErrMsg">Length zero means no error</param>
        /// <returns></returns>
        public static int sqlQueryToDt(ref DataTable dt, string query, string ConnString, out string ErrMsg, int CommandTimeout = -1)  // General sql query to Datatable method
        {
            ErrMsg = null;
            int Nrows = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnString))
                {
                    conn.Open();
                    Nrows = sqlQueryToDt(ref dt, query, conn, out ErrMsg, CommandTimeout);
                }
            }
            catch (Exception ex)
            {
                Nrows = 0;
                ErrMsg = ex.Message;
            }
            return Nrows;
        }

        public static Result sqlQueryToDt(ref DataTable dt, string query, string ConnString, int CommandTimeout = -1)  // General sql query to Datatable method
        {
            string ErrMsg = null;
            int Nrows = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnString))
                {
                    conn.Open();
                    Nrows = sqlQueryToDt(ref dt, query, conn, out ErrMsg, CommandTimeout);
                }
            }
            catch (Exception ex)
            {
                Nrows = 0;
                return Result.Fail<int>(ex.Message, -1);

            }
            return Result.Ok<int>(Nrows);
        }

        /// <summary>
        /// General sql query to Datatable method
        /// </summary>
        /// <param name="dt">Created automaically if null, or uses the structure provided</param>
        /// <param name="sqlQuery"></param>
        /// <param name="sqlConn"></param>
        /// <param name="ErrMsg">Length zero means no error</param>
        /// <returns>The number of rows returned</returns>
        public static int sqlQueryToDt(ref DataTable dt, string sqlQuery, SqlConnection sqlConn, out string ErrMsg, int CommandTimeout = -1) // General sql query to Datatable method
        {
            if (dt == null) dt = new DataTable();
            ErrMsg = null;
            int Nrows = 0;
            try
            {
                using (SqlCommand sqlCmd = new SqlCommand(sqlQuery, sqlConn))
                {
                    if (CommandTimeout >= 0) sqlCmd.CommandTimeout = CommandTimeout; // 9/9/19

                    // create data adapter
                    SqlDataAdapter da = new SqlDataAdapter(sqlCmd);

                    // this will query your database and return the result to your datatable
                    Nrows = da.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
            }
            return Nrows;
        }


        /// <summary>
        /// This is normally the one to use
        /// General sql query to Datatable method that automatically sorts out the datatable definition to match the SQL
        /// source table but adjusts float and Int16/byte types to be double and Int32 so one can easily a maintain standard
        /// set of DataTypes in one's program!
        /// Updated version that always replicates the case of the column names
        /// 13.4.19 Updated again to cope with other statements before the select. Now modifies the first instance of select. 
        /// 11.9.19 Updated again to cope with UNION
        /// </summary>
        /// <param name="query">Sql query</param>
        /// <param name="ConnString">Database connection string</param>
        /// <param name="dt">This version has the DataTable as an output parameter as it is always dertermined within this method</param>
        /// <param name="ErrMsg">Returned error if the query fails</param>
        /// <returns>No. of rows returned</returns>
        public static int sqlQueryToDt(string query, string ConnString, out DataTable dt, out string ErrMsg, int CommandTimeout = -1)  // Comprehensive sql query to Datatable method
        {
            ErrMsg = null;
            int Nrows = 0;
            dt = new DataTable();
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnString))
                {
                    conn.Open();

                    // 4th refinement - lose all but the first query in the event of a UNION  11/9/19
                    string query2 = query;
                    int iUNION = query.IndexOf("UNION", StringComparison.OrdinalIgnoreCase);
                    if (iUNION > -1) query2 = query.Substring(0, iUNION);


                    // 1st go so simply cope with DISTINCT being present
                    // string moddedQuery = query.ToUpper().Replace("DISTINCT", "").Replace("SELECT", "SELECT TOP 0");

                    // 2nd refinement - to better cope with DISTINCT being present
                    var regexDistinct = new System.Text.RegularExpressions.Regex("DISTINCT", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    var regexSelect = new System.Text.RegularExpressions.Regex("SELECT", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    string moddedQuery = regexDistinct.Replace(query2, "");   // change "DISTINCT" to ""  irrespective of case
                    moddedQuery = regexSelect.Replace(moddedQuery, "SELECT TOP 0");   // change "SELECT" to "SELECT TOP 0" irrespctive of case

                    // 3rd refinement - to cope with both DISTINCT and TOP 
                    string cQuery = CollapseSpacesOutsideOfQuotedStrings(query2, new char[] { '\'' });
                    List<string> parts = cQuery.Split(' ').ToList();

                    // find first select
                    int iSelect = 0;
                    for (int i = 0; i < parts.Count; i++) if (parts[i].ToUpper() == "SELECT") { iSelect = i; break; }

                    if (parts.Count > iSelect && parts[iSelect + 1].ToUpper() == "DISTINCT") parts.RemoveAt(iSelect + 1);
                    if (parts.Count > iSelect && parts[iSelect + 1].ToUpper() == "TOP") parts.RemoveRange(iSelect + 1, 2);
                    parts[iSelect] = "SELECT TOP 0";
                    moddedQuery = string.Join(" ", parts);

                    Nrows = SqlHelpers.sqlQueryToDt(ref dt, moddedQuery, conn, out ErrMsg); // modify the SELECT clause to return no rows
                    if (!String.IsNullOrEmpty(ErrMsg)) return 0;

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (dt.Columns[j].DataType == typeof(float)) dt.Columns[j].DataType = typeof(double);
                        if (dt.Columns[j].DataType == typeof(Int16)) dt.Columns[j].DataType = typeof(int);
                        if (dt.Columns[j].DataType == typeof(byte)) dt.Columns[j].DataType = typeof(int);
                    }

                    Nrows = SqlHelpers.sqlQueryToDt(ref dt, query, conn, out ErrMsg, CommandTimeout);
                    if (!String.IsNullOrEmpty(ErrMsg)) return 0;
                }
            }
            catch (Exception ex)
            {
                Nrows = 0;
                ErrMsg = ex.Message;
            }
            return Nrows;
        }

        static string CollapseSpacesOutsideOfQuotedStrings(string text, char[] quotesChars = null)
        {
            StringBuilder sb = new StringBuilder(text.Length);
            bool quotesOpen = false;
            if (quotesChars == null) quotesChars = new char[] { '\'', '"' };

            for (int i = 0; i < text.Length - 1; i++)
            {
                char thisChar = Convert.ToChar(text.Substring(i, 1));
                char nextChar = Convert.ToChar(text.Substring(i + 1, 1));
                if (quotesChars.Contains(thisChar)) quotesOpen = !quotesOpen;

                if (thisChar != ' ')
                    sb.Append(thisChar);
                else if (quotesOpen || (thisChar == ' ' && nextChar != ' '))
                    sb.Append(thisChar);
            }
            sb.Append(text.Substring(text.Length - 1, 1));
            return sb.ToString();
        }

        /*
        /// <summary>
        /// This is normally the one to use
        /// General sql query to Datatable method that automatically sorts out the datatable definition to match the SQL
        /// source table but adjusts float and Int16/byte types to be double and Int32 so one can easily a maintain standard
        /// set of DataTypes in one's program!
        /// </summary>
        /// <param name="query">Sql query</param>
        /// <param name="ConnString">Database connection string</param>
        /// <param name="dt">This version has the DataTable as an output parameter as it is always dertermined within this method</param>
        /// <param name="ErrMsg">Returned error if the query fails</param>
        /// <returns>No. of rows returned</returns>
        public static int sqlQueryToDt(string query, string ConnString, out DataTable dt, out string ErrMsg)  // Comprehensive sql query to Datatable method
        {
            ErrMsg = null;
            int Nrows = 0;
            dt = new DataTable();
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnString))
                {
                    conn.Open();

                    string moddedQuery = query.ToUpper().Replace("DISTINCT", "").Replace("SELECT", "SELECT TOP 0");
                    Nrows = sqlQueryToDt(ref dt, moddedQuery, conn, out ErrMsg); // modify the SELECT clause to return no rows
                    if (!String.IsNullOrEmpty(ErrMsg)) return 0;

                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (dt.Columns[j].DataType == typeof(float)) dt.Columns[j].DataType = typeof(double);
                        if (dt.Columns[j].DataType == typeof(Int16)) dt.Columns[j].DataType = typeof(int);
                        if (dt.Columns[j].DataType == typeof(byte)) dt.Columns[j].DataType = typeof(int);
                    }

                    Nrows = sqlQueryToDt(ref dt, query, conn, out ErrMsg);
                    if (!String.IsNullOrEmpty(ErrMsg)) return 0;
                }
            }
            catch (Exception ex)
            {
                Nrows = 0;
                ErrMsg = ex.Message;
            }
            return Nrows;
        } */

        /// <summary>
        /// This is normally the one to use
        /// General sql query to Datatable method that automatically sorts out the datatable definition to match the SQL
        /// source table but adjusts float and Int16/byte types to be double and Int32 so one can easily a maintain standard
        /// set of DataTypes in one's program!
        /// Updated version that always replicates the case of the column names
        /// </summary>
        /// <param name="query">Sql query</param>
        /// <param name="sqlConn">Database connection</param>
        /// <param name="dt">This version has the DataTable as an output parameter as it is always dertermined within this method</param>
        /// <param name="ErrMsg">Returned error if the query fails</param>
        /// <returns>No. of rows returned</returns>
        public static int sqlQueryToDt(string query, SqlConnection sqlConn, out DataTable dt, out string ErrMsg, int CommandTimeout = -1)  // Comprehensive sql query to Datatable method
        {
            ErrMsg = null;
            int Nrows = 0;
            dt = new DataTable();
            try
            {
                // string moddedQuery = query.ToUpper().Replace("DISTINCT", "").Replace("SELECT", "SELECT TOP 0");
                var regexDistinct = new System.Text.RegularExpressions.Regex("DISTINCT", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                var regexSelect = new System.Text.RegularExpressions.Regex("SELECT", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                string moddedQuery = regexDistinct.Replace(query, "");   // change "DISTINCT" to ""  irrespctive of case
                moddedQuery = regexSelect.Replace(moddedQuery, "SELECT TOP 0");   // change "SELECT" to "SELECT TOP 0" irrespctive of case

                Nrows = sqlQueryToDt(ref dt, moddedQuery, sqlConn, out ErrMsg); // modify the SELECT clause to return no rows
                if (!String.IsNullOrEmpty(ErrMsg)) return 0;

                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Columns[j].DataType == typeof(float)) dt.Columns[j].DataType = typeof(double);
                    if (dt.Columns[j].DataType == typeof(Int16)) dt.Columns[j].DataType = typeof(int);
                    if (dt.Columns[j].DataType == typeof(byte)) dt.Columns[j].DataType = typeof(int);
                }

                Nrows = sqlQueryToDt(ref dt, query, sqlConn, out ErrMsg, CommandTimeout);
                if (!String.IsNullOrEmpty(ErrMsg)) return 0;
            }
            catch (Exception ex)
            {
                Nrows = 0;
                ErrMsg = ex.Message;
            }
            return Nrows;
        }

        /// <summary>
        /// Handy method to prepare an empty datatable that matches a sql table, automatically converting the column types
        /// to my usual c# norms. After using this go on to use the resulting table in a data query with SqlQueryToDt(...).
        /// Obsolete now
        /// </summary>
        /// <param name="SqlTablename"></param>
        /// <param name="SqlConn"></param>
        /// <param name="ErrMsg">Length zero means no error</param>
        /// <returns>Empty datatable matching the sql table</returns>
        public static DataTable sqlQueryGetDataTableMatchingSqlTable(string SqlTablename, string ColumnListOrStar, SqlConnection SqlConn, out string ErrMsg)
        {
            string query = "SELECT TOP 0 " + ColumnListOrStar + " FROM " + SqlTablename;

            DataTable dt = new DataTable();
            int nRows = sqlQueryToDt(ref dt, query, SqlConn, out ErrMsg);

            // adjust the dt column datatypes from SQL to C# norms
            foreach (DataColumn col in dt.Columns)
            {
                if (col.DataType == typeof(float)) col.DataType = typeof(double);
                if (col.DataType == typeof(Int16)) col.DataType = typeof(int);
            }
            return dt;
        }

        static int sqlSelectToDataTableObsolete(string SelectString, string ConnectionString, out DataTable dtOutput, out string ErrMsg)
        {
            int nrows = -1;
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    conn.Open();
                    nrows = sqlSelectToDataTableObsolete(SelectString, conn, out dtOutput, out ErrMsg);
                }
            }
            catch (Exception ex)
            {
                dtOutput = null;
                ErrMsg = ex.Message;
                return -1;
            }
            ErrMsg = null;
            return nrows;
        }

        /// <summary>
        /// Return true if the (say) appropriate meta table has a matching entry.
        /// </summary>
        /// <param name="SelectString">e.g SELECT * FROM table WHERE ... ORDER ...</param>
        /// <param name="sqlConn"></param>
        /// <param name="dtOutput">The resulting datatable which will contain one row if matching data exists - the matching row</param>
        /// <param name="ErrMsg"></param>
        /// <returns>The no. of rows found, -1 if error encountered</returns>
        static int sqlSelectToDataTableObsolete(string SelectString, SqlConnection sqlConn, out DataTable dtOutput, out string ErrMsg)
        {
            dtOutput = new DataTable();

            //  DataTable dt = new DataTable();
            int NRows = sqlQueryToDt(ref dtOutput, SelectString, sqlConn, out ErrMsg);
            if (!String.IsNullOrEmpty(ErrMsg)) return -1;
            return NRows;
        }

        #endregion Query to DataTable

        #region Query To List

        public static int sqlQueryToList(string sqlQuery, string ConnString, out IList<dynamic> Vector, out string ErrMsg)
        {
            using (SqlConnection conn = new System.Data.SqlClient.SqlConnection(ConnString))
            {
                conn.Open();
                return sqlQueryToList(sqlQuery, conn, out Vector, out ErrMsg);
            }
        }

        /// <summary>
        /// UNTESTED dynamic
        /// </summary>
        /// <param name="sqlQuery"></param>
        /// <param name="sqlConn"></param>
        /// <param name="Vector"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static int sqlQueryToList(string sqlQuery, SqlConnection sqlConn, out IList<dynamic> Vector, out string ErrMsg)
        {
            ErrMsg = null;
            Vector = null;
            DataTable dt = new DataTable();

            int nrows = sqlQueryToDt(ref dt, sqlQuery, sqlConn, out ErrMsg);

            if (dt.Columns.Count > 0)
            {
                Type type = dt.Columns[0].DataType;
                Vector = new List<dynamic>(dt.Rows.Count);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    object val = dt.Rows[i][0];
                    if (type == typeof(int))
                        Vector.Add((int)val);
                    else if (type == typeof(double))
                        Vector.Add((double)val);
                    else if (type == typeof(bool))
                        Vector.Add((bool)val);
                    else if (type == typeof(string) && val.ToString().Length > 0)
                        Vector.Add((string)val);

                }
            }
            return nrows;
        }

        /// <summary>
        /// Untested T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sqlQuery"></param>
        /// <param name="sqlConn"></param>
        /// <param name="Vector"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static int sqlQueryToListOBSOLETE<T>(string sqlQuery, SqlConnection sqlConn, out IList<T> Vector, out string ErrMsg) // General sql query to Datatable method
        {
            ErrMsg = null;
            Vector = null;
            DataTable dt = new DataTable();

            int nrows = sqlQueryToDt(ref dt, sqlQuery, sqlConn, out ErrMsg);

            if (dt.Columns.Count > 0)
            {
                Type type = dt.Columns[0].DataType;
                Vector = new List<T>(dt.Rows.Count);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    object val = dt.Rows[i][0];
                    Vector.Add((T)val);
                }
            }
            return nrows;
        }

        /// <summary>
        /// Better than dynamic option.
        /// Usage: List[string] list=sqlQueryToList[string](query,...)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sqlQuery"></param>
        /// <param name="ConnString"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static List<T> sqlQueryToList<T>(string sqlQuery, string ConnString, out string ErrMsg) // General sql query to Datatable method
        {
            using (SqlConnection conn = new System.Data.SqlClient.SqlConnection(ConnString))
            {
                conn.Open();
                return sqlQueryToList<T>(sqlQuery, conn, out ErrMsg);
            }
        }

        /// <summary>
        /// Better than dynamic option.
        /// Usage: List[string] list=sqlQueryToList[string](query,...)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sqlQuery"></param>
        /// <param name="sqlConn"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static List<T> sqlQueryToList<T>(string sqlQuery, SqlConnection sqlConn, out string ErrMsg) // General sql query to Datatable method
        {
            ErrMsg = null;
            List<T> Vector = null;
            DataTable dt = new DataTable();

            int nrows = sqlQueryToDt(ref dt, sqlQuery, sqlConn, out ErrMsg);

            if (dt.Columns.Count > 0)
            {
                Type type = dt.Columns[0].DataType;
                Vector = new List<T>(dt.Rows.Count);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    object val = dt.Rows[i][0];
                    if (val != DBNull.Value)
                        Vector.Add((T)val);
                    //else if (Nullable.GetUnderlyingType(type) != null)
                    //    Vector.Add((T)Convert.ChangeType(null, type));
                    else if (type == typeof(string))
                        Vector.Add((T)Convert.ChangeType(null, type));
                    else if (type == typeof(Double))
                        Vector.Add((T)Convert.ChangeType(null, type));

                    else
                        throw new Exception("Cannot unbox DBNull.Value into a " + val.GetType().ToString() + " in sqlQueryToList<T>");

                }
            }
            return Vector;
        }


        /// <summary>
        /// Gets a list of column names using SELECT TOP 0
        /// </summary>
        /// <param name="SqlTablename"></param>
        /// <param name="SqlConn"></param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static List<string> sqlQueryGetColumnNames(string SqlTablename, SqlConnection SqlConn, out string ErrMsg)
        {
            string query = "SELECT TOP 0 * FROM " + SqlTablename;

            DataTable dt = new DataTable();
            int nRows = sqlQueryToDt(ref dt, query, SqlConn, out ErrMsg);

            List<string> ret = new List<string>(nRows);
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ret.Add(dt.Columns[j].ColumnName);
            }
            return ret;
        }

        /// <summary>
        /// Get Sql Table Column Names using SELECT TOP 0 * FROM " + TableName;
        /// </summary>
        /// <param name="TableName"></param>
        /// <param name="conn"></param>
        /// <returns></returns>
        public static List<string> sqlGetTableColumnNames(string TableName, SqlConnection conn)
        {
            List<string> list = new List<string>();

            //SqlDataReader myReader;
            DataTable schemaTable;
            string cmd = "SELECT TOP 0 * FROM " + TableName;
            SqlCommand SqCmd = new SqlCommand(cmd, conn);

            using (SqlDataReader myReader = SqCmd.ExecuteReader(CommandBehavior.KeyInfo))
            {
                //Retrieve column schema into a DataTable.
                schemaTable = myReader.GetSchemaTable();

                foreach (DataRow row in schemaTable.Rows)
                {
                    list.Add(row.Field<string>("ColumnName"));
                }
            }
            return list;
        }

        #endregion Query To List
    }
}
