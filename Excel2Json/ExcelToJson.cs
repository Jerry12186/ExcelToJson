using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace Excel2Json
{
    class ExcelToJson
    {
        /// <summary>
        /// excel文件的路径
        /// </summary>
        private string excelPath;

        /// <summary>
        /// Excel连接
        /// </summary>
        OleDbConnection connection = null;

        public ExcelToJson(string path)
        {
            if (path == "")
            {
                Console.WriteLine("Excel文件路径错误");
                return;
            }
            excelPath = path;
            connection = InitConntion();
            connection.Open();
        }

        private OleDbConnection InitConntion()
        {
            if (!File.Exists(excelPath))
            {
                Console.WriteLine("指定文件不存在--->" + excelPath);
                return null;
            }

            string strExtension = Path.GetExtension(excelPath);
            string initStr = string.Empty;

            switch (strExtension)
            {
                case ".xls":
                    initStr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1;\"", excelPath);
                    break;
                case ".xlsx":
                    initStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1;\"", excelPath);
                    break;
                default:
                    Console.WriteLine("指定文件不是Excel文件。");
                    break;
            }
            return new OleDbConnection(initStr);
        }

        /// <summary>
        /// 获取Excel表单名字
        /// </summary>
        /// <returns></returns>
        private List<string> GetExecelSheetNames()
        {
            List<string> sheetNames = new List<string>();
            DataTable dataTable = null;
            dataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                sheetNames.Add(dataTable.Rows[i]["Table_Name"].ToString().Split('$')[0]);
            }

            return sheetNames;
        }

        /// <summary>
        /// 根据sheet名字获取Excel内容
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private DataTable GetExcelContent(string sheetName)
        {
            if (sheetName == "_xlnm#_FilterDatabase")
                return null;
            DataSet ds = new DataSet();
            string commandStr = string.Format("SELECT * FROM [{0}$]", sheetName);
            OleDbCommand command = new OleDbCommand(commandStr, connection);
            OleDbDataAdapter data = new OleDbDataAdapter(commandStr, connection);
            data.Fill(ds, sheetName);

            DataTable table = ds.Tables[sheetName];
            for (int i = 0; i < table.Rows[0].ItemArray.Length; i++)
            {
                var cloumnName = table.Rows[0].ItemArray[i].ToString();
                if (!string.IsNullOrEmpty(cloumnName))
                    table.Columns[i].ColumnName = cloumnName;
            }
            table.Rows.RemoveAt(0);

            return table;
        }

        public List<string> GetSheetNames()
        {
            return GetExecelSheetNames();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string ToJson()
        {
            JObject json = new JObject();
            List<string> tableNames = GetExecelSheetNames();

            tableNames.ForEach(tableName =>
            {
                var table = new JArray() as dynamic;
                DataTable dataTable = GetExcelContent(tableName);
                foreach (DataRow dataRow in dataTable.Rows)
                {
                    dynamic row = new JObject();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        row.Add(column.ColumnName, dataRow[column.ColumnName].ToString());
                    }
                    table.Add(row);
                }
                json.Add(tableName, table);
            });

            return json.ToString();
        }

        public void Close()
        {
            if (connection != null)
            {
                connection.Close();
                connection.Dispose();
            }
        }
    }
}
