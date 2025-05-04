using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows;

using Access = Microsoft.Office.Interop.Access;

namespace AccessTableExport
{
    internal class AccessControl
    {
        private const string Provider = "Microsoft.ACE.OLEDB.12.0";

        private const string PARAM = "@criterion";

        private string connectionString;

        public string ConnectionString { get { return connectionString; } }

        public AccessControl(string sheetInfoPath, string passWord = "")
        {
            OleDbConnectionStringBuilder builder = new System.Data.OleDb.OleDbConnectionStringBuilder();
            builder["Provider"] = Provider;
            builder["Data Source"] = sheetInfoPath;

            if (!string.IsNullOrEmpty(passWord))
            {
                builder["Jet OLEDB:Database Password"] = passWord;
            }
            this.connectionString = builder.ConnectionString;
        }

        private void CREATEDataTable(OleDbConnection oDConnection, OleDbTransaction oDTransaction, DataTable dataTable)
        {
            // テーブル名を取得
            string tableName = dataTable.TableName;

            // 各列の定義を生成
            var columnDefinitions = new List<string>();
            foreach (DataColumn column in dataTable.Columns)
            {
                string columnName = column.ColumnName;
                string columnType = GetOleDbType(column.DataType); // データ型を変換
                string nullable = column.AllowDBNull ? "NULL" : "NOT NULL";
                string unique = column.Unique ? "UNIQUE" : ""; // UNIQUE制約を確認

                // デフォルト値を確認
                string defaultValue = column.DefaultValue != DBNull.Value && column.DefaultValue != null
                    ? $"DEFAULT {FormatDefaultValue(column.DefaultValue, column.DataType)}"
                    : "";

                columnDefinitions.Add($"[{columnName}] {columnType} {nullable} {unique}  {defaultValue}");
            }

            // PRIMARY KEYの定義を生成
            if (dataTable.PrimaryKey.Length > 0)
            {
                string primaryKeyColumns = string.Join(", ", dataTable.PrimaryKey.Select(c => $"[{c.ColumnName}]"));
                columnDefinitions.Add($"PRIMARY KEY ({primaryKeyColumns})");
            }

            //既存のAccessテーブルエクスポート機能でも外部キーはコピーされないので、コメントアウト
            //// 外部キー制約の定義を生成
            //foreach (DataRelation relation in dataSet.Relations)
            //{
            //    if (relation.ChildTable == dataTable)
            //    {
            //        string childColumns = string.Join(", ", relation.ChildColumns.Select(c => $"[{c.ColumnName}]"));
            //        string parentTable = relation.ParentTable.TableName;
            //        string parentColumns = string.Join(", ", relation.ParentColumns.Select(c => $"[{c.ColumnName}]"));

            //        columnDefinitions.Add($"FOREIGN KEY ({childColumns}) REFERENCES [{parentTable}] ({parentColumns})");
            //    }
            //}

            // CREATE TABLE文を生成
            string createTableQuery = $"CREATE TABLE [{tableName}] ({string.Join(", ", columnDefinitions)})";

            // クエリを実行
            using (OleDbCommand command = new OleDbCommand(createTableQuery, oDConnection, oDTransaction))
            {
                command.ExecuteNonQuery();
            }
        }

        // DataColumnの型をOleDbの型に変換するヘルパーメソッド
        private string GetOleDbType(Type type)
        {
            if (type == typeof(string)) return "TEXT";
            if (type == typeof(int)) return "INTEGER";
            if (type == typeof(long)) return "BIGINT";
            if (type == typeof(bool)) return "YESNO";
            if (type == typeof(DateTime)) return "DATETIME";
            if (type == typeof(double)) return "DOUBLE";
            if (type == typeof(decimal)) return "CURRENCY";
            if (type == typeof(byte[])) return "BINARY";
            return "TEXT"; // デフォルトはTEXT型
        }

        // デフォルト値を適切な形式にフォーマットするヘルパーメソッド
        private string FormatDefaultValue(object defaultValue, Type dataType)
        {
            if (dataType == typeof(string) || dataType == typeof(DateTime))
            {
                return $"'{defaultValue}'"; // 文字列や日付はシングルクォートで囲む
            }
            if (dataType == typeof(bool))
            {
                return (bool)defaultValue ? "TRUE" : "FALSE"; // AccessではTRUE/FALSEを使用
            }
            return defaultValue.ToString(); // その他の型はそのまま文字列化
        }

        private void INSERTDataTable(OleDbConnection oDConnection, OleDbTransaction oDTransaction, DataTable dtInsert)
        {
            // カラム名を自動生成
            IEnumerable<DataColumn> dcInsert = dtInsert.Columns.Cast<DataColumn>();

            string columns = string.Join(", ", dcInsert.Select(c => $"[{c.ColumnName}]"));
            string parameters = string.Join(", ", dcInsert.Select(c => "@" + c.ColumnName));

            //INSERT文を生成
            string query = $"INSERT INTO [{dtInsert.TableName}] ({columns}) VALUES ({parameters})";

            foreach (DataRow drInsert in dtInsert.Rows)
            {
                using (OleDbCommand insertCommand = new OleDbCommand(query, oDConnection))
                {
                    insertCommand.Transaction = oDTransaction;
                    foreach (DataColumn column in dtInsert.Columns)
                    {
                        insertCommand.Parameters.AddWithValue("@" + column.ColumnName, drInsert[column]);
                    }
                    insertCommand.ExecuteNonQuery();
                }
            }
        }

        private void SetAllowDBNullFromSchema(DataTable dataTable, OleDbConnection connection)
        {
            // スキーマ情報を取得  
            DataTable schemaTable = connection.GetSchema("Columns", new string[] { null, null, dataTable.TableName, null });
            foreach (DataRow schemaRow in schemaTable.Rows)
            {
                string columnName = schemaRow["COLUMN_NAME"].ToString();
                //bool isNullable = schemaRow["IS_NULLABLE"].ToString() == "YES";
                bool isNullable = Convert.ToBoolean(schemaRow["IS_NULLABLE"]);

                // DataTableの該当カラムにAllowDBNullを設定  
                if (dataTable.Columns.Contains(columnName))
                {
                    dataTable.Columns[columnName].AllowDBNull = isNullable;
                }
            }
        }

        private void SetPrimaryKeyFromSchema(DataTable dataTable, OleDbConnection connection)
        {
            // スキーマ情報を取得  
            DataTable schemaTable = connection.GetSchema("Indexes", new string[] { null, null, dataTable.TableName, null });
            List<DataColumn> primaryKeyColumns = new List<DataColumn>();

            foreach (DataRow schemaRow in schemaTable.Rows)
            {
                // 主キー情報を確認  
                //bool isPrimaryKey = schemaRow["PRIMARY_KEY"].ToString() == "True";
                bool isPrimaryKey = Convert.ToBoolean(schemaRow["PRIMARY_KEY"]);
                if (isPrimaryKey)
                {
                    string columnName = schemaRow["COLUMN_NAME"].ToString();
                    if (dataTable.Columns.Contains(columnName))
                    {
                        primaryKeyColumns.Add(dataTable.Columns[columnName]);
                    }
                }
            }

            // 主キーを設定  
            if (primaryKeyColumns.Count > 0)
            {
                dataTable.PrimaryKey = primaryKeyColumns.ToArray();
            }
        }

        private void SetDefaultValuesFromSchema(DataTable dataTable, OleDbConnection connection)
        {
            // スキーマ情報を取得
            DataTable schemaTable = connection.GetSchema("Columns", new string[] { null, null, dataTable.TableName, null });

            foreach (DataRow schemaRow in schemaTable.Rows)
            {
                string columnName = schemaRow["COLUMN_NAME"].ToString();

                // デフォルト値を取得
                object defaultValue = schemaRow["COLUMN_DEFAULT"];

                if (dataTable.Columns.Contains(columnName) && defaultValue != DBNull.Value && defaultValue != null)
                {
                    // DataTableの該当カラムにデフォルト値を設定
                    dataTable.Columns[columnName].DefaultValue = defaultValue;
                }
            }
        }

        private DataTable DtLoadReader(string query, string tableName, string criterion = "")
        {
            DataTable dtSelect = new DataTable(tableName);
            using (var conn = new System.Data.OleDb.OleDbConnection(this.connectionString))
            {
                conn.Open();

                using (OleDbCommand command = new OleDbCommand(query, conn))
                {
                    if (criterion != "")
                    {
                        command.Parameters.AddWithValue(PARAM, criterion);
                    }

                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        dtSelect.Load(reader);
                    }
                }

                // スキーマ情報からAllowDBNullを設定  
                SetAllowDBNullFromSchema(dtSelect, conn);

                // スキーマ情報から主キーを設定  
                SetPrimaryKeyFromSchema(dtSelect, conn);

                // スキーマ情報からデフォルト値を設定
                SetDefaultValuesFromSchema(dtSelect, conn);

                conn.Close();
            }

            return dtSelect;
        }


        public DataSet GetDataSet(List<string> copyTableList)
        {
            //コピーデータセット
            DataSet dsInsert = new DataSet();

            // コピー元のデータベースからデータをデータセットとして取得

            //コピー元からテーブルを取得する処理
            foreach (string copyTable in copyTableList)
            {
                string selectQuery = $"SELECT * FROM [{copyTable}]"; // テーブル名を指定

                DataTable dataTable = new DataTable();

                dataTable = this.DtLoadReader(selectQuery, copyTable);

                dataTable.TableName = copyTable;

                dsInsert.Tables.Add(dataTable);

                //using (OleDbCommand selectCommand = new OleDbCommand(selectQuery, sourceConnection))
                //{
                //    using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommand))
                //    {
                //        DataTable dataTable = new DataTable(copyTable);
                //        adapter.Fill(dataTable);
                //        dsInsert.Tables.Add(dataTable);
                //    }
                //}
            }

            return dsInsert;

        }

        private void DELETETable(string tableName, OleDbConnection oDConnection, ref OleDbTransaction oDTransaction)
        {
            //INSERT先のテーブルを全行削除
            using (OleDbCommand deleteCommand = new OleDbCommand($"DELETE FROM [{tableName}];", oDConnection))
            {
                deleteCommand.Transaction = oDTransaction;
                deleteCommand.ExecuteNonQuery();
            }
        }

        private void DROPTable(string tableName, OleDbConnection oDConnection, ref OleDbTransaction oDTransaction)
        {
            //INSERT先のテーブルを削除
            using (OleDbCommand deleteCommand = new OleDbCommand($"DROP TABLE [{tableName}];", oDConnection))
            {
                deleteCommand.Transaction = oDTransaction;
                deleteCommand.ExecuteNonQuery();
            }
        }

        public void ExportTable(DataTable dtInsert, OleDbConnection destinationConnection, OleDbTransaction TRN)
        {
            //INSERT先のテーブルを全行削除
            //this.DELETETable(dtInsert.TableName, destinationConnection, ref TRN);

            this.DROPTable(dtInsert.TableName, destinationConnection, ref TRN);

            this.CREATEDataTable(destinationConnection, TRN, dtInsert);

            this.INSERTDataTable(destinationConnection, TRN, dtInsert);
        }

        public object[] GetTableNames()
        {
            DataTable dt;

            using (var conn = new System.Data.OleDb.OleDbConnection(this.connectionString))
            {
                conn.Open();
                dt = conn.GetSchema("Tables");
                conn.Close();
            }

            DataTable TypeTable = dt.Select("[TABLE_TYPE] = 'TABLE'").CopyToDataTable();

            object[] TableNames = TypeTable.AsEnumerable().Select(i => i["TABLE_NAME"]).ToArray();

            string[] stTableNames = new string[TableNames.Length];

            TableNames.CopyTo(stTableNames, 0);

            return stTableNames;
        }
    }
}
