using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
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

        private string ConnectionString;

        public AccessControl(string sheetInfoPath, string passWord = "")
        {
            OleDbConnectionStringBuilder builder = new System.Data.OleDb.OleDbConnectionStringBuilder();
            builder["Provider"] = Provider;
            builder["Data Source"] = sheetInfoPath;

            if (!string.IsNullOrEmpty(passWord))
            {
                builder["Jet OLEDB:Database Password"] = passWord;
            }
            this.ConnectionString = builder.ConnectionString;
        }

        private void INSERTDataTable(string query, OleDbConnection oDConnection, OleDbTransaction oDTransaction, DataTable dtInsert)
        {
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
        private DataTable DtLoadReader(string query, string criterion = "")
        {
            DataTable dtSelect = new DataTable();
            using (var conn = new System.Data.OleDb.OleDbConnection(this.ConnectionString))
            {
                //ココにデータベースにアクセスするコードを書く
                // Accessのデータベースファイルに接続する
                conn.Open();

                // OleDbCommandインスタンスを生成する
                using (OleDbCommand command = new OleDbCommand(query, conn))
                {
                    // パラメータを追加する
                    if (criterion != "")
                    {
                        command.Parameters.AddWithValue(PARAM, criterion);
                    }

                    using (OleDbDataReader reader = command.ExecuteReader()) {
                        dtSelect.Load(reader);
                    }
                }
                // Accessのデータベースファイルの接続を閉じる
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

                dataTable = this.DtLoadReader(selectQuery);

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

        public List<string> ExportTable(DataSet dsInsert)
        {
            //コピー先にInsertする処理

            //結果メッセージ
            List<string> stResult = new List<string>();

            // コピー先のデータベースにデータを挿入
            using (OleDbConnection destinationConnection = new OleDbConnection(this.ConnectionString))
            {
                //トランザクション開始
                destinationConnection.Open();
                OleDbTransaction TRN = destinationConnection.BeginTransaction();

                try
                {
                    //sqlで更新する方法
                    foreach (DataTable dtInsert in dsInsert.Tables)
                    {
                        //INSERT先のテーブルを全行削除
                        this.DELETETable(dtInsert.TableName, destinationConnection, ref TRN);

                        // カラム名を自動生成
                        IEnumerable<DataColumn> dcInsert = dtInsert.Columns.Cast<DataColumn>();

                        string columns = string.Join(", ", dcInsert.Select(c => $"[{c.ColumnName}]"));
                        string parameters = string.Join(", ", dcInsert.Select(c => "@" + c.ColumnName));

                        //INSERT文を生成
                        string insertQueryTemplate = $"INSERT INTO [{dtInsert.TableName}] ({columns}) VALUES ({parameters})";

                        this.INSERTDataTable(insertQueryTemplate, destinationConnection, TRN, dtInsert);

                    }

                    //トランザクションをコミット
                    TRN.Commit();

                    ////データベースの最適化実行
                    //string tempDbPath = destinationDbPath + "_temp";
                    //string tempConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={tempDbPath};Jet OLEDB:Database Password={dbPass};";

                    //OleDbConnection tempConnection = new OleDbConnection(tempConnectionString);

                    ////// JRO.JetEngineを使用してデータベースをコンパクトおよび修復
                    ////JRO.JetEngine jetEngine = new JRO.JetEngine();
                    ////jetEngine.CompactDatabase(destinationConnectionString, tempConnectionString);

                    //CompactDatabaseHelper.JetCompact(IntPtr.Zero, destinationConnectionString, tempConnectionString, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

                    //// 元のデータベースを削除し、テンポラリデータベースをリネーム
                    //System.IO.File.Delete(destinationDbPath);
                    //System.IO.File.Move(tempDbPath, destinationDbPath);

                }
                catch (Exception ex)
                {
                    stResult.Add(ex.Message + Environment.NewLine + ex.ToString());
                    stResult.Add("コピー中にエラーが発生しました。ロールバックします。");

                    //MessageBox.Show(ex.Message + Environment.NewLine + ex.ToString(), "エラー");
                    //MessageBox.Show("ロールバック", "コピー中にエラーが発生しました。ロールバックします。");

                    try
                    {
                        TRN.Rollback();
                        stResult.Add("ロールバック成功");

                        //MessageBox.Show("ロールバック", "ロールバック成功");
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        //MessageBox.Show(ex2.Message + Environment.NewLine + ex2.ToString(), "エラー");
                        //MessageBox.Show("失敗", "ロールバック中にエラーが発生しました。");

                        stResult.Add(ex2.Message + Environment.NewLine + ex2.ToString());
                        stResult.Add("ロールバック中にエラーが発生しました。");


                    }

                }
                finally
                {
                    destinationConnection.Close();
                }
            }
            return stResult;
        }

        public object[] GetTableNames() {
            DataTable dt;

            using (var conn = new System.Data.OleDb.OleDbConnection(this.ConnectionString))
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
