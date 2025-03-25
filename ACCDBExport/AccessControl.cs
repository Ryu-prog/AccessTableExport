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

            public void ExportTable(string sourceDbPath, string dbPass, List<string> copyTableList, string destinationDbPath, string toDBPass = "")
        {

            string universalConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Jet OLEDB:Database Password={1};";

            //コピー元からテーブルを取得する処理と
            //コピー先にInsertする処理を分ける

            //コピーデータセット
            DataSet dsInsert = new DataSet();

            // コピー元のデータベースからデータをデータセットとして取得
            using (OleDbConnection sourceConnection = new OleDbConnection(String.Format(universalConnectionString, sourceDbPath, dbPass)))
            {
                sourceConnection.Open();

                //コピー元からテーブルを取得する処理
                foreach (string copyTable in copyTableList) {
                    string selectQuery = $"SELECT * FROM {copyTable}"; // テーブル名を指定

                    using (OleDbCommand selectCommand = new OleDbCommand(selectQuery, sourceConnection)) {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommand))
                        {
                            DataTable dataTable = new DataTable(copyTable);
                            adapter.Fill(dataTable);
                            dsInsert.Tables.Add(dataTable);
                        }
                    }
                }

                sourceConnection.Close();

            }

            //コピー先にInsertする処理

            //sql成否フラグ
            bool isSuccess = false;

            string destinationConnectionString = String.Format(universalConnectionString, destinationDbPath, dbPass);

            // コピー先のデータベースにデータを挿入
            using (OleDbConnection destinationConnection = new OleDbConnection(destinationConnectionString))
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
                        using (OleDbCommand deleteCommand = new OleDbCommand($"DELETE FROM [{dtInsert.TableName}];", destinationConnection))
                        {
                            deleteCommand.Transaction = TRN;
                            deleteCommand.ExecuteNonQuery();
                        }

                        // カラム名を自動生成
                        IEnumerable<DataColumn> dcInsert = dtInsert.Columns.Cast<DataColumn>();

                        string columns = string.Join(", ", dcInsert.Select(c => $"[{c.ColumnName}]"));
                        string parameters = string.Join(", ", dcInsert.Select(c => "@" + c.ColumnName));

                        //INSERT文を生成
                        string insertQueryTemplate = $"INSERT INTO {dtInsert.TableName} ({columns}) VALUES ({parameters})";

                        foreach (DataRow drInsert in dtInsert.Rows)
                        {
                            using (OleDbCommand insertCommand = new OleDbCommand(insertQueryTemplate, destinationConnection))
                            {
                                insertCommand.Transaction = TRN;
                                foreach (DataColumn column in dtInsert.Columns)
                                {
                                    insertCommand.Parameters.AddWithValue("@" + column.ColumnName, drInsert[column]);
                                }
                                insertCommand.ExecuteNonQuery();
                            }
                        }
                    }

                    //トランザクションをコミット
                    TRN.Commit();

                    isSuccess = true;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + Environment.NewLine + ex.ToString(), "エラー");
                    MessageBox.Show("ロールバック", "コピー中にエラーが発生しました。ロールバックします。");

                    try
                    {
                        TRN.Rollback();
                        MessageBox.Show("ロールバック", "ロールバック成功");
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred
                        // on the server that would cause the rollback to fail, such as
                        // a closed connection.
                        MessageBox.Show(ex2.Message + Environment.NewLine + ex2.ToString(), "エラー");
                        MessageBox.Show("失敗", "ロールバック中にエラーが発生しました。");
                    }

                }
                finally
                {
                    destinationConnection.Close();
                }
            }

            if (isSuccess) MessageBox.Show("正常終了", "正常終了");

        }

        public object[] GetTableNames(string SheetInfoPath, string DBPass) {
            var builder = new System.Data.OleDb.OleDbConnectionStringBuilder();
            builder["Provider"] = Provider;
            builder["Data Source"] = SheetInfoPath;
            builder["Jet OLEDB:Database Password"] = DBPass;

            DataTable dt;

            using (var conn = new System.Data.OleDb.OleDbConnection(builder.ConnectionString))
            {
                conn.Open();
                dt = conn.GetSchema("Tables");
                conn.Close();
            }

            DataTable TypeTable = dt.Select("[TABLE_TYPE] = 'TABLE'").CopyToDataTable();

            object[] TableNames = TypeTable.AsEnumerable().Select(i => i["TABLE_NAME"]).ToArray();

            return TableNames;
        }
    }
}
