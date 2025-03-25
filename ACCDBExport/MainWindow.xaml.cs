using Microsoft.Win32;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace AccessTableExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            if (string.IsNullOrEmpty(Properties.Settings.Default.DBPassword) == false)
            {
                this.pbDBPass.Password = Properties.Settings.Default.DBPassword;
            }

            if (string.IsNullOrEmpty(Properties.Settings.Default.TableListPath) == false && System.IO.Path.Exists(Properties.Settings.Default.TableListPath)) {
                this.tblTXTPath.Text = Properties.Settings.Default.TableListPath;
            }

            Closing += FormMain_FormClosing;
        }

        /// <summary>
        /// アプリ終了時のプロパティ保存
        /// </summary>
        private void FormMain_FormClosing(object sender, CancelEventArgs e)
        {
            Properties.Settings.Default.DBFolder = this.tbxDBFilePath.Text;
            Properties.Settings.Default.DBPassword = this.pbDBPass.Password;
            Properties.Settings.Default.TableListPath = this.tblTXTPath.Text;
            Properties.Settings.Default.ToDBFolder = this.tbxToDBFilePath.Text;
            Properties.Settings.Default.Save();
        }


        private void DBSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opd = new OpenFileDialog();
            opd.Filter = "データベース(*.accdb)|*.accdb";

            if (!String.IsNullOrEmpty(Properties.Settings.Default.DBFolder))
            {
                opd.InitialDirectory = System.IO.Path.GetDirectoryName(Properties.Settings.Default.DBFolder);
                opd.FileName = System.IO.Path.GetFileName(Properties.Settings.Default.DBFolder);
            }

            if (opd.ShowDialog() == true)
            {
                Properties.Settings.Default.DBFolder = opd.FileName;
                Properties.Settings.Default.Save();
                this.tbxDBFilePath.Text = opd.FileName;
                //this.setSheetTypeCmb();

                this.setTableList();

                if (string.IsNullOrEmpty(this.tblTXTPath.Text) == false) {
                    this.selectTableList(this.tblTXTPath.Text);
                }
            }
        }

        private void setTableList() {

            AccessControl AccessTable = new AccessControl();
            object[] tableNames = AccessTable.GetTableNames(this.tbxDBFilePath.Text, this.pbDBPass.Password);

            foreach (object tableName in tableNames) {
                this.TableList.Items.Add(tableName.ToString());
            }
        }

        private void selectTableList(string tblTXTPath)
        {
            TableList.SelectedItems.Clear();

            // 読み込みたいTXTファイルのパスを指定して開く
            StreamReader sr = new StreamReader(@tblTXTPath);

            while (!sr.EndOfStream) {
                // TXTファイルの一行を読み込む
                string line = sr.ReadLine();

                foreach(string tableItem in TableList.Items) 
                {
                    if (tableItem == line)
                    {
                        TableList.SelectedItems.Add(tableItem);
                    }
                }

            }
            sr.Close();
        }

        private void TXTSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opd = new OpenFileDialog();
            opd.Filter = "テーブル一覧ファイル(*.txt)|*.txt";

            if (!String.IsNullOrEmpty(Properties.Settings.Default.TableListPath))
            {
                opd.InitialDirectory = System.IO.Path.GetDirectoryName(Properties.Settings.Default.TableListPath);
                opd.FileName = System.IO.Path.GetFileName(Properties.Settings.Default.TableListPath);
            }

            if (opd.ShowDialog() == true)
            {
                this.tblTXTPath.Text = opd.FileName;

                this.selectTableList(this.tblTXTPath.Text);
            }
        }

        private void TXTClearButton_Click(object sender, RoutedEventArgs e)
        {
            this.tblTXTPath.Text = string.Empty;
        }

        private void ToDBSelectButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opd = new OpenFileDialog();
            opd.Filter = "データベース(*.accdb)|*.accdb";

            if (!String.IsNullOrEmpty(Properties.Settings.Default.DBFolder))
            {
                opd.InitialDirectory = System.IO.Path.GetDirectoryName(Properties.Settings.Default.ToDBFolder);
                opd.FileName = System.IO.Path.GetFileName(Properties.Settings.Default.ToDBFolder);
            }

            if (opd.ShowDialog() == true)
            {
                Properties.Settings.Default.ToDBFolder = opd.FileName;
                Properties.Settings.Default.Save();
                this.tbxToDBFilePath.Text = opd.FileName;
            }
        }

        private void CopyExistButton_Click(object sender, RoutedEventArgs e)
        {
            //コピー対象のテーブルをまとめる
            List<string> CopyTableList = new List<string>();

            foreach (string oneItemLine in this.TableList.SelectedItems)
            {
                CopyTableList.Add(oneItemLine);
            }

            //string[] CopyTableLists = this.TableList.SelectedItems;

            AccessControl AccessTable = new AccessControl();

            AccessTable.ExportTable(this.tbxDBFilePath.Text, this.pbDBPass.Password, CopyTableList, this.tbxToDBFilePath.Text);

           
        }
    }
}