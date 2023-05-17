using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Timers;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SWIFT_СПФС
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string connectionString;
        static OleDbDataAdapter adapter;
        static DataTable FilesTable;
        static OleDbConnection connection;
        static bool ProcessRun = false;
        static int Cnt = 0;
        static int Error = 0;
        static bool ErrorNow = false;
        //---------------------------------------------------init---------------------------------------------------
        public MainWindow()
        {
            InitializeComponent();
            datePickerFrom.SelectedDate = GlobalTrash.datePickerFrom;
            datePickerTo.SelectedDate = GlobalTrash.datePickerTo;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = SPFS.MDB;Mode = ReadWrite;";
        }
        //------------------------------------------------data base-------------------------------------------------
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            FilesTable = new DataTable();
            try
            {
                connection = new OleDbConnection(connectionString);
                OleDbCommand command = new OleDbCommand(null, connection);
                adapter = new OleDbDataAdapter(command);
                Refresh_FilesTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection != null)
                    connection.Close();
            }
        }

        private void Refresh_FilesTable()
        {
            if (adapter != null)
            {
                adapter.SelectCommand.CommandText = "select f.*\n" +
                    "from files f\n" +
                    "where FiledateTime between #" + GlobalTrash.datePickerFrom.ToString("MM/dd/yyyy").Replace(".", "/") +
                    "# and #" + GlobalTrash.datePickerTo.ToString("MM/dd/yyyy").Replace(".", "/") + "#\n" +
                    "order by FiledateTime desc";
                connection.Open();
                FilesTable.Clear();
                adapter.Fill(FilesTable);
                // определяем первичный ключ таблицы Files
                FilesTable.PrimaryKey = new DataColumn[] { FilesTable.Columns["ID"] };
                if (FilesGrid.ItemsSource == null) FilesGrid.ItemsSource = FilesTable.DefaultView;
                connection.Close();
            }
        }

        private void Refresh(object sender, RoutedEventArgs e)
        {
            Refresh_FilesTable();
        }

        //-------------------------------------------------settings-------------------------------------------------
        private void Button_Click_Settings(object sender, RoutedEventArgs e)
        {
            Настройки Settings = new Настройки();
            Settings.Owner = this;
            //Settings.ShowDialog();
            if (Settings.ShowDialog() == true)
            {
                GlobalTrash.CodeBnk = Settings.My_CodeBnk;
                GlobalTrash.manager.WritePrivateString("General", "CodeBnk", GlobalTrash.CodeBnk);

                GlobalTrash.SWIFTBnk = Settings.My_SWIFTBnk;
                GlobalTrash.manager.WritePrivateString("General", "SWIFTBnk", GlobalTrash.SWIFTBnk);

                GlobalTrash.NameBnk = Settings.My_NameBnk;
                GlobalTrash.manager.WritePrivateString("General", "NameBnk", GlobalTrash.NameBnk);

                GlobalTrash.Mist = Settings.My_Mist;
                GlobalTrash.manager.WritePrivateString("General", "Mist", GlobalTrash.Mist);

                GlobalTrash.PathABSOut = Settings.My_PathABSOut;
                GlobalTrash.manager.WritePrivateString("SWIFT", "PathABSOut", GlobalTrash.PathABSOut);

                GlobalTrash.PathABSIn = Settings.My_PathABSIn;
                GlobalTrash.manager.WritePrivateString("SWIFT", "PathABSIn", GlobalTrash.PathABSIn);

                GlobalTrash.FileMaskOut = Settings.My_FileMaskOut;
                GlobalTrash.manager.WritePrivateString("SWIFT", "FileMaskOut", GlobalTrash.FileMaskOut);

                GlobalTrash.PathTransit = Settings.My_PathTransit;
                GlobalTrash.manager.WritePrivateString("SWIFT", "PathTransit", GlobalTrash.PathTransit);

                GlobalTrash.PathOut = Settings.My_PathOut;
                GlobalTrash.manager.WritePrivateString("SWIFT", "PathOut", GlobalTrash.PathOut);

                GlobalTrash.PathIn = Settings.My_PathIn;
                GlobalTrash.manager.WritePrivateString("SWIFT", "PathIn", GlobalTrash.PathIn);

                GlobalTrash.FileMaskIn = Settings.My_FileMaskIn;
                GlobalTrash.manager.WritePrivateString("SWIFT", "FileMaskIn", GlobalTrash.FileMaskIn);

                GlobalTrash.PathArc = Settings.My_PathArc;
                GlobalTrash.manager.WritePrivateString("SWIFT", "PathArc", GlobalTrash.PathArc);

                GlobalTrash.SPFSPathIn = Settings.My_SPFSPathIn;
                GlobalTrash.manager.WritePrivateString("SPFS", "SPFSPathIn", GlobalTrash.SPFSPathIn);

                GlobalTrash.FileMask = Settings.My_FileMask;
                GlobalTrash.manager.WritePrivateString("SPFS", "FileMask", GlobalTrash.FileMask);

                GlobalTrash.SPFSPathOut = Settings.My_SPFSPathOut;
                GlobalTrash.manager.WritePrivateString("SPFS", "SPFSPathOut", GlobalTrash.SPFSPathOut);

                GlobalTrash.SPFSPathTransit = Settings.My_SPFSPathTransit;
                GlobalTrash.manager.WritePrivateString("SPFS", "SPFSPathTransit", GlobalTrash.SPFSPathTransit);

                GlobalTrash.SPFSPathArc = Settings.My_SPFSPathArc;
                GlobalTrash.manager.WritePrivateString("SPFS", "SPFSPathArc", GlobalTrash.SPFSPathArc);
                //MessageBox.Show("b" + GlobalTrash.manager.GetPrivateString("General", "CodeBnk"));   изменить ini
            }
        }

        private void Button_Click_Exit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_Document(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView row = (DataRowView)FilesGrid.SelectedItems[0];
                string FileName = "" + row["FileName"];
                string[] ext = FileName.Split('.');
                string State = "" + row["State"];
                FileInfo fi;
                if (State == "INIT")
                {
                    if (String.Compare(ext[1], "in") == 0 || String.Compare(ext.Last(), "out") == 0)
                        FileName = GlobalTrash.PathTransit + "\\" + FileName;
                    else FileName = GlobalTrash.SPFSPathTransit + "\\" + FileName;
                    fi = new FileInfo(FileName);
                    string text = File.ReadAllText(@fi.FullName);
                    MessageBox.Show(text);
                }
                else
                {
                    if (State == "ERROR") MessageBox.Show("Ошибка чтения файла");
                    else MessageBox.Show("Файл уже отправлен");
                }
            }
            catch
            { 
                MessageBox.Show("Не выбран документ");
            }
        }

        //---------------------------------------------------date---------------------------------------------------
        private void From_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            DateTime? selectedDate = datePickerFrom.SelectedDate;
            GlobalTrash.datePickerFrom = selectedDate.Value;
            Refresh_FilesTable();
        }
        private void To_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            DateTime? selectedDate = datePickerTo.SelectedDate;
            GlobalTrash.datePickerTo = selectedDate.Value;
            Refresh_FilesTable();
        }


        //--------------------------------------------------timer---------------------------------------------------

        private static System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
        private void initializeTimer()
        {
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 5);
            dispatcherTimer.Start();
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            if (!ProcessRun)
            {
                ProcessRun = true;
                ReadFiles();
                if (Cnt > 0 || Error > 0)
                {
                    Refresh_FilesTable();
                    Processed.Text = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + "\nУспешно обработано: " + Cnt + "\nОбнаружено ошибок: ";
                    Errors.Text = "" + Error;
                    if (Error != 0) Errors.Foreground = Brushes.Red;
                    Cnt = 0;
                    Error = 0;
                }
                ProcessRun = false;
            }
        }

        public bool IsDuplicate(string fiName)
        {
            //Проверка на дубликат
            bool duplicate = false;
            string sql = "select * from Files where FileName='" + fiName + "'";
            OleDbCommand com = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataReader reader = com.ExecuteReader();
            while (reader.Read() == true)
                duplicate = true;
            reader.Close();
            connection.Close();
            return duplicate;
        }
        private void ReadFiles()
        {
            bool duplicate;
            //Исходящие SWIFT
            DirectoryInfo diPathABSOut = new DirectoryInfo(@GlobalTrash.PathABSOut);
            foreach (var fi in diPathABSOut.GetFiles(GlobalTrash.FileMaskOut))
            {
                if (fi.Exists)
                {
                    duplicate = IsDuplicate(fi.Name);
                    if (duplicate)
                    {
                        Error++;
                        fi.CopyTo(GlobalTrash.Mist + "\\" + fi.Name, true);
                        fi.Delete();
                        continue;
                    }

                    Cnt++;
                    //Записать информацию в базу
                    InsertFile(fi, "SWIFT", "");
                    if (ErrorNow == false)
                    {
                        //Скопировать в архив
                        fi.CopyTo(GlobalTrash.PathArc + "\\" + fi.Name, true);
                        //Переместить в транзитный каталог
                        fi.CopyTo(GlobalTrash.PathTransit + "\\" + fi.Name, true);
                        fi.Delete();
                    }
                    else
                    {
                        fi.CopyTo(GlobalTrash.Mist + "\\" + fi.Name, true);
                        fi.Delete();
                        ErrorNow = false;
                    }
                }
            }
            //Входящие SWIFT
            DirectoryInfo diPathIn = new DirectoryInfo(@GlobalTrash.PathIn);
            foreach (var fi in diPathIn.GetFiles(GlobalTrash.FileMaskIn))
            {
                if (fi.Exists)
                {
                    duplicate = IsDuplicate(fi.Name);
                    if (duplicate)
                    {
                        Error++;
                        fi.CopyTo(GlobalTrash.Mist + "\\" + fi.Name, true);
                        fi.Delete();
                        continue;
                    }

                    Cnt++;
                    //Записать информацию в базу
                    InsertFile(fi, "SWIFT", GlobalTrash.SWIFTBnk);
                    if (ErrorNow == false)
                    {

                        //Скопировать в архив
                        fi.CopyTo(GlobalTrash.PathArc + "\\" + fi.Name, true);
                        //Переместить в транзитный каталог
                        fi.CopyTo(GlobalTrash.PathTransit + "\\" + fi.Name, true);
                        fi.Delete();
                    }
                    else
                    {
                        fi.CopyTo(GlobalTrash.Mist + "\\" + fi.Name, true);
                        fi.Delete();
                        ErrorNow = false;
                    }
                }
            }
            //СПФС
            DirectoryInfo diSPFSPathIn = new DirectoryInfo(@GlobalTrash.SPFSPathIn);
            foreach (var fi in diSPFSPathIn.GetFiles(GlobalTrash.FileMask))
            {
                if (fi.Exists)
                {
                    duplicate = IsDuplicate(fi.Name);
                    if (duplicate)
                    {
                        Error++;
                        fi.CopyTo(GlobalTrash.Mist + "\\" + fi.Name, true);
                        fi.Delete();
                        continue;
                    }

                    Cnt++;
                    //Записать информацию в базу
                    InsertFile(fi, "СПФС", "");
                    if (ErrorNow == false)
                    {
                        //Скопировать в архив
                        fi.CopyTo(GlobalTrash.SPFSPathArc + "\\" + fi.Name, true);
                        //Переместить в транзитный каталог
                        fi.CopyTo(GlobalTrash.SPFSPathTransit + "\\" + fi.Name, true);
                        fi.Delete();
                    }
                    else
                    {
                        fi.CopyTo(GlobalTrash.Mist + "\\" + fi.Name, true);
                        fi.Delete();
                        ErrorNow = false;
                    }
                }
            }
        }

        private void InsertFile(FileInfo fi, string pSysFrom, string pReceiverSWIFT)
        {
            string SenderSWIFT = "", MsgCode = "", ReceiverSWIFT = "", MsgRef = "", ValueDate = "", Curr = "", Amount = "", Payer = "";
            try
            {
                if (String.Compare(pReceiverSWIFT, "") == 0)
                {
                    //Парсинг файла
                    string[] readText = File.ReadAllLines(@fi.FullName);
                    string message = "";
                    foreach (string s in readText)
                        message += s + "\n";
                    //MessageBox.Show(message);
                    string[] separatingStrings = { "{1:", "{2:", ":20:", ":32A:", ":50" };
                    string[] parts = message.Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries);
                    //MessageBox.Show(parts[5]);

                    //SWIFT отправителя сообщения
                    SenderSWIFT = parts[1].Substring(3, 12);
                    if (SenderSWIFT.Substring(9) == "XXX")
                        SenderSWIFT = SenderSWIFT.Substring(0, 8);
                    else SenderSWIFT = SenderSWIFT.Substring(0, 8) + SenderSWIFT.Substring(9);
                    //Код сообщения   //SWIFT получателя сообщения
                    MsgCode = parts[2].Substring(1, 3);
                    ReceiverSWIFT = parts[2].Substring(4, 12);
                    if (ReceiverSWIFT.Substring(9) == "XXX")
                        ReceiverSWIFT = ReceiverSWIFT.Substring(0, 8);
                    else ReceiverSWIFT = ReceiverSWIFT.Substring(0, 8) + ReceiverSWIFT.Substring(9);
                    //Референс сообщения
                    string[] Refs = parts[3].Split(':');
                    MsgRef = Refs[0];
                    MsgRef = MsgRef.Substring(0, MsgRef.Length - 1);
                    //Дата валютирования   //Код валюты   //Сумма сообщения
                    if (message.Contains(":32A:"))
                    {
                        string[] vca = parts[4].Split('\n');
                        ValueDate = vca[0].Substring(0, 6);
                        Curr = vca[0].Substring(6, 3);
                        Amount = vca[0].Substring(9);
                    }
                    //Плательщик
                    Payer = "";
                    string[] Pays;
                    if (message.Contains(":32A:") && message.Contains(":50"))
                    {
                        Pays = parts[5].Split('\n');
                        foreach (string P in Pays)
                        {
                            if (P.StartsWith(":"))
                                break;
                            if (!P.StartsWith("F") && !P.StartsWith("K"))
                                Payer += P + "\n";
                        }
                        Payer = Payer.Substring(0, Payer.Length - 1);
                    }
                    if (!message.Contains(":32A:") && message.Contains(":50"))
                    {
                        Pays = parts[4].Split('\n');
                        foreach (string P in Pays)
                        {
                            if (P.StartsWith(":"))
                                break;
                            if (!P.StartsWith("F") && !P.StartsWith("K"))
                                Payer += P + "\n";
                        }
                        Payer = Payer.Substring(0, Payer.Length - 1);
                    }
                    if (!message.Contains(":50"))
                    {
                        Payer = "";
                    }
                }
                else ReceiverSWIFT = pReceiverSWIFT;

                //Сохранение файла в базу
                string sql = "insert into Files (FileName, SysFrom, FiledateTime, State, MsgCode, MsgRef, SenderSWIFT, ReceiverSWIFT, ValueDate, Curr, Amount, Payer)" +
                            "values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                OleDbCommand com = new OleDbCommand(sql, connection);
                com.Parameters.Add("@FileName", OleDbType.VarWChar, 255).Value = fi.Name;
                com.Parameters.Add("@SysFrom", OleDbType.VarWChar, 5).Value = pSysFrom;
                com.Parameters.Add("@FiledateTime", OleDbType.Date, 40).Value = fi.CreationTime;
                //Состояние
                com.Parameters.Add("@State", OleDbType.VarWChar, 10).Value = "INIT";

                //Инфоормация по платежному документу
                //Код сообщения
                com.Parameters.Add("@MsgCode", OleDbType.VarWChar, 3).Value = MsgCode;
                //Референс сообщения
                com.Parameters.Add("@MsgRef", OleDbType.VarWChar, 20).Value = MsgRef;
                //SWIFT отправителя сообщения
                com.Parameters.Add("@SenderSWIFT", OleDbType.VarWChar, 12).Value = SenderSWIFT;
                //SWIFT получателя сообщения
                com.Parameters.Add("@ReceiverSWIFT", OleDbType.VarWChar, 12).Value = ReceiverSWIFT;
                //Дата валютирования
                if (ValueDate != "")
                    com.Parameters.Add("@ValueDate", OleDbType.Date, 40).Value = DateTime.ParseExact(ValueDate, "yyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                else com.Parameters.Add("@ValueDate", OleDbType.Date, 40).Value = DBNull.Value;
                //Код валюты
                com.Parameters.Add("@Curr", OleDbType.VarWChar, 3).Value = Curr;
                //Сумма сообщения
                if (Amount != "")
                {
                    string am = Amount;
                    if (am.EndsWith(",")) am = am + "00";
                    decimal amount = decimal.Parse(am);
                    com.Parameters.Add("@Amount", OleDbType.Currency, 40).Value = amount;
                }
                else com.Parameters.Add("@Amount", OleDbType.Currency, 40).Value = DBNull.Value;
                //Плательщик
                com.Parameters.Add("@Payer", OleDbType.VarWChar, 200).Value = Payer;

                connection.Open();
                OleDbTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);
                com.Transaction = transaction;
                com.ExecuteNonQuery();
                transaction.Commit();
                connection.Close();
            }
            catch
            {
                Error++;
                Cnt--;
                ErrorNow = true;
            }
        }

        private void Start_Timer(object sender, RoutedEventArgs e)
        {
            initializeTimer();
            StartItem.IsEnabled = false;
            StartStop.Text = "Обработка файлов запущена";
            StopItem.IsEnabled = true;
            StartButton.IsEnabled = false;
            StopButton.IsEnabled = true;
        }
        private void Stop_Timer(object sender, RoutedEventArgs e)
        {
            dispatcherTimer.Stop();
            StartItem.IsEnabled = true;
            StopItem.IsEnabled = false;
            StartStop.Text = "Обработка файлов остановлена";
            StartButton.IsEnabled = true;
            StopButton.IsEnabled = false;
        }


        //---------------------------------------------------send---------------------------------------------------
        public void GoToID(int id)
        {
            int id2 = 0;
            string idS2;
            DataRowView row2;
            for (int i = 0; i<FilesGrid.Items.Count; i++)
            {
                row2 = (DataRowView) FilesGrid.Items[i];
                idS2 = "" + row2["ID"];
                id2 = int.Parse(idS2);
                if (id2 == id)
                {
                    FilesGrid.Focus();
                    FilesGrid.SelectedItem = FilesGrid.Items[i];
                    break;
                }
            }
        }
        
        private void Button_ToSWIFT(object sender, RoutedEventArgs e)
        {
            if (FilesGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)FilesGrid.SelectedItems[0];
                
                string idS = "" + row["ID"];
                int id = int.Parse(idS);
                MessageBoxResult result = MessageBox.Show("Выполнить отправку в SWIFT?", "Отправка в SWIFT", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    string Err = SendToSWIFT(id);
                    if (Err != "") MessageBox.Show("Ошибка обработки сообщения\n" + Err);
                    Refresh_FilesTable();
                    GoToID(id);  
                }
            }
        }

        private string SendToSWIFT(int id)
        {
            string Result = "";
            DirectoryInfo di = new DirectoryInfo(GlobalTrash.PathOut);
            if (!di.Exists) return "Не задан каталог для отправки файлов SWIFT";

            string sql = "select * from Files where ID=" + id.ToString();
            OleDbCommand com = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataReader reader = com.ExecuteReader();
            string FileName = "", SysFrom = "", SenderSWIFT= "", ReceiverSWIFT = "";
            while (reader.Read() == true)
            {
                if (String.Compare(reader["ID"] as string, "") == 0) { Result = "Документ не найден"; break; }
                if (String.Compare(Result, "") == 0 && String.Compare(reader["State"] as string, "INIT") != 0) { Result = "Документ уже обработан"; break; }
                if (String.Compare(Result, "") == 0 && String.Compare(reader["SysTo"] as string, "") == 0) { Result = "Документ уже обработан"; break; }
                FileName = reader["FileName"] as string;
                if (String.Compare(reader["SysFrom"] as string, "SWIFT") == 0)
                    FileName = GlobalTrash.PathTransit + "\\" + FileName;
                else
                    FileName = GlobalTrash.SPFSPathTransit + "\\" + FileName;
                SysFrom = reader["SysFrom"] as string;
                SenderSWIFT = reader["SenderSWIFT"] as string;
                ReceiverSWIFT = reader["ReceiverSWIFT"] as string;
            }
            reader.Close();
            //Ошибок не найдено
            if (String.Compare(Result, "") == 0)
            {
                FileInfo fi = new FileInfo(FileName);
                OleDbTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);
                com.Transaction = transaction;
                if (!fi.Exists)
                {
                    //Файл удалили из транзитного каталога, помечаем как ошибку
                    com.CommandText = "update files set State=?, ErrorText=? where ID=" + id.ToString();
                    com.Parameters.Clear();
                    com.Parameters.Add("@State", OleDbType.VarWChar, 10).Value = "ERROR";
                    com.Parameters.Add("@ErrorText", OleDbType.VarWChar, 255).Value = "Не найден файл" + fi.Name;
                    com.ExecuteNonQuery();
                    transaction.Commit();

                }
                else
                {
                    if (String.Compare(SysFrom, "SWIFT") == 0 && String.Compare(SenderSWIFT, GlobalTrash.SWIFTBnk) == 0)
                    {
                        //Исходящий SWIFT в SWIFT
                        try
                        {
                            //Запишем информацию в базу
                            com.CommandText = "update files set State=?, SysTo=?, FileNameOutput=? where ID=" + id.ToString();
                            com.Parameters.Clear();
                            com.Parameters.Add("@State", OleDbType.VarWChar, 10).Value = "SWIFT";
                            com.Parameters.Add("@SysTo", OleDbType.VarWChar, 5).Value = "SWIFT";
                            com.Parameters.Add("@FileNameOutput", OleDbType.VarWChar, 255).Value = fi.Name;
                            com.ExecuteNonQuery();
                            transaction.Commit();
                            //Скопируем файл
                            fi.CopyTo(GlobalTrash.PathOut + "\\" + fi.Name, true);
                            fi.Delete();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            Result = ex.Message;
                        }

                    }
                    else if ((String.Compare(SysFrom, "SWIFT") == 0 && String.Compare(ReceiverSWIFT, GlobalTrash.SWIFTBnk) == 0) ||
                            (String.Compare(SysFrom, "СПФС") == 0))
                    {
                        //Входящий SWIFT на прием в АБС по каналу SWIFT
                        //Входящий СПФС на прием в АБС по каналу SWIFT
                        try
                        {
                            //Запишем информацию в базу
                            com.CommandText = "update files set State=?, SysTo=?, FileNameOutput=? where ID=" + id.ToString();
                            com.Parameters.Clear();
                            com.Parameters.Add("@State", OleDbType.VarWChar, 10).Value = "SWIFT";
                            com.Parameters.Add("@SysTo", OleDbType.VarWChar, 5).Value = "SWIFT";
                            com.Parameters.Add("@FileNameOutput", OleDbType.VarWChar, 255).Value = fi.Name;
                            com.ExecuteNonQuery();
                            transaction.Commit();
                            //Скопируем файл
                            fi.CopyTo(GlobalTrash.PathABSIn + "\\" + fi.Name, true);
                            fi.Delete();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            Result = ex.Message;
                        }

                    }
                }
            }
            connection.Close();
            return Result;
        }

        private void Button_ToSPFSru(object sender, RoutedEventArgs e)
        {
            if (FilesGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)FilesGrid.SelectedItems[0];
                string idS = "" + row["ID"];
                int id = int.Parse(idS);
                MessageBoxResult result = MessageBox.Show("Выполнить отправку в СПФС-RU?", "Отправка в СПФС-RU?", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    string Err = SendToSPFS(id, "СПФС-RU");
                    if (Err != "") MessageBox.Show("Ошибка обработки сообщения\n" + Err);
                    Refresh_FilesTable();
                    GoToID(id);
                }
            }
        }

        private void Button_ToSPFSby(object sender, RoutedEventArgs e)
        {
            if (FilesGrid.SelectedItem != null)
            {
                DataRowView row = (DataRowView)FilesGrid.SelectedItems[0];
                string idS = "" + row["ID"];
                int id = int.Parse(idS);
                MessageBoxResult result = MessageBox.Show("Выполнить отправку в СПФС-BY?", "Отправка в СПФС-BY?", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    string Err = SendToSPFS(id, "СПФС-BY");
                    if (Err != "") MessageBox.Show("Ошибка обработки сообщения\n" + Err);
                    Refresh_FilesTable();
                    GoToID(id);
                }
            }
        }

        //Преобразование числа Value в строковое представление числа в системе счисления ss
        private string DecToStringSS(int Value, int ss)
        {
            string s = "";
            int i = 0, v = Value;
            while (v > 0) 
            {
                i = v % ss;
                v = v / ss;
                if (i > 9) s = Convert.ToChar(i - 10 + (int)'A') + s;
                else s = Convert.ToChar(i + (int)'0') + s;
            }
            return s;
        }

        private string SendToSPFS(int id, string systo)
        {
            string Result = "", s = "", CodeBnk = "";
            int N = 0;
            DirectoryInfo di = new DirectoryInfo(GlobalTrash.SPFSPathOut);
            if (!di.Exists) return "Не задан каталог для отправки файлов СПФС";

            string sql = "select * from Files where ID=" + id.ToString();
            OleDbCommand com = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataReader reader = com.ExecuteReader();
            string FileName = "", NewFileName = "", ReceiverSWIFT = "", MsgCode = "";
            DateTime D;
            while (reader.Read() == true)
            {
                if (String.Compare(reader["ID"] as string, "") == 0) { Result = "Документ не найден"; break; }
                if (String.Compare(Result, "") == 0 && String.Compare(reader["State"] as string, "INIT") != 0) { Result = "Документ уже обработан"; break; }
                if (String.Compare(Result, "") == 0 && String.Compare(reader["SysTo"] as string, "") == 0) { Result = "Документ уже обработан"; break; }
                if (String.Compare(reader["SysFrom"] as string, "СПФС") == 0) { Result = "Документ нельзя отправить в СПФС"; break; }
                if (String.Compare(reader["SysFrom"] as string, "SWIFT") == 0 && String.Compare(reader["SenderSWIFT"] as string, GlobalTrash.SWIFTBnk) != 0) { Result = "Входящий документ SWIFT нельзя отправить в СПФС"; break; }
                FileName = reader["FileName"] as string;
                FileName = GlobalTrash.PathTransit + "\\" + FileName;
                ReceiverSWIFT = reader["ReceiverSWIFT"] as string;
                MsgCode = reader["MsgCode"] as string;
            }
            reader.Close();
            //Ошибок не найдено
            if (String.Compare(Result, "") == 0)
            {
                FileInfo fi = new FileInfo(FileName);
                OleDbTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);
                com.Transaction = transaction;
                if (!fi.Exists)
                {
                    //Файл удалили из транзитного каталога, помечаем как ошибку
                    com.CommandText = "update files set State=?, ErrorText=? where ID=" + id.ToString();
                    com.Parameters.Clear();
                    com.Parameters.Add("@State", OleDbType.VarWChar, 10).Value = "ERROR";
                    com.Parameters.Add("@ErrorText", OleDbType.VarWChar, 255).Value = "Не найден файл" + fi.Name;
                    com.ExecuteNonQuery();
                    transaction.Commit();
                }
                else
                {
                    try
                    {
                        //Блок {S: заменить на пустой блок {5:}
                        string text = File.ReadAllText(@fi.FullName);
                        if (text.Contains("-}{S:"))
                        {
                            string[] sepStrings = { "-}{S:" };
                            string[] texts = text.Split(sepStrings, System.StringSplitOptions.RemoveEmptyEntries);
                            File.WriteAllText(@fi.FullName, texts[0] + "-}{5:}");
                        }

                        if (String.Compare(systo, "СПФС-RU") == 0)
                        {
                            //Читаем счётчик отправок за день из таблицы FileNum
                            D = DateTime.Today;
                            com.CommandText = "select N from FileNum where OperDate=?";
                            com.Parameters.Clear();
                            com.Parameters.Add("@OperDate", OleDbType.Date, 40).Value = D;
                            reader = com.ExecuteReader();
                            N = 0;
                            while (reader.Read() == true)
                                N = reader.GetInt32(0) + 1;
                            reader.Close();
                            //Записываем значение счётчика в базу
                            if (N == 0)
                            {
                                N = 1;
                                com.CommandText = "insert into FileNum(N, OperDate) values(?, ?)";
                            }
                            else com.CommandText = "update FileNum set N=? where OperDate=?";
                            com.Parameters.Clear();
                            com.Parameters.Add("@N", OleDbType.Integer, 10).Value = N;
                            com.Parameters.Add("@OperDate", OleDbType.Date, 40).Value = D;
                            com.ExecuteNonQuery();

                            //Генерируем новое имя файла для RU
                            s = N.ToString();
                            s = s.PadLeft(6, '0');
                            NewFileName = "ED503." +
                                          D.ToString("ddMM") +
                                          s +
                                          '.' + 
                                          GlobalTrash.CodeBnk +
                                          ".IN.txt";
                        }
                        else if (String.Compare(systo, "СПФС-BY") == 0)
                        {
                            //Читаем код UNUR gjkexfntkz из таблицы Bank
                            com.CommandText = "select UNUR from Bank where SWIFT=?";
                            com.Parameters.Clear();
                            com.Parameters.Add("@SWIFT", OleDbType.VarWChar, 12).Value = ReceiverSWIFT;
                            reader = com.ExecuteReader();
                            while (reader.Read() == true)
                                CodeBnk = reader["UNUR"] as string;
                            reader.Close();
                            //Читаем счётчик отправок за день из таблицы FileNum
                            D = new DateTime(1980, 01, 01);
                            com.CommandText = "select N from FileNum where OperDate=?";
                            com.Parameters.Clear();
                            com.Parameters.Add("@OperDate", OleDbType.Date, 40).Value = D;
                            reader = com.ExecuteReader();
                            N = 0;
                            while (reader.Read() == true)
                                N = reader.GetInt32(0) + 1;
                            reader.Close();
                            //Записываем значение счётчика в базу
                            if (N == 0)
                            {
                                N = 1;
                                com.CommandText = "insert into FileNum(N, OperDate) values(?, ?)";
                            }
                            else com.CommandText = "update FileNum set N=? where OperDate=?";
                            com.Parameters.Clear();
                            com.Parameters.Add("@N", OleDbType.Integer, 10).Value = N;
                            com.Parameters.Add("@OperDate", OleDbType.Date, 40).Value = D;
                            com.ExecuteNonQuery();

                            s = DecToStringSS(N, 36);
                            s = s.PadLeft(3, '0');
                            //Префикс
                            string Prefix = "";
                            if ("103,200,202".Contains(MsgCode)) Prefix = "XP";
                            else if ("900,910,940,950".Contains(MsgCode)) Prefix = "XV";
                            else if ("199,299".Contains(MsgCode)) Prefix = "XE";
                            else Prefix = "XS";

                            NewFileName = Prefix + s + GlobalTrash.CodeBnk + '.' + CodeBnk;
                        }

                        //Запишем информацию в базу
                        com.CommandText = "update files set State=?, SysTo=?, FileNameOutput=? where ID=" + id.ToString();
                        com.Parameters.Clear();
                        com.Parameters.Add("@State", OleDbType.VarWChar, 10).Value = "SPFS";
                        com.Parameters.Add("@SysTo", OleDbType.VarWChar, 5).Value = "СПФС";
                        com.Parameters.Add("@FileNameOutput", OleDbType.VarWChar, 255).Value = NewFileName;
                        com.ExecuteNonQuery();
                        transaction.Commit();

                        //Скопируем файл
                        fi.CopyTo(GlobalTrash.SPFSPathOut + "\\" + NewFileName, true);
                        fi.CopyTo(GlobalTrash.SPFSPathArc + "\\" + NewFileName, true);
                        fi.Delete();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        Result = ex.Message;
                    }
                }
            }
            connection.Close();
            return Result;
        }
    }
}