using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace IMcorrectionTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);
        List<Warming> WarningListCurrentMonth { get; set; }
        List<Warming> WarningListLastMonth { get; set; }
        List<Warming> WarningListKGID { get; set; }
        List<Warming> WarningListItog { get; set; }
        string CurrentMonthFileName { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            WarningListItog = new List<Warming>();
            updProgress = new UpdateProgressBarDelegate(progressBar.SetValue);
        }
        private void InsertRDUResult(string RduName)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Книга Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                var warningList = WarningFarm.GetWarningListFromCduFormatExcel(filePath);

                foreach (var wrn in warningList)
                {
                    var kgidWrn = WarningListKGID.Where(x => x.ID == wrn.ID).FirstOrDefault();
                    if (kgidWrn != null)
                    {
                        if (kgidWrn.ModelingAuthoritySet == RduName)
                        {
                            kgidWrn.Comment = wrn.Comment;
                        }
                    }
                }

            }
        }
        private void CopyPreviousCommentsToCurren()
        {
            UpdateProgressBarDelegate updProgress = new UpdateProgressBarDelegate(progressBar.SetValue);
            double value = 0;
            progressBar.Maximum = WarningListLastMonth.Count;
            Task.Run(() =>
            {
                foreach (var wrn in WarningListLastMonth)
                {
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++value });
                    var curr = WarningListCurrentMonth.Where(x => x.ID == wrn.ID).FirstOrDefault();
                    if (curr != null)
                    {
                        curr.IsNewInMonth = false;
                        curr.PreviousComment = wrn.Comment;
                    }
                }
                value = 0;
                Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value });
            });
        }
        //private void CopyPreviousCommentsToKGIG()
        //{
        //    foreach (var wrn in WarningListLastMonth)
        //    {
        //        var curr = WarningListCurrentMonth.Where(x => x.ID == wrn.ID).FirstOrDefault();
        //        if (curr != null)
        //        {
        //            curr.IsNewInMonth = false;
        //            curr.PreviousComment = wrn.Comment;
        //        }
        //    }
        //}

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Книга Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                currentMonthSatus.Items.Clear();
                currentMonthSatus.Items.Add(new TextBlock() { Text = $"Загрузка файла..." });
                CurrentMonthFileName = openFileDialog.FileName;
                WarningListCurrentMonth = WarningFarm.GetWarningListFromCduFormatExcel(CurrentMonthFileName);
                dataGridWarning.ItemsSource = WarningListCurrentMonth;
                currentMonthSatus.Items.Clear();
                currentMonthSatus.Items.Add(new TextBlock() { Text = $"Всего по ОЗ ОДУ Урала: {WarningListCurrentMonth.Count()}" });

                foreach (var dc in WarningListCurrentMonth.Select(x => x.ModelingAuthoritySet).Distinct())
                {
                    currentMonthSatus.Items.Add(new Separator());
                    currentMonthSatus.Items.Add(new TextBlock() { Text = $"{dc}: {WarningListCurrentMonth.Count(x => x.ModelingAuthoritySet == dc)}" });

                }
            }

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Книга Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                currentMonthSatus.Items.Clear();
                currentMonthSatus.Items.Add(new TextBlock() { Text = $"Загрузка файла..." });

                string filePath = openFileDialog.FileName;
                WarningListLastMonth = WarningFarm.GetWarningListFromCduFormatExcel(filePath);
                dataGridWarningLastMonth.ItemsSource = WarningListLastMonth;

                lastMonthSatus.Items.Clear();
                lastMonthSatus.Items.Add(new TextBlock() { Text = $"Всего по ОЗ ОДУ Урала: {WarningListLastMonth.Count()}" });

                foreach (var dc in WarningListLastMonth.Select(x => x.ModelingAuthoritySet).Distinct())
                {
                    lastMonthSatus.Items.Add(new Separator());
                    lastMonthSatus.Items.Add(new TextBlock() { Text = $"{dc}: {WarningListLastMonth.Count(x => x.ModelingAuthoritySet == dc)}" });

                }
            }

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV (разделитель точка с запятой) (*.csv)|*.csv|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                currentMonthSatus.Items.Clear();
                currentMonthSatus.Items.Add(new TextBlock() { Text = $"Загрузка файла..." });

                string filePath = openFileDialog.FileName;
                WarningListKGID = WarningFarm.GetWarningListFromCK11Format(filePath);
                dataGridWarningKGID.ItemsSource = WarningListKGID;

                kgidSatus.Items.Clear();
                kgidSatus.Items.Add(new TextBlock() { Text = $"Всего по ОЗ ОДУ Урала: {WarningListKGID.Count()}" });

                foreach (var dc in WarningListKGID.Select(x => x.ModelingAuthoritySet).Distinct())
                {
                    kgidSatus.Items.Add(new Separator());
                    kgidSatus.Items.Add(new TextBlock() { Text = $"{dc}: {WarningListKGID.Count(x => x.ModelingAuthoritySet == dc)}" });

                }
            }
        }

        private void button4_Click(object sender, RoutedEventArgs e)
        {
            if (WarningListLastMonth != null && WarningListCurrentMonth != null && WarningListLastMonth.Count() > 0 && WarningListCurrentMonth.Count() > 0)
            {
                CopyPreviousCommentsToCurren();
                colPrevCommentCurrTable.Visibility = Visibility.Visible;
            }
            else
            {
                MessageBox.Show("Файл текущего или прошлого месяца не выбран или не содержит записей. Перенос не выполнен.", "Ошибка");
            }
        }
        private UpdateProgressBarDelegate updProgress;
        private double value;
        private void button5_Click(object sender, RoutedEventArgs e)
        {
            currentMonthSatus.Items.Clear();
            currentMonthSatus.Items.Add(new TextBlock() { Text = $"Формирование итога..." });
            WarningListItog = new List<Warming>();
            value = 0;
            progressBar.Maximum = WarningListCurrentMonth.Count + WarningListKGID.Count;
            Task.Run(() => 
            {
                foreach (var wrn in WarningListCurrentMonth)
                {
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++value });
                    WarningListItog.Add(wrn);
                    var lastMonthWrn = WarningListLastMonth?.Where(x => x.ID == wrn.ID).FirstOrDefault();
                    if (lastMonthWrn != null)
                        wrn.PreviousComment = lastMonthWrn.Comment;

                    var kgidWrn = WarningListKGID.Where(x => x.ID == wrn.ID).FirstOrDefault();
                    if (kgidWrn != null)
                    {
                        wrn.IsCorrectedInKGID = false;
                        wrn.Comment = kgidWrn.Comment;
                    }
                    else
                    {
                        wrn.IsCorrectedInKGID = true;
                        wrn.Comment = "Устранено";
                    }

                }
                foreach (var wrn in WarningListKGID)
                {
                    Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, ++value });
                    var currMonthWrn = WarningListCurrentMonth.Where(x => x.ID == wrn.ID).FirstOrDefault();
                    if (currMonthWrn == null)
                    {
                        wrn.IsNewInKGID = true;
                        wrn.IsCorrectedInKGID = false;
                        WarningListItog.Add(wrn);
                    }
                }
            }).ContinueWith(SetDataGrid, TaskScheduler.FromCurrentSynchronizationContext());
        }
        private void SetDataGrid(Task obj)
        {
            dataGridWarningItog.ItemsSource = WarningListItog;
            value = 0;
            Dispatcher.Invoke(updProgress, new object[] { ProgressBar.ValueProperty, value });
            itogSatus.Items.Clear();
            itogSatus.Items.Add(new TextBlock() { Text = $"Всего по ОЗ ОДУ Урала: {WarningListItog.Count()} (Н:{WarningListItog.Count(x => x.IsNewInKGID)} И:{WarningListItog.Count(x => x.IsCorrectedInKGID)})" });

            foreach (var dc in WarningListItog.Select(x => x.ModelingAuthoritySet).Distinct())
            {
                itogSatus.Items.Add(new Separator());
                itogSatus.Items.Add(new TextBlock() { Text = $"{dc}: {WarningListItog.Count(x => x.ModelingAuthoritySet == dc)} (Н:{WarningListItog.Count(x => x.ModelingAuthoritySet == dc && x.IsNewInKGID)} И:{WarningListItog.Count(x => x.ModelingAuthoritySet == dc && x.IsCorrectedInKGID)})" });
            }
        }
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            InsertRDUResult((sender as MenuItem).Header.ToString());
        }

        private void buttonExport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Книга Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (!string.IsNullOrEmpty(CurrentMonthFileName))
            {
                int shortNameStartIndex = CurrentMonthFileName.LastIndexOf(@"\") + 1;
                int length = CurrentMonthFileName.Length - 5;
                saveFileDialog.FileName = CurrentMonthFileName.Substring(shortNameStartIndex, length - shortNameStartIndex) + " (Комментарии ОДУ Урала)";
            }
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    var filePath = saveFileDialog.FileName;
                    WarningFarm.SaveToExcelBasedOnCurrentMonth(CurrentMonthFileName, filePath, WarningListItog);
                    MessageBox.Show("Сохранение успешно завершено", "Готово");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка при сохранении файла");
                }
            }
        }
    }
}
