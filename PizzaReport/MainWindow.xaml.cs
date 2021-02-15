using System;
using System.Windows;
using System.Windows.Forms;
using System.Threading.Tasks;
using MilanoExtraReport.BL;

namespace MilanoExtraReport.UI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            if (Type.GetTypeFromProgID("Excel.Application") == null)
            {
                Header.Text = "Ошибка";
                Body.Text = "MS Excel не установлен на данном ПК.";
                return;
            }

            Converter.RowRead += Converter_Read;
            Converter.RowWritten += Converter_Write;
            Converter.Completed += Converter_Converted;
            Converter.ErrorOccurred += Converter_Error;
            this.Loaded += Window_Loaded;
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string fileName = SelectFile();

            if (!string.IsNullOrEmpty(fileName))
            {
                new Task(() => Converter.Convert(fileName)).Start();
                Progress.Visibility = Visibility.Visible;
            }
            else
            {
                Close();
            }
        }

        private string SelectFile()
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Выберите отчет для преобразования",
                Filter = "Excel files (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return dialog.FileName;
            }

            return null;
        }

        private void Converter_Read(int percentСomplete)
        {
            Dispatcher.Invoke(() => {
                Header.Text = "Чтение данных";
                Progress.Value = percentСomplete;
            });
        }

        private void Converter_Write(int number)
        {
            Dispatcher.Invoke(() => {
                Header.Text = "Запись данных";
                if (Progress.Maximum != number)
                {
                    Progress.Value = 0;
                }
                Progress.Maximum = number;
                Progress.Value++;
            });
        }

        private void Converter_Converted(string fileName)
        {
            Dispatcher.Invoke(() => {
                Header.Text = "Отчет готов!";
                Body.Text = fileName;
            });
        }

        private void Converter_Error(Exception ex)
        {
            Dispatcher.Invoke(() => {
                Header.Text = "Ошибка";
                Body.Text = ex.Message;
            });
        }
    }
}
