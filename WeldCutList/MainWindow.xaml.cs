using SolidWorks.Interop.sldworks;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using WeldCutList.ViewModel;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace WeldCutList
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        CancellationTokenSource cancellationTokenSource;
        CancellationToken cancellationToken;
        DrawingViewModel drawingViewModel;

        public MainWindow()
        {
            InitializeComponent();
            btnDuplicate.Click += Button_Click_1;
            btnDuplicate.Click += Button_Click_4;
            drawingViewModel = new DrawingViewModel();
            DataContext = drawingViewModel;
            this.Loaded += (s, e) =>
            {
                var screen = System.Windows.SystemParameters.WorkArea;
                this.Left = screen.Right - this.Width;
                this.Top = (screen.Height - this.Height) / 2;
            };
        }

        /// <summary>
        /// (write to DB)from cut-list folders and how to get the custom properties for the solid bodies
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        async private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            this.progressBar1.IsIndeterminate = true;
            cancellationTokenSource = new CancellationTokenSource();
            cancellationToken = cancellationTokenSource.Token;

            try
            {
            await Task.Run(() =>
            {
                var macro = new Macro1CSharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                macro.Main();

                string excelFilePath = "cutlist.xlsx";
                List<CutList> cutLists = new List<CutList>();

                if (File.Exists(excelFilePath))
                {
                    using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[1];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            cutLists.Add(new CutList
                            {
                                Folder_Name = worksheet.Cells[row, 1].Value?.ToString(),
                                Body_Name = worksheet.Cells[row, 2].Value?.ToString(),
                                MaterialProperty = worksheet.Cells[row, 3].Value?.ToString()
                            });
                        }
                    }
                }

                var query = cutLists.Select(product => new { product.Folder_Name, product.Body_Name, product.MaterialProperty });

                Application.Current.Dispatcher.Invoke(() =>
                {
                    dataGrid1.ItemsSource = query.ToList();
                    textBox1.Text = "View 的数量是: " + query.Count().ToString();
                });
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Operation was canceled.");
            }

            this.progressBar1.IsIndeterminate = false;
            this.progressBar1.Value = 100;
        }

        /// <summary>
        /// check the DB
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            string excelFilePath = "cutlist.xlsx";
            List<CutList> cutLists = new List<CutList>();

            if (File.Exists(excelFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        cutLists.Add(new CutList
                        {
                            Folder_Name = worksheet.Cells[row, 1].Value?.ToString(),
                            Body_Name = worksheet.Cells[row, 2].Value?.ToString(),
                            MaterialProperty = worksheet.Cells[row, 3].Value?.ToString()
                        });
                    }
                }
            }

            var query = cutLists.Select(product => new { product.Folder_Name, product.Body_Name, product.MaterialProperty });

            dataGrid1.ItemsSource = query.ToList();

            Application.Current.Dispatcher.Invoke(() =>
            {
                textBox1.Text = "是: " + query.Count().ToString();
            });
        }

        /// <summary>
        /// duplicate views from an existing view, 当前是16个, 改成了30个
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private  void Button_Click_1(object sender, RoutedEventArgs e)
        {
                var macro = new InsertUnfoldedView_CSharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                //用于复制多个视图
                macro.Main();
        }

        /// <summary>
        /// copy and paste drawing sheets. 改成 view/30
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Button_Click_4(object sender, RoutedEventArgs e)
        {
            this.progressBar1.IsIndeterminate = true;
            cancellationTokenSource = new CancellationTokenSource();
            cancellationToken = cancellationTokenSource.Token;

            try
            {
            await Task.Run(() =>
            {
                string excelFilePath = "cutlist.xlsx";
                List<CutList> cutLists = new List<CutList>();

                if (File.Exists(excelFilePath))
                {
                    using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[1];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            cutLists.Add(new CutList
                            {
                                Folder_Name = worksheet.Cells[row, 1].Value?.ToString(),
                                Body_Name = worksheet.Cells[row, 2].Value?.ToString()
                            });
                        }
                    }
                }

                var query = cutLists.Select(product => new { product.Folder_Name, product.Body_Name });
                var macro = new CopyAndPasteCsharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                var sheetQuantity = Math.Ceiling(Convert.ToDouble(query.Count()) / 30);
                Console.WriteLine(query.Count() + "   " + Convert.ToDouble(query.Count()) / 30 + "  " + sheetQuantity);
                macro.Main(query.Count() / 30);
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Operation was canceled.");
            }

            this.progressBar1.IsIndeterminate = false;
            this.progressBar1.Value = 100;
        }

        /// <summary>
        /// Traverse the drawing view
        /// </summary>
        /// <param="sender"></param>
        /// <param name="e"></param>
        async private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            //this.btn2.IsEnabled = false;
            this.progressBar1.IsIndeterminate = true;
            cancellationTokenSource = new CancellationTokenSource();
            cancellationToken = cancellationTokenSource.Token;

            try
            {
            await Task.Run(() =>
            {
                var macro = new CenterOfMass_CSharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                //最核心的方法:获取新的集合,气球, 摆正等功能
                macro.Main(drawingViewModel);
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Operation was canceled.");
            }

            this.progressBar1.IsIndeterminate = false;
            this.progressBar1.Value = 100;
        }

        async private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            //this.btn2.IsEnabled = false;
            this.progressBar1.IsIndeterminate = true;
            cancellationTokenSource = new CancellationTokenSource();
            cancellationToken = cancellationTokenSource.Token;

            try
            {
            await Task.Run(() =>
            {
                var macro = new Dimensioning.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                macro.Main();
                }, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Operation was canceled.");
            }

            this.progressBar1.IsIndeterminate = false;
            this.progressBar1.Value = 100;
        }

        private void Button_Click_Cancel(object sender, RoutedEventArgs e)
        {
            if (cancellationTokenSource != null)
            {
                cancellationTokenSource.Cancel();
            }
        }
    }
}
