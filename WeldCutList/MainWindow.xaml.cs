using SolidWorks.Interop.sldworks;
using System.Windows;
using System.Data.Entity.Core;
using System.Linq;
using Microsoft.VisualBasic.Logging;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using System.Threading;

namespace WeldCutList
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        CancellationTokenSource cancellationTokenSource;
        CancellationToken cancellationToken;

        public MainWindow()
        {
            InitializeComponent();
            //btn2.Click += Button_Click_2;
            //btn2.Click += Button_Click_3;
        }

        /// <summary>
        /// (write to DB)from cut-list folders and how to get the custom properties for the solid bodies
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        async private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            //this.btn2.IsEnabled = false;
            this.progressBar1.IsIndeterminate = true;

            await Task.Run(() =>
            {
                var macro = new Macro1CSharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                macro.Main();
            });

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
            using (CutListSample01Entities cutListSample01Entities1 = new CutListSample01Entities())
            {
                var query =
                    from product in cutListSample01Entities1.CutLists
                        //where product.Color == "Red"
                        //orderby product.ListPrice
                    select new { product.Folder_Name, product.Body_Name, product.MaterialProperty };

                dataGrid1.ItemsSource = query.ToList();

                Application.Current.Dispatcher.Invoke(() =>
                {
                    textBox1.Text =
                    "View 的数量是: "+
                    query.Count().ToString()
                    ;
                });
            }
        }

        /// <summary>
        /// duplicate views from an existing view, 当前是16个, 改成了30个
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var macro = new InsertUnfoldedView_CSharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
            macro.Main();
        }

        /// <summary>
        /// copy and paste drawing sheets. 改成 view/30
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            using (CutListSample01Entities cutListSample01Entities1 = new CutListSample01Entities())
            {
                var query =
                    from product in cutListSample01Entities1.CutLists
                        //where product.Color == "Red"
                        //orderby product.ListPrice
                    select new { product.Folder_Name, product.Body_Name };
                var macro = new CopyAndPasteCsharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                var sheetQuantity = Math.Ceiling(Convert.ToDouble(query.Count()) / 30);
                Console.WriteLine(query.Count() + "   " + Convert.ToDouble(query.Count()) / 30 + "  " + sheetQuantity);
                macro.Main(query.Count() / 30);
            }
        }

        /// <summary>
        /// Traverse the drawing view
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        async private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            //this.btn2.IsEnabled = false;
            this.progressBar2.IsIndeterminate = true;

            await Task.Run(() =>
            {
                var macro = new CenterOfMass_CSharp.csproj.SolidWorksMacro() { swApp = new SldWorks() };
                macro.Main();
            });

            this.progressBar2.IsIndeterminate = false;
            this.progressBar2.Value = 100;
        }

    }
}
