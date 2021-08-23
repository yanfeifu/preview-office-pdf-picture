using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;
using Path = System.IO.Path;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Xps.Packaging;

namespace PreviewDocPicWpf
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();

            
            //Page_Change.Content = new Frame()
            //{
            //    Content = pdf
            //};
        }
        
        public void LoadFile(object sender, RoutedEventArgs e)
        {
            //officePage office = new officePage();
            //Page_Change.Content = new Frame()
            //{
            //    Content = office
            //};
            pdfPage pdf = new pdfPage();
            pdf.Loadpdf();
            Page_Change.Content = new Frame()
            {
                Content = pdf
            };
        }

        //private void Button_Click(object sender, RoutedEventArgs e)
        //{
        //    //PageChange();
        //    officePage office = new officePage();
        //    Page_Change.Content = new Frame()
        //    {
        //        Content = office
        //    };
        //}

        //private void PageChange()
        //{
        //    officePage office = new officePage();
        //    Page_Change.Content = new Frame()
        //    {
        //        Content = office
        //    };
        //}

        #region 预览图片
        //public void PreviewFile(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog openFileDialog = new OpenFileDialog();
        //    if (openFileDialog.ShowDialog() == true)
        //    {
        //        Uri fileUri = new Uri(openFileDialog.FileName);
        //        this.img.Source = new BitmapImage(fileUri);
        //    }
        //}
        #endregion

        #region 预览pdf
        //private void PreviewFile(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog dlg = new OpenFileDialog();
        //    dlg.CheckFileExists = true;
        //    dlg.Filter = "PowerPoint Format (*.pdf)|*.pdf|" +
        //                 "All files (*.*)|*.*";
        //    if ((bool)dlg.ShowDialog(this))
        //    {
        //        string filePath = dlg.FileName;
        //        pdfWebBrowser.Navigate(filePath);
        //    }
        //}
        #endregion


    }
}



