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

namespace PreviewDocPicWpf
{
    /// <summary>
    /// pdfPage.xaml 的交互逻辑
    /// </summary>
    public partial class pdfPage : Page
    {
        public pdfPage()
        {
            InitializeComponent();
        }

        public void Loadpdf()
        {
            var dialog = new OpenFileDialog();
            string filePath = dialog.FileName;
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                try
                {
                    moonPdfPanel.OpenFile(filePath);
                    _isLoaded = true;
                }
                catch (Exception)
                {
                    _isLoaded = false;
                }
            }
        }

        #region 预览pdf
        private bool _isLoaded = false;
        private void FileButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();

            if (dialog.ShowDialog().GetValueOrDefault())
            {
                string filePath = dialog.FileName;

                try
                {
                    moonPdfPanel.OpenFile(filePath);
                    _isLoaded = true;
                }
                catch (Exception)
                {
                    _isLoaded = false;
                }
            }
        }

        private void ZoomInButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isLoaded)
            {
                moonPdfPanel.ZoomIn();
            }
        }

        private void ZoomOutButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isLoaded)
            {
                moonPdfPanel.ZoomOut();
            }
        }

        private void NormalButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isLoaded)
            {
                moonPdfPanel.Zoom(1.0);
            }
        }

        private void FitToHeightButton_Click(object sender, RoutedEventArgs e)
        {
            moonPdfPanel.ZoomToHeight();
        }

        private void FacingButton_Click(object sender, RoutedEventArgs e)
        {
            moonPdfPanel.ViewType = MoonPdfLib.ViewType.Facing;
        }

        private void SinglePageButton_Click(object sender, RoutedEventArgs e)
        {
            moonPdfPanel.ViewType = MoonPdfLib.ViewType.SinglePage;
        }
        #endregion
    }
}
