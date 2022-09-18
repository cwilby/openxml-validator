using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace OpenXMLValidator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<ValidationErrorInfo> errors = new ObservableCollection<ValidationErrorInfo>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();

            this.dataGrid.ItemsSource = errors;

            if (ofd.ShowDialog() == true)
            {
                textFile.Text = ofd.FileName;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var filename = textFile.Text;

            using (var doc = WordprocessingDocument.Open(filename, true))
            {
                try
                {
                    var validator = new OpenXmlValidator();

                    errors = new ObservableCollection<ValidationErrorInfo>(validator.Validate(doc));

                    dataGrid.ItemsSource = errors;

                    MessageBox.Show("Found " + errors.Count() + " errors");
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                doc.Close();
            }
        }
    }
}
