using Microsoft.WindowsAPICodePack.Dialogs;
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
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace DigitalDataCopy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void FolderFromButton_Click(object sender, RoutedEventArgs e)
        {
            MyOpenFolder(sender, e, FolderFromTextBox);
        }


        private void xlsFromButton_Click(object sender, RoutedEventArgs e)
        {
            MyOpenFolder(sender, e, xlsFromTextBox);
        }

        private void MyOpenFolder(object sender, RoutedEventArgs e, TextBox ResultTexbox)
        {
            var screen = new CommonOpenFileDialog();
            screen.IsFolderPicker = true;
            var folderOpen = screen.ShowDialog();

            if (folderOpen == CommonFileDialogResult.Ok)
            {
                ResultTexbox.Text = screen.FileName;
            }
        }

        private void FolderToButton_Click(object sender, RoutedEventArgs e)
        {
            MyOpenFolder(sender, e, FolderToTextBox);
        }

        private void xlsToButton_Click(object sender, RoutedEventArgs e)
        {
            MyOpenFolder(sender, e, xlsToTextBox);
        }

        private async void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CopyButton.IsEnabled = false;
                CopyButton.Content = "COPYING...";

                var subFolder = $"{ SubfolderPrefixTextBox.Text.Trim() }{ SubfolderPostFixTextBox.Text.Trim()}";
                var SourceFolderPath = $"{FolderFromTextBox.Text.Trim()}\\{subFolder}";
                var targetFolderPath = $"{FolderToTextBox.Text.Trim()}\\{subFolder}";

                var excelFileExt = ExcelExtendsionTextBox.Text.Trim();
                var excelFileName = $"{xlsPrefixTextBox.Text.Trim()}{xlsPostfixTextBox.Text.Trim()}{excelFileExt}";
                var SourceExcelFile = $"{xlsFromTextBox.Text.Trim()}\\{excelFileName}";
                var TargetxcelFile = $"{xlsToTextBox.Text.Trim()}\\{excelFileName}";

                CanCopy(SourceFolderPath, SourceExcelFile);

                Directory.CreateDirectory(targetFolderPath);

                await  Task.Run(() => FileSystem.CopyFile(SourceExcelFile, TargetxcelFile, 
                                                           UIOption.AllDialogs, UICancelOption.DoNothing));

                await Task.Run(() => FileSystem.CopyDirectory(SourceFolderPath,
                                                               targetFolderPath,
                                                               UIOption.AllDialogs,
                                                               UICancelOption.DoNothing));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CopyButton.IsEnabled = true;
                CopyButton.Content = "Press Enter Or Click To Copy";
            }

        }

        public void CanCopy(string subFolderPath, string sourceExcelFile)
        {
            var checkStr = CheckEmptyTextBox();
            if (checkStr != string.Empty)
            {

                throw new Exception(checkStr);
            }

            if (Directory.Exists(subFolderPath) == false)
            {

                throw new Exception($"Source Sub Folder Not Exits: {subFolderPath}");
            }

            if (File.Exists(sourceExcelFile) == false)
            {

                throw new Exception($"Source Excel File Not Exits: {sourceExcelFile}");
            }

        }
        private string CheckEmptyTextBox()
        {
            if (FolderFromTextBox.Text == string.Empty)
            {
                return "Please Input Folder From First";
            }

            if (SubfolderPrefixTextBox.Text == string.Empty || SubfolderPostFixTextBox.Text == string.Empty)
            {
                return "Please Input Sub Folder First";
            }

            if (FolderToTextBox.Text == string.Empty)
            {
                return "Please Input Source Folder First";
            }
            if (xlsFromTextBox.Text == string.Empty)
            {
                return "Please Input Source Excel Folder First";
            }
            if (xlsPrefixTextBox.Text == string.Empty)
            {
                return "Please Input Prefix Excel File Name First";
            }
            if (xlsPostfixTextBox.Text == string.Empty)
            {
                return "Please Input Postfix Excel File Name First";
            }
            if (xlsToTextBox.Text == string.Empty)
            {
                return "Please Input Target Excel Folder First";

            }

            return string.Empty;
        }

        private void SubfolderPrefixTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (AutoFillExcelNameCheckBox.IsChecked == true)
            {
                xlsPrefixTextBox.Text = SubfolderPrefixTextBox.Text;
            }

        }

        private void SubfolderPostFixTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (AutoFillExcelNameCheckBox.IsChecked == true)
            {
                xlsPostfixTextBox.Text = SubfolderPostFixTextBox.Text;
            }

        }

        private void ContactButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("If You Need Support, Please Email To hieutrinh87@gmail.com\nThanks", "Question", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
