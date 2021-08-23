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

                if (CanCopyFolder(SourceFolderPath) == false)
                {
                    return;
                }

                if (Directory.Exists(targetFolderPath) == false)
                {
                    Directory.CreateDirectory(targetFolderPath);
                }

                foreach (string filename in Directory.EnumerateFiles(SourceFolderPath))
                {
                    using (FileStream SourceStream = File.Open(filename, FileMode.Open))
                    {
                        var ShortName = filename.Substring(filename.LastIndexOf('\\'));
                        using (FileStream DestinationStream = File.Create(targetFolderPath + ShortName))
                        {
                            await SourceStream.CopyToAsync(DestinationStream);
                        }
                    }
                }
                MessageBox.Show("Copy Successfuly", "Copy",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CopyButton.IsEnabled = true;
                CopyButton.Content = "Press Enter Or Click To Copy";
            }

        }

        public bool CanCopyFolder(string subFolderPath)
        {
            var checkStr = CheckEmptyTextBox();
            if (checkStr != string.Empty)
            {
                MessageBox.Show(checkStr, "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            return true;

            if (Directory.Exists(subFolderPath) == false)
            {
                MessageBox.Show("Target Sub Folder Not Exits", "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
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
                return "Target Folder Not Exits";
            }

            return string.Empty;
        }
    }
}
