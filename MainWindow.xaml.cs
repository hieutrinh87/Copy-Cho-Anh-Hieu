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
using Newtonsoft.Json;
using System.Drawing;
using System.Diagnostics;

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

        private async void CopyAllButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveConfig();
                CopyAllButton.IsEnabled = false;
                CopyDataButton.IsEnabled = false;
                CopyExcelButton.IsEnabled = false;

                CheckBeforeCopyData();
                CheckBeforeCopyExcelFile();

                await CopyData();
                await CopyExcel();
                await Task.Delay(1000);
                MessageBox.Show("Data and Excel file is copied", "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CopyAllButton.IsEnabled = true;
                CopyDataButton.IsEnabled = true;
                CopyExcelButton.IsEnabled = true;
            }
        }
        private async void CopyDataButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveConfig();
                CopyDataButton.IsEnabled = false;

                CheckBeforeCopyData();
                await CopyData();
                await Task.Delay(1000);
                MessageBox.Show("Data is copied", "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CopyDataButton.IsEnabled = true;
            }
        }

        public void CheckBeforeCopyData()
        {
            if (FolderFromTextBox.Text == string.Empty)
            {
                throw new Exception("Please Input Folder From First");
            }

            if (Directory.Exists(FolderFromTextBox.Text) == false)
            {
                throw new Exception($"Source Sub Folder Not Exits: {FolderFromTextBox.Text}");
            }

            if (SubfolderPrefixTextBox.Text == string.Empty || SubfolderPostFixTextBox.Text == string.Empty)
            {
                throw new Exception("Please Input Sub Folder First");
            }

            var subFolder = $"{ SubfolderPrefixTextBox.Text.Trim() }{ SubfolderPostFixTextBox.Text.Trim()}";
            var SourceFolderPath = $"{FolderFromTextBox.Text.Trim()}\\{subFolder}";

            if (Directory.Exists(SourceFolderPath) == false)
            {
                throw new Exception($"Source Sub Folder Not Exits: {SourceFolderPath}");
            }

            if (FolderToTextBox.Text == string.Empty)
            {
                throw new Exception("Please Input Source Folder First");
            }
            if (Directory.Exists(FolderToTextBox.Text.Trim()) == false)
            {
                throw new Exception($"Target Folder Not Exits: {FolderToTextBox.Text.Trim()}");
            }
        }

        public async Task<bool> CopyData()
        {
            try
            {
                var subFolder = $"{ SubfolderPrefixTextBox.Text.Trim() }{ SubfolderPostFixTextBox.Text.Trim()}";
                var SourceFolderPath = $"{FolderFromTextBox.Text.Trim()}\\{subFolder}";
                var targetFolderPath = $"{FolderToTextBox.Text.Trim()}\\{subFolder}";

                Directory.CreateDirectory(targetFolderPath);
                await Task.Run(() => FileSystem.CopyDirectory(SourceFolderPath,
                                                   targetFolderPath,
                                                   UIOption.AllDialogs,
                                                   UICancelOption.DoNothing));
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private async void CopyExcelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveConfig();
                CopyExcelButton.IsEnabled = false;

                CheckBeforeCopyExcelFile();
                await CopyExcel();
                await Task.Delay(1000);
                MessageBox.Show("Excel file is copied", "Copy", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Copy", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally
            {
                CopyExcelButton.IsEnabled = true;
            }
        }

        private void CheckBeforeCopyExcelFile()
        {

            if (xlsFromTextBox.Text == string.Empty)
            {
                throw new Exception("Please Input Source Excel Folder First");
            }
            if (Directory.Exists(xlsFromTextBox.Text) == false)
            {
                throw new Exception($"Source Excel Folder Not Exits: {xlsFromTextBox.Text}");
            }
            if (xlsPrefixTextBox.Text == string.Empty)
            {
                throw new Exception("Please Input Prefix Excel File Name First");
            }
            if (xlsPostfixTextBox.Text == string.Empty)
            {
                throw new Exception("Please Input Postfix Excel File Name First");
            }


            var excelFileExt = ExcelExtendsionTextBox.Text.Trim();
            var excelFileName = $"{xlsPrefixTextBox.Text.Trim()}{xlsPostfixTextBox.Text.Trim()}{excelFileExt}";
            var sourceExcelFile = $"{xlsFromTextBox.Text.Trim()}\\{excelFileName}";

            if (File.Exists(sourceExcelFile) == false)
            {
                throw new Exception($"Source Excel File Not Exits: {sourceExcelFile}");
            }

            if (xlsToTextBox.Text.Trim() == string.Empty)
            {
                throw new Exception("Please Input Target Excel Folder First");
            }
            if (Directory.Exists(xlsToTextBox.Text) == false)
            {
                throw new Exception($"Target Excel Folder Not Exits: {xlsToTextBox.Text}");
            }
        }
        public async Task<bool> CopyExcel()
        {
            try
            {
                var excelFileExt = ExcelExtendsionTextBox.Text.Trim();
                var excelFileName = $"{xlsPrefixTextBox.Text.Trim()}{xlsPostfixTextBox.Text.Trim()}{excelFileExt}";
                var SourceExcelFile = $"{xlsFromTextBox.Text.Trim()}\\{excelFileName}";
                var TargetxcelFile = $"{xlsToTextBox.Text.Trim()}\\{excelFileName}";


                await Task.Run(() => FileSystem.CopyFile(SourceExcelFile, TargetxcelFile,
                                                          UIOption.AllDialogs, UICancelOption.DoNothing));
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private void SubfolderPrefixTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (AutoFillExcelNameCheckBox.IsChecked == true)
            {
                xlsPrefixTextBox.Text = SubfolderPrefixTextBox.Text;
            }

            if (AutoFillImageNameCheckBox.IsChecked == true)
            {
                PrefixScreenFileTextBox.Text = SubfolderPrefixTextBox.Text;
            }
        }


        private void SubfolderPostFixTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (AutoFillExcelNameCheckBox.IsChecked == true)
            {
                xlsPostfixTextBox.Text = SubfolderPostFixTextBox.Text;

            }

            if (AutoFillImageNameCheckBox.IsChecked == true)
            {
                PostfixScreenFileTextBox.Text = SubfolderPostFixTextBox.Text;
            }
        }

        private void ContactButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("If You Need Support, Please Email To hieutrinh87@gmail.com\nThanks", "Question", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfigIfExits();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            SaveConfig();
        }

        public void SaveConfig()
        {
            Config _config = new Config();

            _config.FolderFromTextBox = FolderFromTextBox.Text;
            _config.SubfolderPrefixTextBox = SubfolderPrefixTextBox.Text;
            _config.SubfolderPostFixTextBox = SubfolderPostFixTextBox.Text;

            _config.xlsFromTextBox = xlsFromTextBox.Text;
            _config.xlsPrefixTextBox = xlsPrefixTextBox.Text;
            _config.xlsPostfixTextBox = xlsPostfixTextBox.Text;
            _config.ExcelExtendsionTextBox = ExcelExtendsionTextBox.Text;

            _config.FolderToTextBox = FolderToTextBox.Text;

            _config.xlsToTextBox = xlsToTextBox.Text;

            _config.PrintSreenFolder = PrintSreenFolderTextBox.Text;
            _config.PrefixScreenFile = PrefixScreenFileTextBox.Text;
            _config.PostfixScreenFile = PostfixScreenFileTextBox.Text;

            _config.AutoFillExcel = (bool)AutoFillExcelNameCheckBox.IsChecked;
            _config.AutoFillImage = (bool)AutoFillImageNameCheckBox.IsChecked;

            string ConfigStr = JsonConvert.SerializeObject(_config);

            File.WriteAllText("Config.JSON", ConfigStr);
        }

        public void LoadConfigIfExits()
        {
            var ConfigFilePath = "Config.JSON";
            if (File.Exists(ConfigFilePath))
            {
                var CongifStr = File.ReadAllText(ConfigFilePath);

                Config _config = new Config();
                _config = JsonConvert.DeserializeObject<Config>(CongifStr);

                FolderFromTextBox.Text = _config.FolderFromTextBox;
                SubfolderPrefixTextBox.Text = _config.SubfolderPrefixTextBox;
                SubfolderPostFixTextBox.Text = _config.SubfolderPostFixTextBox;

                xlsFromTextBox.Text = _config.xlsFromTextBox;
                xlsPrefixTextBox.Text = _config.xlsPrefixTextBox;
                xlsPostfixTextBox.Text = _config.xlsPostfixTextBox;
                ExcelExtendsionTextBox.Text = _config.ExcelExtendsionTextBox;

                FolderToTextBox.Text = _config.FolderToTextBox;

                xlsToTextBox.Text = _config.xlsToTextBox;

                PrintSreenFolderTextBox.Text = _config.PrintSreenFolder;
                PrefixScreenFileTextBox.Text = _config.PrefixScreenFile;
                PostfixScreenFileTextBox.Text = _config.PostfixScreenFile;
                AutoFillExcelNameCheckBox.IsChecked = _config.AutoFillExcel;
                AutoFillImageNameCheckBox.IsChecked = _config.AutoFillImage;
            }
        }

        private void AutoFillExcelNameCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            EnableOfDisableExcelTextbox(false);

            xlsPrefixTextBox.Text = SubfolderPrefixTextBox.Text;
            xlsPostfixTextBox.Text = SubfolderPostFixTextBox.Text;
        }


        private void AutoFillExcelNameCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            EnableOfDisableExcelTextbox(true);
        }
        private void EnableOfDisableExcelTextbox(bool Status)
        {
            xlsPrefixTextBox.IsEnabled = Status;
            xlsPostfixTextBox.IsEnabled = Status;
        }

        private void AutoFillImageNameCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            EnableOfDisableImageTexttbox(false);
            PrefixScreenFileTextBox.Text = SubfolderPrefixTextBox.Text;
            PostfixScreenFileTextBox.Text = SubfolderPostFixTextBox.Text;
            PostfixScreenFileTextBox.Text = SubfolderPostFixTextBox.Text;
        }

        private void AutoFillImageNameCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            EnableOfDisableImageTexttbox(true);
        }

        private void EnableOfDisableImageTexttbox(bool Status)
        {
            PostfixScreenFileTextBox.IsEnabled = Status;
            PrefixScreenFileTextBox.IsEnabled = Status;
        }

        private void ScreenFolderButton_Click(object sender, RoutedEventArgs e)
        {
            MyOpenFolder(sender, e, PrintSreenFolderTextBox);
        }

        private void SolButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveConfig();
                save_ScreenShot_as_File("SOL");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Screen Shot", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        private void EolButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveConfig();
                save_ScreenShot_as_File("EOL");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Screen Shot", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void save_ScreenShot_as_File(string FilePreFix)

        {
            if (PrintSreenFolderTextBox.Text.Trim() == string.Empty)
            {
                throw new Exception($"Please Input Screen Folder First");

            }

            if (PrefixScreenFileTextBox.Text == string.Empty)
            {
                throw new Exception($"Please Input Screen Sub Folder First");

            }
            if (PostfixScreenFileTextBox.Text == string.Empty)
            {
                throw new Exception($"Please Input Screen Sub Folder First");

            }

            var MainFolder = PrintSreenFolderTextBox.Text.Trim();
            if (Directory.Exists(MainFolder) == false)
            {
                throw new Exception($"The Folder Path Not Exists: {MainFolder}");
            }

            var Subfolder = PrefixScreenFileTextBox.Text.Trim() + PostfixScreenFileTextBox.Text.Trim();

            var FullFolderPath = MainFolder + "\\" + Subfolder;

            var FileName = $"{FilePreFix}_{DateTime.Now.ToString("ddMMyyyy_hhmmss")}.png";
            var FullFileName = $"{MainFolder}\\{Subfolder}\\{FileName}";

            int screenLeft = (int)SystemParameters.VirtualScreenLeft;

            int screenTop = (int)SystemParameters.VirtualScreenTop;

            int screenWidth = (int)SystemParameters.VirtualScreenWidth;

            int screenHeight = (int)SystemParameters.VirtualScreenHeight;

            Bitmap bitmap_Screen = new Bitmap(screenWidth, screenHeight);

            Graphics g = Graphics.FromImage(bitmap_Screen);

            g.CopyFromScreen(screenLeft, screenTop, 0, 0, bitmap_Screen.Size);


            Directory.CreateDirectory(FullFolderPath);

            bitmap_Screen.Save(FullFileName);

            //ProcessStartInfo Info = new ProcessStartInfo()
            //{
            //    FileName = "mspaint.exe",
            //    Arguments = FullFileName
            //};
            //Process.Start(Info);

            ProcessStartInfo startInfo = new ProcessStartInfo(FullFileName);
            startInfo.Verb = "edit";

            Process.Start(startInfo);

            // Đây là branch Test

        }


    }
}
