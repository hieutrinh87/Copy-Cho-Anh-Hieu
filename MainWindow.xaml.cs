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

        private async void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveConfig();
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

                await Task.Run(() => FileSystem.CopyFile(SourceExcelFile, TargetxcelFile,
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
                PrefixScreenFileTextBox.Text = SubfolderPrefixTextBox.Text;
            }

        }

        private void SubfolderPostFixTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (AutoFillExcelNameCheckBox.IsChecked == true)
            {
                xlsPostfixTextBox.Text = SubfolderPostFixTextBox.Text;
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

            _config.AutoFill = (bool)AutoFillExcelNameCheckBox.IsChecked;

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
                AutoFillExcelNameCheckBox.IsChecked = _config.AutoFill;
            }
        }

        private void AutoFillExcelNameCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            EnableOfDisableTextbox(false);
        }


        private void AutoFillExcelNameCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            EnableOfDisableTextbox(true);
        }
        private void EnableOfDisableTextbox(bool Status)
        {
            xlsPrefixTextBox.IsEnabled = Status;
            xlsPostfixTextBox.IsEnabled = Status;
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
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message , "Screen Shot", MessageBoxButton.OK, MessageBoxImage.Information);
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
