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

namespace Main
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


        private void TargetWordFileSelectButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog objFileDialog = new System.Windows.Forms.OpenFileDialog();

            objFileDialog.Title = "請選取目標 Word 檔案";
            objFileDialog.InitialDirectory = "C:\\";
            if (objFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.TargetWordFilePathTextBox.Text = objFileDialog.FileName;
                string sUnzipDirecory = string.Empty;
                string sZipFilePath = this.TargetWordFilePathTextBox.Text.Replace("docx", "zip");

                System.IO.File.Copy(this.TargetWordFilePathTextBox.Text, sZipFilePath);
                sUnzipDirecory = ZipHelper.Decompress(sZipFilePath);
                System.IO.File.Delete(sZipFilePath);
                this.ResetTree(sUnzipDirecory);
            }
            else
            {

            }
        }

        #region TreeView

        private object dummyNode = null;

        public string SelectedImagePath { get; set; }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //foreach (string s in Directory.GetLogicalDrives())
            //{
            //    TreeViewItem item = new TreeViewItem();
            //    item.Header = s;
            //    item.Tag = s;
            //    item.FontWeight = FontWeights.Normal;
            //    item.Items.Add(dummyNode);
            //    item.Expanded += new RoutedEventHandler(folder_Expanded);
            //    foldersItem.Items.Add(item);
            //}
        }

        private void ResetTree(string pi_sRoot)
        {
            foldersItem.Items.Clear();
            TreeViewItem item = new TreeViewItem();
            item.Header = pi_sRoot;
            item.Tag = pi_sRoot;
            item.FontWeight = FontWeights.Normal;
            item.Items.Add(dummyNode);
            item.Expanded += new RoutedEventHandler(folder_Expanded);
            foldersItem.Items.Add(item);
            //foreach (string s in Directory.GetDirectories(pi_sRoot))//Directory.GetLogicalDrives())
            //{
            //    TreeViewItem item = new TreeViewItem();
            //    item.Header = s;
            //    item.Tag = s;
            //    item.FontWeight = FontWeights.Normal;
            //    item.Items.Add(dummyNode);
            //    item.Expanded += new RoutedEventHandler(folder_Expanded);
            //    foldersItem.Items.Add(item);
            //}
        }


        void folder_Expanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = (TreeViewItem)sender;
            if (item.Items.Count == 1 && item.Items[0] == dummyNode)
            {
                item.Items.Clear();
                try
                {
                    foreach (string s in Directory.GetDirectories(item.Tag.ToString()))
                    {
                        TreeViewItem subitem = new TreeViewItem();
                        subitem.Header = s.Substring(s.LastIndexOf("\\") + 1);
                        subitem.Tag = s;
                        subitem.FontWeight = FontWeights.Normal;
                        subitem.Items.Add(dummyNode);
                        subitem.Expanded += new RoutedEventHandler(folder_Expanded);
                        item.Items.Add(subitem);
                    }

                    foreach (string s in Directory.GetFiles(item.Tag.ToString()))
                    {
                        TreeViewItem subitem = new TreeViewItem();
                        subitem.Header = s.Substring(s.LastIndexOf("\\") + 1);
                        subitem.Tag = s;
                        subitem.FontWeight = FontWeights.Normal;
                        subitem.Items.Add(dummyNode);
                        subitem.Expanded += new RoutedEventHandler(folder_Expanded);
                        item.Items.Add(subitem);
                    }
                }
                catch (Exception) { }
            }
        }

        private void foldersItem_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeView tree = (TreeView)sender;
            TreeViewItem temp = ((TreeViewItem)tree.SelectedItem);

            if (temp == null)
                return;
            SelectedImagePath = "";
            string temp1 = "";
            string temp2 = "";
            while (true)
            {
                temp1 = temp.Header.ToString();
                if (temp1.Contains(@"\"))
                {
                    temp2 = "";
                }
                SelectedImagePath = temp1 + temp2 + SelectedImagePath;
                if (temp.Parent.GetType().Equals(typeof(TreeView)))
                {
                    break;
                }
                temp = ((TreeViewItem)temp.Parent);
                temp2 = @"\";
            }
            //show user selected path
            MessageBox.Show(SelectedImagePath);
        }

        #endregion


    }
}
