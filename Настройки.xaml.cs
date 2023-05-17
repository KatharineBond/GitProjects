using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SWIFT_СПФС
{
    /// <summary>
    /// Логика взаимодействия для Настройки.xaml
    /// </summary>
    public partial class Настройки : Window
    {
        public Настройки()
        {
            InitializeComponent();
            CodeBnk.Text = GlobalTrash.CodeBnk;
            SWIFTBnk.Text = GlobalTrash.SWIFTBnk;
            NameBnk.Text = GlobalTrash.NameBnk;
            Mist.Text = GlobalTrash.Mist;
            PathABSOut.Text = GlobalTrash.PathABSOut;
            PathABSIn.Text = GlobalTrash.PathABSIn;
            FileMaskOut.Text = GlobalTrash.FileMaskOut;
            PathTransit.Text = GlobalTrash.PathTransit;
            PathOut.Text = GlobalTrash.PathOut;
            PathIn.Text = GlobalTrash.PathIn;
            FileMaskIn.Text = GlobalTrash.FileMaskIn;
            PathArc.Text = GlobalTrash.PathArc;
            SPFSPathIn.Text = GlobalTrash.SPFSPathIn;
            FileMask.Text = GlobalTrash.FileMask;
            SPFSPathOut.Text = GlobalTrash.SPFSPathOut;
            SPFSPathTransit.Text = GlobalTrash.SPFSPathTransit;
            SPFSPathArc.Text = GlobalTrash.SPFSPathArc;
        }

        private void Accept_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        public string My_CodeBnk
        {
            get { return CodeBnk.Text; }
        }
        public string My_SWIFTBnk
        {
            get { return SWIFTBnk.Text; }
        }
        public string My_NameBnk
        {
            get { return NameBnk.Text; }
        }
        public string My_Mist
        {
            get { return Mist.Text; }
        }
        public string My_PathABSOut
        {
            get { return PathABSOut.Text; }
        }
        public string My_PathABSIn
        {
            get { return PathABSIn.Text; }
        }
        public string My_FileMaskOut
        {
            get { return FileMaskOut.Text; }
        }
        public string My_PathTransit
        {
            get { return PathTransit.Text; }
        }
        public string My_PathOut
        {
            get { return PathOut.Text; }
        }
        public string My_PathIn
        {
            get { return PathIn.Text; }
        }
        public string My_FileMaskIn
        {
            get { return FileMaskIn.Text; }
        }
        public string My_PathArc
        {
            get { return PathArc.Text; }
        }
        public string My_SPFSPathIn
        {
            get { return SPFSPathIn.Text; }
        }
        public string My_FileMask
        {
            get { return FileMask.Text; }
        }
        public string My_SPFSPathOut
        {
            get { return SPFSPathOut.Text; }
        }
        public string My_SPFSPathTransit
        {
            get { return SPFSPathTransit.Text; }
        }
        public string My_SPFSPathArc
        {
            get { return SPFSPathArc.Text; }
        }
    }
}