using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.IO;
using OfficeOpenXml;
using System.Windows.Controls;

namespace epplus_testWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       
        public MainWindow()
        {
            InitializeComponent();
            String filePath = "..\\..\\..\\test.xlsx";

            //readExcel(filePath);
            ExcelColorList ecl = new ExcelColorList(filePath);
            listView.ItemsSource = ecl.getCellList();
        }
    }
}
