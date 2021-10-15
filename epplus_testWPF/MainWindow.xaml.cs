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

            // test
            highlightTextBlock.RichText = new RichTextContainer("hoge");

            readExcel(filePath);
        }

        public void readExcel(String filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            List<CellList> Items = new List<CellList>();
            if (!fileInfo.Exists) return;
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                var workbook = package.Workbook;
                int numberOfSheets = workbook.Worksheets.Count;

                for(int SheetNo = 0; SheetNo < numberOfSheets; SheetNo++)
                {
                    var sheet = workbook.Worksheets[SheetNo];
                    for (int row = 1; row < 10; row++)
                    {
                        List<String> line = new List<string>();
                        for (int col = 1; col < 10; col++)
                        {
                            var cell = sheet.Cells[row, col];
                            if (cell.Value != null)
                            {

                                textBlock.Text += "\r\n" + cell.Value;
                                if (cell.Style.Fill.BackgroundColor.Rgb != null)
                                {
                                    textBlock.Text += cell.Style.Font.Color.Rgb;
                                    if(cell.IsRichText)
                                    {
                                        var richitext = cell.RichText[1].Color;
                                        textBlock.Text += richitext.ToString();
                                    }
                                    line.Add(cell.Value.ToString());
                                }
                            }
                        }
                        if (line.Count == 0) continue;
                        CellList cellList = new CellList(row, line, CellBkColorToBlushColor(sheet.Cells[row, 2].Style.Fill));
                        Items.Add(cellList);
                    }
                    Items.Add(new CellList());
                }
                listView.ItemsSource = Items;

            }
        }

        public SolidColorBrush CellBkColorToBlushColor(OfficeOpenXml.Style.ExcelFill excelFill)
        {
            // no color
            if (excelFill.BackgroundColor.Rgb == null) return new SolidColorBrush(Color.FromRgb(0,0,0));

            // Theme は　固定色 green
            if(excelFill.BackgroundColor.Theme != "") return new SolidColorBrush(Color.FromRgb(100, 150, 200));

            // 色付き
            string rgb = excelFill.BackgroundColor.Rgb;
            byte r = Convert.ToByte(rgb.Substring(2, 2), 16);
            byte g = Convert.ToByte(rgb.Substring(4, 2), 16);
            byte b = Convert.ToByte(rgb.Substring(6, 2), 16);

            return new SolidColorBrush(Color.FromRgb(r, g, b));
        }

        //https://social.msdn.microsoft.com/Forums/aspnet/ja-JP/8dbcf3bd-5ceb-4eaa-84f4-3def7017e35c/listview2086912398text123981996837096124342580520316123771242726041?forum=wpfja
    }
}
