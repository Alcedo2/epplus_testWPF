using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Media;

namespace epplus_testWPF
{
    class ExcelColorList
    {
        public int MaxRow = 100;
        public int MaxCol = 100;
        public int ItemColNo = 3;

        private string filePath;
        public ExcelColorList(string fileName)
        {
            filePath = fileName;
        }

        public List<CellList> getCellList()
        {
            FileInfo fileInfo = new FileInfo(filePath);
            List<CellList> Items = new List<CellList>();

            if (!fileInfo.Exists) return Items;
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                var workbook = package.Workbook;
                int numberOfSheets = workbook.Worksheets.Count;

                // 全シート対象
                for (int SheetNo = 0; SheetNo < numberOfSheets; SheetNo++)
                {
                    var sheet = workbook.Worksheets[SheetNo];
                    // シート名
                    Items.Add(new CellList(sheet.Name));

                    // 最大行まで走査
                    for (int row = 1; row < MaxRow; row++)
                    {
                        List<String> line = new List<string>();

                        // 最大列まで走査
                        for (int col = 1; col < MaxCol; col++)
                        {
                            var ItemCell = sheet.Cells[row, ItemColNo];
                            string ItemName;
                            if (ItemCell.Value == null)
                                ItemName = "no Item Name";
                            else
                                ItemName = ItemCell.Value.ToString();

                            var cell = sheet.Cells[row, col];

                            RichTextContainer rich = new RichTextContainer();
                            if (cell.Value != null)
                            {

                                var backGroundColor = cell.Style.Fill.BackgroundColor;
                                Brush bkBrush = null;

                                // BackColor
                                if(backGroundColor.Rgb != null && backGroundColor.Rgb != "")
                                {
                                    bkBrush = ExcelColorToBrush(backGroundColor);
                                }
                                // テーマは固定
                                if (backGroundColor.Theme != null && backGroundColor.Theme != "")
                                {
                                    bkBrush = Brushes.LightGoldenrodYellow;
                                }

                                // TextColor
                                // RichText
                                if (cell.IsRichText)
                                {
                                    foreach (OfficeOpenXml.Style.ExcelRichText richText in cell.RichText)
                                    {
                                        rich.add(richText.Text, ExcelColorToBrush(richText.Color),bkBrush);
                                    }
                                }
                                // 全部色
                                else if (cell.Style.Font.Color.Rgb != null && cell.Style.Font.Color.Rgb != "")
                                {
                                    rich.add(cell.Value.ToString(), ExcelColorToBrush(cell.Style.Font.Color), bkBrush);
                                }
                                else if (cell.Style.Font.Color.Theme != null && cell.Style.Font.Color.Theme != "1") // ここから修正する
                                {
                                    // テーマは固定色
                                    rich.add(cell.Value.ToString(), Brushes.IndianRed, bkBrush);
                                }
                                else
                                    continue;
                                Items.Add(new CellList(row, ItemName, col, rich));
                            }
                        }

                    }

                }
            }
            return Items;
        }

        public Brush ExcelColorToBrush(System.Drawing.Color color)
        {
            if (color == null) return Brushes.Black;

            var MColor = Color.FromRgb(color.R, color.G, color.B);
            var brush = new SolidColorBrush(MColor);
            return brush;
        }

        public Brush ExcelColorToBrush(OfficeOpenXml.Style.ExcelColor color)
        {
            if (color == null) return Brushes.Black;

            byte[] argb = ColorStringToByte(color.Rgb);
            var MColor = Color.FromRgb(argb[1],argb[2],argb[3]);
            var brush = new SolidColorBrush(MColor);
            return brush;
        }

        // colorString format is #FF123456
        public byte[] ColorStringToByte(String colorString)
        {
            byte[] argb = new byte[4];
            try
            {
                for(int i = 0; i < 4; i++)
                    argb[i] = Convert.ToByte(colorString.Substring(i*2, 2),16);

            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            

            return argb;
        }
    }



}
