using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Media;

namespace epplus_testWPF
{
    class CellList
    {
        private const int argb = 0x000000;

        public string RowNo { get; set; }
        public string Text1 { get; set; }
        public string Text2 { get; set; }
        public SolidColorBrush bkColor { get; set;}
        public RichTextContainer RichText { get; set; }

        public CellList()
        {
            Text1 = "";
        }

        public CellList(int rowno, string item, int colno, RichTextContainer rtc)
        {
            RowNo = rowno.ToString();
            Text1 = item;
            Text2 = colno.ToString();
            RichText = rtc;
        }

        public CellList(int rowno,List<string> textlist, SolidColorBrush color, RichTextContainer rtc)
        {
            RowNo = rowno.ToString();
            if (textlist == null || textlist.Count < 2) return;
            Text1 = textlist[0];
            Text2 = textlist[1];
            bkColor = color;
            RichText = rtc;
        }

        public CellList(string rowno, List<string> textlist, SolidColorBrush color, RichTextContainer rtc)
        {
            RowNo = rowno;
            if (textlist == null || textlist.Count < 2) return;
            Text1 = textlist[0];
            Text2 = textlist[1];
            bkColor = color;
            RichText = rtc;
        }

        public CellList(string rowno)
        {
            RowNo = rowno;
        }
    }
}
