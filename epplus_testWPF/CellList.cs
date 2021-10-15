using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Media;

namespace epplus_testWPF
{
    class CellList
    {
        private const int argb = 0x000000;

        public int RowNo { get; set; }
        public string Text1 { get; set; }
        public string Text2 { get; set; }
        public SolidColorBrush bkColor { get; set;}

        public CellList()
        {
            Text1 = "";
        }
        public CellList(int rowno,List<string> textlist, SolidColorBrush color)
        {
            RowNo = rowno;
            if (textlist == null || textlist.Count < 2) return;
            Text1 = textlist[0];
            Text2 = textlist[1];
            bkColor = color;
        }
    }
}
