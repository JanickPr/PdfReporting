using System;
using System.Collections.Generic;
using System.Windows;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    public class PageNumberSettings
    {
        public Point Position
        {
            get; set;
        }

        public bool OverallNumeration
        {
            get; set;
        }

        public Typeface FontFamily
        {
            get; set;
        }

        public int FontSize
        {
            get; set;
        }

        public Brush FontBrush
        {
            get; set;
        }

        public string PagePrefix
        {
            get; set;
        }

        public bool PageOfAllPagesNotation
        {
            get; set;
        }

        public PageNumberSettings(Point position, Typeface fontFamily, int fontSize, Brush fontBrush, bool overallNumeration, string pagePrefix, bool pageOfAllPagesNotation)
        {
            this.Position = position;
            this.FontFamily = fontFamily;
            this.FontSize = fontSize;
            this.FontBrush = fontBrush;
            this.OverallNumeration = overallNumeration;
            this.PagePrefix = pagePrefix;
            this.PageOfAllPagesNotation = pageOfAllPagesNotation;
        }
    }
}
