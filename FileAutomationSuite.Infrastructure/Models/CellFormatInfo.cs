using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Infrastructure.Models
{
    public class CellFormatInfo
    {
        public object Value { get; set; }
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }
        public ExcelVerticalAlignment VerticalAlignment { get; set; }
        public string NumberFormat { get; set; }

        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public float FontSize { get; set; }
        public string FontName { get; set; }

        public Color? FontColor { get; set; }
        public Color? BackgroundColor { get; set; }
    }
}
