using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Infrastructure.Models
{
    internal class ExcelReadModel
    {
    }

    public class ExcelCellData
    {
        public string Header { get; set; }
        public object Value { get; set; }
        public string NumberFormat { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public string FontColor { get; set; }
        public string BackgroundColor { get; set; }
    }

    public class ExcelHeaderInfo
    {
        public string HeaderName { get; set; }
        public string Value { get; set; }
        public string FontColor { get; set; }
        public string BackgroundColor { get; set; }
        public float FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
    }

}
