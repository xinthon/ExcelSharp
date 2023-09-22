using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp.Sample
{
    public static class FileDialog
    {
        public static OpenFileDialog Open { get; } = new()
        {
            Filter = "Excel Files|*.xlsx;*.xls"
        };

        public static SaveFileDialog Save { get; } = new()
        {
            Filter = "Excel Files|*.xlsx;*.xls"
        };
    }
}
