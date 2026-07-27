using System;
using System.Collections.Generic;
using System.Text;

namespace Raporlama.Models
{
    public class ExcelDataRow
    {
        public int RowNumber { get; set; }

        public Dictionary<string, string?> Values { get; set; } = new();

        public List<string> Errors { get; set; } = new();

        public bool IsValid => Errors.Count == 0;
    }
}
