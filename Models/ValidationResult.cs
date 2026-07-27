using System;
using System.Collections.Generic;
using System.Text;

namespace Raporlama.Models
{
    public class ValidationResult
    {
        public bool IsValid { get; set; }

        public List<string> Errors { get; set; } = new();
    }
}

