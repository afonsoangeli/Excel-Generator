using System.Collections.Generic;

namespace Excel_Generator
{
    public class Excel
    {
        public Excel(string workSheetName, string headerName, List<string[]> rows)
        {
            WorkSheetName = workSheetName;
            HeaderName = headerName;
            Rows = rows;
        }

        public string WorkSheetName { get; set; }
        public string HeaderName { get; set; }
        public List<string[]> Rows { get; set; }
    }
}
