using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChartFromExcelToWord.DynamicCharts
{
    public class DynamicCharts
    {
        public class ChartTableDef
        {
            public ChartTableDef()
            {
                Columns = new List<ChartTableColumn>();
                Rows = new List<ChartTableRow>();
            }
            public int Id { get; set; }
            public string Name { get; set; }
            public string chartType { get; set; }
            public string Title { get; set; }
            public string XAxisTitle { get; set; }
            public string YAxisTitle { get; set; }
            public int startingColumnIndex { get; set; }
            public int startingRowIndex { get; set; }
            public List<ChartTableColumn> Columns { get; set; }
            public List<ChartTableRow> Rows { get; set; }
        }

        public class ChartTableColumn
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string Format { get; set; }
            public int columnIndex { get; set; }
        }

        public class ChartTableRow
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }
    }
}
