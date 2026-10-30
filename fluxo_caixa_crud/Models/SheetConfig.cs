using System.Collections.Generic;

namespace fluxo_caixa_crud.Models
{
    public enum SheetType
    {
        Tabular,
        Matrix,
        ReadOnly
    }

    public class TabularField
    {
        public string Key { get; set; }
        public string Column { get; set; }
        public string Type { get; set; } = "text";
    }

    public class SheetConfig
    {
        public string SheetName { get; set; }
        public string Entity { get; set; }
        public SheetType Type { get; set; }
        public int HeaderRow { get; set; }
        public int DataStartRow { get; set; }
        public string IdColumn { get; set; } = "A";
        public List<TabularField> Fields { get; set; } = new List<TabularField>();
        public string MonthStartCol { get; set; } = "D";
        public string MonthEndCol { get; set; } = "P";
        public int MonthCount { get; set; } = 12;
        public string RubricaCol { get; set; } = "B";
        public string SectionCol { get; set; } = "A";
        public int MonthHeaderRow { get; set; } = 7;
        public int TypeHeaderRow { get; set; } = 8;
        public int DataStartRowMatrix { get; set; } = 11;
    }
}
