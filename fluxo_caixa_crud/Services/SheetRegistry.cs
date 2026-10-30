using System;
using System.Collections.Generic;
using System.Linq;
using fluxo_caixa_crud.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    public class SheetRegistry
    {
        private const string ConfigSheetName = "_CONFIG";
        private readonly Dictionary<string, SheetConfig> _byEntity = new Dictionary<string, SheetConfig>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, SheetConfig> _bySheet = new Dictionary<string, SheetConfig>(StringComparer.OrdinalIgnoreCase);
        private readonly List<SheetConfig> _all = new List<SheetConfig>();
        private LookupService _lookups;

        public IReadOnlyList<SheetConfig> All => _all;
        public LookupService Lookups => _lookups ?? (_lookups = new LookupService());

        public void Initialize()
        {
            _all.Clear();
            _byEntity.Clear();
            _bySheet.Clear();
            _all.AddRange(GetDefaultConfigs());
            EnsureConfigSheet();
            LoadFromConfigSheet();
            foreach (var c in _all)
            {
                if (!string.IsNullOrEmpty(c.Entity))
                    _byEntity[c.Entity] = c;
                _bySheet[c.SheetName] = c;
            }
            _lookups = new LookupService();
            _lookups.Load(this);
        }

        public SheetConfig GetByEntity(string entity)
        {
            if (string.IsNullOrEmpty(entity) || !_byEntity.TryGetValue(entity, out var cfg))
                throw new InvalidOperationException("Entidade desconhecida: " + entity);
            return cfg;
        }

        public SheetConfig GetBySheet(string sheetName)
        {
            if (!_bySheet.TryGetValue(sheetName, out var cfg))
                throw new InvalidOperationException("Aba não configurada: " + sheetName);
            return cfg;
        }

        public IEnumerable<SheetConfig> MatrixSheets => _all.Where(c => c.Type == SheetType.Matrix);

        public object GetMeta()
        {
            var lists = Lookups.ToMetaObject();
            var unitLabels = UnitDisplayService.ReadAllUnitLabels(this);
            return new
            {
                entities = _byEntity.Keys.OrderBy(x => x).ToList(),
                tabular = _all.Where(c => c.Type == SheetType.Tabular).Select(c => new { c.Entity, c.SheetName }).ToList(),
                matrix = MatrixSheets.Select(c =>
                {
                    string sheet = c.SheetName.Trim();
                    UnitDisplayInfo info = null;
                    if (UnitDisplayService.IsUnitMatrix(c.Entity))
                        unitLabels.TryGetValue(sheet, out info);
                    return new
                    {
                        sheet = c.SheetName,
                        entity = c.Entity,
                        label = info?.Label ?? sheet,
                        code = info?.Code ?? sheet,
                        nome = info?.Nome ?? sheet
                    };
                }).ToList(),
                unitLabels = unitLabels.Values
                    .OrderBy(x => x.Label, StringComparer.OrdinalIgnoreCase)
                    .Select(x => new { x.Sheet, x.Code, x.Nome, x.Label })
                    .ToList(),
                readonlySheets = _all.Where(c => c.Type == SheetType.ReadOnly).Select(c => c.SheetName).ToList(),
                meses = new[] { "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO" },
                tipos = new[] { "Realizado", "Previsto" },
                naturezas = new[] { "Receita", "Despesa", "Resultado", "Outros" },
                monthEndCol = "P",
                unidades = Lookups.Unidades,
                origensCapex = Lookups.OrigensCapex,
                centrosCusto = Lookups.CentrosCusto,
                classificacoesCapex = Lookups.ClassificacoesCapex,
                responsaveis = Lookups.Responsaveis,
                simNao = new[] { "Sim", "Não" }
            };
        }

        private void EnsureConfigSheet()
        {
            ExcelHelper.RunWithExcel(() =>
            {
                Excel.Worksheet ws;
                try
                {
                    ws = ExcelHelper.GetSheet(ConfigSheetName);
                    if (ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", 1)) == "SheetName") return;
                }
                catch
                {
                    ws = ExcelHelper.GetOrCreateSheet(ConfigSheetName);
                }
                ws.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                if (ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", 1)) == "SheetName") return;

                ws.Cells.Clear();
                string[] headers = { "SheetName", "Type", "Entity", "HeaderRow", "DataStartRow", "IdColumn", "MonthStartCol", "MonthCount", "RubricaCol", "MonthHeaderRow", "TypeHeaderRow", "DataStartRowMatrix" };
                for (int i = 0; i < headers.Length; i++)
                    ExcelHelper.SetCell(ws, ExcelHelper.ColLetter(i + 1), 1, headers[i]);

                int row = 2;
                foreach (var c in _all)
                {
                    ExcelHelper.SetCell(ws, "A", row, c.SheetName);
                    ExcelHelper.SetCell(ws, "B", row, c.Type.ToString());
                    ExcelHelper.SetCell(ws, "C", row, c.Entity ?? "");
                    ExcelHelper.SetCell(ws, "D", row, c.HeaderRow);
                    ExcelHelper.SetCell(ws, "E", row, c.DataStartRow);
                    ExcelHelper.SetCell(ws, "F", row, c.IdColumn ?? "");
                    ExcelHelper.SetCell(ws, "G", row, c.MonthStartCol ?? "");
                    ExcelHelper.SetCell(ws, "H", row, c.MonthCount);
                    ExcelHelper.SetCell(ws, "I", row, c.RubricaCol ?? "");
                    ExcelHelper.SetCell(ws, "J", row, c.MonthHeaderRow);
                    ExcelHelper.SetCell(ws, "K", row, c.TypeHeaderRow);
                    ExcelHelper.SetCell(ws, "L", row, c.DataStartRowMatrix);
                    row++;
                }
            });
        }

        private void LoadFromConfigSheet()
        {
            try
            {
                var ws = ExcelHelper.GetSheet(ConfigSheetName);
                int row = 2;
                while (true)
                {
                    string name = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", row));
                    if (string.IsNullOrEmpty(name)) break;
                    var existing = _all.FirstOrDefault(x => x.SheetName == name);
                    if (existing != null)
                    {
                        string typeStr = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "B", row));
                        if (Enum.TryParse(typeStr, true, out SheetType st)) existing.Type = st;
                        existing.Entity = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "C", row));
                        existing.HeaderRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "D", row));
                        existing.DataStartRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "E", row));
                        existing.IdColumn = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "F", row));
                        existing.MonthStartCol = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "G", row));
                        existing.MonthCount = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "H", row));
                        existing.RubricaCol = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "I", row));
                        existing.MonthHeaderRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "J", row));
                        existing.TypeHeaderRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "K", row));
                        existing.DataStartRowMatrix = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "L", row));
                    }
                    else
                    {
                        var cfg = ParseConfigRow(ws, row);
                        if (cfg != null) _all.Add(cfg);
                    }
                    row++;
                }
            }
            catch { /* usa defaults */ }
        }

        private static SheetConfig ParseConfigRow(Excel.Worksheet ws, int row)
        {
            string name = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", row));
            if (string.IsNullOrEmpty(name)) return null;
            string typeStr = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "B", row));
            if (!Enum.TryParse(typeStr, true, out SheetType st)) return null;
            return new SheetConfig
            {
                SheetName = name,
                Entity = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "C", row)),
                Type = st,
                HeaderRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "D", row)),
                DataStartRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "E", row)),
                IdColumn = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "F", row)),
                MonthStartCol = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "G", row)) ?? "D",
                MonthCount = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "H", row)),
                RubricaCol = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "I", row)) ?? "B",
                MonthHeaderRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "J", row)),
                TypeHeaderRow = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "K", row)),
                DataStartRowMatrix = (int)ExcelHelper.CellNum(ExcelHelper.GetCell(ws, "L", row)),
                MonthEndCol = "P",
                SectionCol = "A"
            };
        }

        public static string EntityFromSheetName(string sheetName)
        {
            return "matrix_" + sheetName
                .Replace(" ", "_")
                .Replace("(", "")
                .Replace(")", "")
                .Replace("%", "");
        }

        private static List<SheetConfig> GetDefaultConfigs()
        {
            var list = new List<SheetConfig>();

            list.Add(Tabular("capex", "CAPEX (2)", 2, 3, new[]
            {
                T("id","A"), T("unidade","B"), T("fluxo","C"), T("pctTerral","D","percent"),
                T("nomeProjeto","E"), T("valor","F","money"), T("pctTerralCalc","G","percent"),
                T("origemCapex","H"), T("observacao","I"), T("obsQueiroga","J"),
                T("centroCusto","K"), T("projeto","L"), T("classificacao","M")
            }));

            list.Add(Tabular("novos_neg", "Novos Neg.", 3, 4, new[]
            {
                T("id","A"), T("unidade","B"), T("codigo","C"), T("descricao","D"),
                T("valor","E","money"), T("payback","F")
            }));

            list.Add(Tabular("premissas_anot", "ANOTAÇÕES DE PREMISSAS", 2, 3, new[]
            {
                T("id","Z"), T("nota","A"), T("unidade","B"), T("criterio","C"), T("responsavel","D")
            }));

            list.Add(Tabular("bonus", "bonus_caixa_2020", 3, 4, new[]
            {
                T("id","A"), T("undPertencente","B"), T("filialPagante","C"), T("centroCusto","D"),
                T("projeto","E"), T("valorProvisao","F","money"), T("marco","G"), T("abril","H"), T("maio","I")
            }));

            string[] matrixUnits = {
                "ALS","BPS","BRV","BUR","VSPAR","GUA","MOG","POR","SHS","MPM","TSC",
                "ALS 50","MOG 50","SPE Rio Verde (88%)","Sociedade Rio Verde (12%)",
                "SPE 200 (51,33%)","SPE 201- Ribeirão Preto","SPE 204 - Palmas ","SPE 207",
                "SPE 208","RIO BRAVO","SPE 800 ","SNJ","PTS","VSPAR "
            };
            foreach (var u in matrixUnits)
                list.Add(Matrix(EntityFromSheetName(u), u));

            list.Add(Matrix("matrix_premissas_cal", "premissas calibradas", 5, 6, 6, "C", 11));
            list.Add(Matrix("matrix_premissas_neg", "PREMISSAS DE NEGÓCIO", 5, 6, 6, "C", 11));
            list.Add(Matrix("matrix_capex0403", "CAPEX_0403", 2, 3, 3, "C", 12));

            string[] readOnly = {
                "CAPA.","Meta NOI","Receita Financeira","Planilha1","PrevxReal","Planilha2",
                "CAPEX","2026 SHO","DRE Desenv."
            };
            foreach (var s in readOnly.Distinct())
                if (!list.Any(x => x.SheetName == s))
                    list.Add(new SheetConfig { SheetName = s, Entity = "readonly_" + s, Type = SheetType.ReadOnly });

            return list;
        }

        private static TabularField T(string key, string col, string type = "text")
        {
            return new TabularField { Key = key, Column = col, Type = type };
        }

        private static SheetConfig Tabular(string entity, string sheet, int header, int dataStart, TabularField[] fields)
        {
            return new SheetConfig
            {
                Entity = entity, SheetName = sheet, Type = SheetType.Tabular,
                HeaderRow = header, DataStartRow = dataStart, IdColumn = "A", Fields = fields.ToList()
            };
        }

        private static SheetConfig Matrix(string entity, string sheet, int monthRow = 7, int typeRow = 8, int dataStart = 11, string rubricaCol = "B", int monthCount = 12)
        {
            return new SheetConfig
            {
                Entity = entity, SheetName = sheet, Type = SheetType.Matrix,
                MonthHeaderRow = monthRow, TypeHeaderRow = typeRow, DataStartRowMatrix = dataStart,
                RubricaCol = rubricaCol, MonthStartCol = "D", MonthEndCol = "P",
                SectionCol = "A", MonthCount = monthCount
            };
        }
    }
}
