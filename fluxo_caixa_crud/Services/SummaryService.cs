using System;
using System.Collections.Generic;
using System.Linq;
using fluxo_caixa_crud.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    public class SummaryService
    {
        private readonly SheetRegistry _registry;
        private readonly TabularCrudService _tabular;

        public SummaryService(SheetRegistry registry, TabularCrudService tabular)
        {
            _registry = registry;
            _tabular = tabular;
        }

        public object GetResumo()
        {
            var alerts = new List<object>();
            var cards = new List<object>();

            string[] checkSheets = { "Meta NOI", "Receita Financeira", "CAPA.", "PrevxReal" };
            foreach (var sheetName in checkSheets)
            {
                try
                {
                    var ws = ExcelHelper.GetSheet(sheetName);
                    var used = ws.UsedRange;
                    if (used == null) continue;
                    int maxCheck = Math.Min(used.Rows.Count, 30);
                    int maxCol = Math.Min(used.Columns.Count, 15);
                    for (int r = 1; r <= maxCheck; r++)
                    for (int c = 1; c <= maxCol; c++)
                    {
                        var val = ExcelHelper.CellStr(((Excel.Range)used.Cells[r, c]).Value2);
                        if (val == "#REF!")
                        {
                            alerts.Add(new { sheet = sheetName, cell = ((Excel.Range)used.Cells[r, c]).Address });
                            break;
                        }
                    }
                }
                catch { }
            }

            long totalCapex = 0;
            try
            {
                var capex = _tabular.List("capex");
                totalCapex = capex.Sum(x => x.ContainsKey("valor") ? Convert.ToInt64(x["valor"]) : 0);
                cards.Add(new { titulo = "Total CAPEX", valorCents = totalCapex, icone = "capex" });
                cards.Add(new { titulo = "Itens CAPEX", valor = capex.Count, icone = "count" });
            }
            catch { }

            var matrixUnits = GetMatrixUnitNames();

            var unitLabels = UnitDisplayService.ReadAllUnitLabels(_registry);
            var prevReal = BuildPrevRealByUnit(matrixUnits, unitLabels);

            return new
            {
                cards,
                alerts = alerts.Take(20).ToList(),
                alertCount = alerts.Count,
                unidades = matrixUnits,
                prevReal
            };
        }

        private List<string> GetMatrixUnitNames()
        {
            return _registry.MatrixSheets
                .Where(c => c.Entity != null && c.Entity.StartsWith("matrix_")
                    && !c.Entity.Contains("premissas") && !c.Entity.Contains("capex0403"))
                .Select(c => c.SheetName.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<object> BuildPrevRealByUnit(List<string> matrixUnits, Dictionary<string, UnitDisplayInfo> unitLabels)
        {
            var fromSheet = new Dictionary<string, Tuple<long, long>>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var ws = ExcelHelper.GetSheet("PrevxReal");
                int last = ExcelHelper.LastRow(ws, "B", 5);
                for (int r = 5; r <= last; r++)
                {
                    string unidade = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "B", r));
                    if (string.IsNullOrWhiteSpace(unidade)) continue;
                    fromSheet[unidade.Trim()] = Tuple.Create(
                        ExcelHelper.ToCents(ExcelHelper.GetCell(ws, "C", r)),
                        ExcelHelper.ToCents(ExcelHelper.GetCell(ws, "D", r)));
                }
            }
            catch { }

            var list = new List<object>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var u in matrixUnits)
            {
                seen.Add(u);
                long prev = 0, real = 0;
                Tuple<long, long> vals;
                if (fromSheet.TryGetValue(u, out vals))
                {
                    prev = vals.Item1;
                    real = vals.Item2;
                }
                list.Add(new
                {
                    unidade = u,
                    label = unitLabels.ContainsKey(u) ? unitLabels[u].Label : u,
                    previstoCents = prev,
                    realizadoCents = real
                });
            }

            foreach (var kv in fromSheet.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                if (seen.Contains(kv.Key)) continue;
                UnitDisplayInfo info;
                unitLabels.TryGetValue(kv.Key, out info);
                list.Add(new
                {
                    unidade = kv.Key,
                    label = info != null ? info.Label : kv.Key,
                    previstoCents = kv.Value.Item1,
                    realizadoCents = kv.Value.Item2
                });
            }

            return list;
        }
    }
}
