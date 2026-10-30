using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    public class LookupService
    {
        private List<string> _unidades = new List<string>();
        private List<string> _origensCapex = new List<string>();
        private List<string> _centrosCusto = new List<string>();
        private List<string> _classificacoesCapex = new List<string>();
        private List<string> _responsaveis = new List<string>();

        public IReadOnlyList<string> Unidades => _unidades;
        public IReadOnlyList<string> OrigensCapex => _origensCapex;
        public IReadOnlyList<string> CentrosCusto => _centrosCusto;
        public IReadOnlyList<string> ClassificacoesCapex => _classificacoesCapex;
        public IReadOnlyList<string> Responsaveis => _responsaveis;

        public void Load(SheetRegistry registry)
        {
            var unidades = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var origens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var centros = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var classif = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var resp = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var m in registry.MatrixSheets)
            {
                if (m.Entity != null && m.Entity.StartsWith("matrix_") &&
                    !m.Entity.Contains("premissas") && !m.Entity.Contains("capex0403"))
                {
                    unidades.Add(m.SheetName.Trim());
                }
            }

            CollectColumn("CAPEX (2)", "B", 3, unidades);
            CollectColumn("CAPEX (2)", "H", 3, origens);
            CollectColumn("CAPEX (2)", "K", 3, centros);
            CollectColumn("CAPEX (2)", "M", 3, classif);
            CollectColumn("Novos Neg.", "B", 4, unidades);
            CollectColumn("ANOTAÇÕES DE PREMISSAS", "B", 3, unidades);
            CollectColumn("ANOTAÇÕES DE PREMISSAS", "D", 3, resp);
            CollectColumn("bonus_caixa_2020", "B", 4, unidades);

            foreach (var r in ConfigService.ReadLookupValues("responsavel"))
                resp.Add(r);

            foreach (var c in new[] {
                "(-) Capex Manutenção", "(-) Capex Benfeitorias", "(-) Capex Comercial",
                "(-) Novos Negócios", "(-) Capex Melhoria de Ativos", "(-) Investimentos Expansão"
            }) classif.Add(c);

            foreach (var o in new[] {
                "CAPEX DAY - EMPREENDEDOR", "CAPEX DAY - COMERCIAL", "CAPEX DAY - EXTRA",
                "DESENVOLVIMENTO", "EPA REF. DEZ/19"
            }) origens.Add(o);

            _unidades = unidades.Where(x => x.Length <= 40 && !x.Contains("Enviar para")).OrderBy(x => x).ToList();
            _origensCapex = origens.Where(x => !string.IsNullOrWhiteSpace(x)).OrderBy(x => x).ToList();
            _centrosCusto = centros.Where(x => !string.IsNullOrWhiteSpace(x)).OrderBy(x => x).ToList();
            _classificacoesCapex = classif.Where(x => !string.IsNullOrWhiteSpace(x)).OrderBy(x => x).ToList();
            _responsaveis = resp.Where(x => !string.IsNullOrWhiteSpace(x)).OrderBy(x => x).ToList();
        }

        public object ToMetaObject()
        {
            return new
            {
                unidades = _unidades,
                origensCapex = _origensCapex,
                centrosCusto = _centrosCusto,
                classificacoesCapex = _classificacoesCapex,
                responsaveis = _responsaveis,
                simNao = new[] { "Sim", "Não" }
            };
        }

        private static void CollectColumn(string sheetName, string col, int startRow, HashSet<string> target)
        {
            try
            {
                var ws = ExcelHelper.GetSheet(sheetName);
                int last = ExcelHelper.LastRow(ws, col, startRow);
                for (int r = startRow; r <= last; r++)
                {
                    string v = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, col, r));
                    if (!string.IsNullOrWhiteSpace(v) && v.Length <= 80)
                        target.Add(v.Trim());
                }
            }
            catch { }
        }
    }
}
