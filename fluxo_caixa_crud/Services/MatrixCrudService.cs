using System;
using System.Collections.Generic;
using System.Linq;
using fluxo_caixa_crud.Models;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    public class MatrixCrudService
    {
        private readonly SheetRegistry _registry;
        private static readonly string[] Meses = {
            "JANEIRO","FEVEREIRO","MARÇO","MARCO","ABRIL","MAIO","JUNHO",
            "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"
        };

        public MatrixCrudService(SheetRegistry registry)
        {
            _registry = registry;
        }

        public List<Dictionary<string, object>> GetRubricas(string sheetName)
        {
            var cfg = _registry.GetBySheet(sheetName);
            var ws = ExcelHelper.GetSheet(sheetName);
            var list = new List<Dictionary<string, object>>();
            int last = ExcelHelper.LastRow(ws, cfg.RubricaCol, cfg.DataStartRowMatrix);
            for (int r = cfg.DataStartRowMatrix; r <= last; r++)
            {
                string rub = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, cfg.RubricaCol, r));
                if (string.IsNullOrWhiteSpace(rub)) continue;
                string section = GetSectionAtRow(ws, cfg, r);
                string natureza = ClassifyNatureza(rub, section);
                list.Add(new Dictionary<string, object>
                {
                    ["nome"] = rub,
                    ["natureza"] = natureza,
                    ["secao"] = section,
                    ["row"] = r
                });
            }
            return list
                .GroupBy(x => x["nome"].ToString(), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        public List<Dictionary<string, object>> List(string sheetName, string rubricaFilter = null)
        {
            var cfg = _registry.GetBySheet(sheetName);
            var ws = ExcelHelper.GetSheet(sheetName);
            var result = new List<Dictionary<string, object>>();
            var rubricas = GetRubricas(sheetName);
            if (!string.IsNullOrEmpty(rubricaFilter))
                rubricas = rubricas.Where(r => string.Equals(r["nome"].ToString().Trim(), rubricaFilter.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();

            var columns = GetMatrixColumns(ws, cfg);

            foreach (var rubEntry in rubricas)
            {
                string rub = rubEntry["nome"].ToString();
                string natureza = rubEntry["natureza"].ToString();
                string secao = rubEntry.ContainsKey("secao") ? rubEntry["secao"]?.ToString() ?? "" : "";
                int row = FindRubricaRow(ws, cfg, rub);
                if (row < 0) continue;
                foreach (var colInfo in columns)
                {
                    object val = ExcelHelper.GetCell(ws, colInfo.Col, row);
                    long cents = ExcelHelper.ToCents(val);
                    result.Add(new Dictionary<string, object>
                    {
                        ["id"] = sheetName + "|" + rub + "|" + colInfo.Col + "|" + colInfo.Tipo,
                        ["sheet"] = sheetName,
                        ["rubrica"] = rub,
                        ["natureza"] = natureza,
                        ["secao"] = secao,
                        ["mes"] = colInfo.Mes,
                        ["coluna"] = colInfo.Col,
                        ["tipo"] = colInfo.Tipo,
                        ["valorCents"] = cents,
                        ["valor"] = ExcelHelper.CellNum(val),
                        ["row"] = row,
                        ["readonly"] = colInfo.ReadOnly
                    });
                }
            }
            return result;
        }

        public Dictionary<string, object> Get(string sheetName, string rubrica, string mes, string tipo)
        {
            var cfg = _registry.GetBySheet(sheetName);
            var ws = ExcelHelper.GetSheet(sheetName);
            int row = FindRubricaRow(ws, cfg, rubrica);
            if (row < 0) throw new InvalidOperationException("Rubrica não encontrada.");
            string col = ResolveColumn(ws, cfg, mes, tipo);
            return new Dictionary<string, object>
            {
                ["sheet"] = sheetName, ["rubrica"] = rubrica, ["mes"] = mes, ["tipo"] = tipo,
                ["coluna"] = col, ["valorCents"] = ExcelHelper.ToCents(ExcelHelper.GetCell(ws, col, row)), ["row"] = row
            };
        }

        public Dictionary<string, object> Upsert(JObject payload)
        {
            string sheet = payload["sheet"]?.ToString();
            string rubrica = payload["rubrica"]?.ToString();
            string mes = payload["mes"]?.ToString();
            string tipo = payload["tipo"]?.ToString() ?? "Realizado";
            string coluna = payload["coluna"]?.ToString()?.Trim();
            string colunaOrigem = payload["colunaOrigem"]?.ToString()?.Trim();
            long cents = payload["valorCents"]?.Value<long>() ?? 0;

            if (string.IsNullOrEmpty(sheet) || string.IsNullOrEmpty(rubrica))
                throw new InvalidOperationException("sheet e rubrica são obrigatórios.");

            Dictionary<string, object> result = null;
            ExcelHelper.RunWithExcel(() =>
            {
                var cfg = _registry.GetBySheet(sheet);
                var ws = ExcelHelper.GetSheet(sheet);
                int row = FindRubricaRow(ws, cfg, rubrica);
                if (row < 0) throw new InvalidOperationException("Rubrica não encontrada na aba.");
                string col = ResolveColumn(ws, cfg, mes, tipo, coluna);
                if (IsReadOnlyColumn(ws, cfg, col))
                    throw new InvalidOperationException("A coluna TOTAL não pode ser editada manualmente.");

                var colMeta = GetMatrixColumns(ws, cfg).FirstOrDefault(c =>
                    c.Col.Equals(col, StringComparison.OrdinalIgnoreCase));
                string headerTipo = colMeta != null ? colMeta.Tipo : NormalizeTipo(tipo);
                string requestedTipo = NormalizeTipo(tipo);

                if (!string.IsNullOrWhiteSpace(colunaOrigem)
                    && !colunaOrigem.Equals(col, StringComparison.OrdinalIgnoreCase))
                {
                    string origCol = ResolveColumnHint(ws, cfg, colunaOrigem);
                    if (!IsReadOnlyColumn(ws, cfg, origCol))
                        ExcelHelper.SetCell(ws, origCol, row, ExcelHelper.FromCents(0));
                }

                if (!TipoMatches(headerTipo, requestedTipo))
                    SetColumnTipoHeader(ws, cfg, col, requestedTipo);

                ExcelHelper.SetCell(ws, col, row, ExcelHelper.FromCents(cents));
                result = new Dictionary<string, object>
                {
                    ["sheet"] = sheet,
                    ["rubrica"] = rubrica,
                    ["mes"] = colMeta != null ? colMeta.Mes : mes,
                    ["tipo"] = requestedTipo,
                    ["coluna"] = col,
                    ["tipoColunaAtualizado"] = !TipoMatches(headerTipo, requestedTipo),
                    ["valorCents"] = cents,
                    ["valor"] = ExcelHelper.CellNum(ExcelHelper.FromCents(cents)),
                    ["row"] = row
                };
            });
            return result;
        }

        public void Delete(JObject payload)
        {
            payload["valorCents"] = 0;
            Upsert(payload);
        }

        public List<Dictionary<string, object>> GetColumnLayout(string sheetName)
        {
            var cfg = _registry.GetBySheet(sheetName);
            var ws = ExcelHelper.GetSheet(sheetName);
            return GetMatrixColumns(ws, cfg)
                .Where(c => !c.ReadOnly)
                .Select(c => new Dictionary<string, object>
                {
                    ["coluna"] = c.Col,
                    ["mes"] = c.Mes,
                    ["tipo"] = c.Tipo
                })
                .ToList();
        }

        public bool HasLaunchValues(string sheetName)
        {
            var cfg = _registry.GetBySheet(sheetName);
            var ws = ExcelHelper.GetSheet(sheetName);
            int last = ExcelHelper.LastRow(ws, cfg.RubricaCol, cfg.DataStartRowMatrix);
            int start = ExcelHelper.ColIndex(cfg.MonthStartCol);
            int end = ExcelHelper.ColIndex(string.IsNullOrEmpty(cfg.MonthEndCol) ? "P" : cfg.MonthEndCol);
            for (int r = cfg.DataStartRowMatrix; r <= last; r++)
            {
                for (int idx = start; idx <= end; idx++)
                {
                    string col = ExcelHelper.ColLetter(idx);
                    if (IsReadOnlyColumn(ws, cfg, col)) continue;
                    if (Math.Abs(ExcelHelper.ToCents(ExcelHelper.GetCell(ws, col, r))) > 0)
                        return true;
                }
            }
            return false;
        }

        public Dictionary<string, object> AddRubrica(string sheetName, string nome, string secao = null)
        {
            nome = (nome ?? "").Trim();
            if (string.IsNullOrEmpty(nome)) throw new InvalidOperationException("Informe o nome da rubrica.");
            if (FindRubricaRow(ExcelHelper.GetSheet(sheetName), _registry.GetBySheet(sheetName), nome) >= 0)
                throw new InvalidOperationException("Rubrica já existe nesta unidade.");

            Dictionary<string, object> result = null;
            ExcelHelper.RunWithExcel(() =>
            {
                var cfg = _registry.GetBySheet(sheetName);
                var ws = ExcelHelper.GetSheet(sheetName);
                int row = Math.Max(cfg.DataStartRowMatrix, ExcelHelper.LastRow(ws, cfg.RubricaCol, cfg.DataStartRowMatrix) + 1);
                if (!string.IsNullOrWhiteSpace(secao))
                    ExcelHelper.SetCell(ws, string.IsNullOrEmpty(cfg.SectionCol) ? "A" : cfg.SectionCol, row, secao.Trim());
                ExcelHelper.SetCell(ws, cfg.RubricaCol, row, nome);
                string section = GetSectionAtRow(ws, cfg, row);
                result = new Dictionary<string, object>
                {
                    ["nome"] = nome,
                    ["natureza"] = ClassifyNatureza(nome, section),
                    ["secao"] = section,
                    ["row"] = row
                };
            });
            return result;
        }

        public void RemoveRubrica(string sheetName, string nome)
        {
            nome = (nome ?? "").Trim();
            if (string.IsNullOrEmpty(nome)) throw new InvalidOperationException("Informe o nome da rubrica.");

            ExcelHelper.RunWithExcel(() =>
            {
                var cfg = _registry.GetBySheet(sheetName);
                var ws = ExcelHelper.GetSheet(sheetName);
                int row = FindRubricaRow(ws, cfg, nome);
                if (row < 0) throw new InvalidOperationException("Rubrica não encontrada.");

                int start = ExcelHelper.ColIndex(cfg.MonthStartCol);
                int end = ExcelHelper.ColIndex(string.IsNullOrEmpty(cfg.MonthEndCol) ? "P" : cfg.MonthEndCol);
                for (int idx = start; idx <= end; idx++)
                {
                    string col = ExcelHelper.ColLetter(idx);
                    if (IsReadOnlyColumn(ws, cfg, col)) continue;
                    if (Math.Abs(ExcelHelper.ToCents(ExcelHelper.GetCell(ws, col, row))) > 0)
                        throw new InvalidOperationException("A rubrica possui valores lançados. Zere-os antes de excluir.");
                }

                ((Excel.Range)ws.Rows[row]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            });
        }

        public List<Dictionary<string, object>> ListPremissas(string entity)
        {
            var cfg = _registry.GetByEntity(entity);
            return List(cfg.SheetName);
        }

        private int FindRubricaRow(Excel.Worksheet ws, SheetConfig cfg, string rubrica)
        {
            int last = ExcelHelper.LastRow(ws, cfg.RubricaCol, cfg.DataStartRowMatrix);
            for (int r = cfg.DataStartRowMatrix; r <= last; r++)
            {
                if (ExcelHelper.CellStr(ExcelHelper.GetCell(ws, cfg.RubricaCol, r)).Trim() == rubrica.Trim())
                    return r;
            }
            return -1;
        }

        private string GetMonthName(Excel.Worksheet ws, SheetConfig cfg, string col)
        {
            string h = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, col, cfg.MonthHeaderRow));
            if (string.IsNullOrEmpty(h))
                h = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, col, cfg.MonthHeaderRow - 1));
            return (h ?? "").Trim().ToUpperInvariant();
        }

        private string ResolveColumn(Excel.Worksheet ws, SheetConfig cfg, string mes, string tipo, string colunaHint = null)
        {
            if (!string.IsNullOrWhiteSpace(colunaHint))
                return ResolveColumnHint(ws, cfg, colunaHint);

            if (string.IsNullOrEmpty(mes)) throw new InvalidOperationException("Mês é obrigatório.");
            string mesNorm = NormalizeMes(mes);
            string tipoNorm = NormalizeTipo(tipo);
            var matches = new List<MatrixColumn>();

            foreach (var colInfo in GetMatrixColumns(ws, cfg))
            {
                if (colInfo.ReadOnly) continue;
                string monthName = NormalizeMes(colInfo.Mes);
                if (string.IsNullOrEmpty(monthName)) continue;
                if (!MesMatches(monthName, mesNorm)) continue;
                matches.Add(colInfo);
            }

            if (!matches.Any())
                throw new InvalidOperationException("Coluna não encontrada para mês: " + mes);

            if (!string.IsNullOrEmpty(tipoNorm))
            {
                var tipoHit = matches.FirstOrDefault(c => TipoMatches(c.Tipo, tipoNorm));
                if (tipoHit != null) return tipoHit.Col;
            }

            if (matches.Count == 1)
                return matches[0].Col;

            var tiposDisp = string.Join(", ", matches.Select(c => c.Tipo).Distinct());
            throw new InvalidOperationException(
                string.IsNullOrEmpty(tipoNorm)
                    ? "Coluna não encontrada para mês: " + mes
                    : "Coluna não encontrada para mês: " + mes + " e tipo: " + tipo
                      + ". Tipos disponíveis: " + tiposDisp);
        }

        private string ResolveColumnHint(Excel.Worksheet ws, SheetConfig cfg, string colunaHint)
        {
            string col = colunaHint.Trim().ToUpperInvariant();
            foreach (var colInfo in GetMatrixColumns(ws, cfg))
            {
                if (colInfo.ReadOnly) continue;
                if (colInfo.Col.Equals(col, StringComparison.OrdinalIgnoreCase))
                    return colInfo.Col;
            }
            throw new InvalidOperationException("Coluna inválida: " + colunaHint);
        }

        /// <summary>Colunas D..P — exclui TOTAL para edição.</summary>
        private List<MatrixColumn> GetMatrixColumns(Excel.Worksheet ws, SheetConfig cfg)
        {
            var cols = new List<MatrixColumn>();
            int start = ExcelHelper.ColIndex(cfg.MonthStartCol);
            int end = ExcelHelper.ColIndex(string.IsNullOrEmpty(cfg.MonthEndCol) ? "P" : cfg.MonthEndCol);
            for (int idx = start; idx <= end; idx++)
            {
                string col = ExcelHelper.ColLetter(idx);
                string mes = GetMonthName(ws, cfg, col);
                if (string.IsNullOrWhiteSpace(mes)) continue;
                bool readOnly = IsReadOnlyColumn(ws, cfg, col, mes);
                cols.Add(new MatrixColumn
                {
                    Col = col,
                    Mes = mes.Trim().ToUpperInvariant(),
                    Tipo = GetTipo(ws, cfg, col, idx - start, mes),
                    ReadOnly = readOnly
                });
            }
            return cols;
        }

        private string GetTipo(Excel.Worksheet ws, SheetConfig cfg, string col, int offsetFromStart, string mes)
        {
            int[] rows = { cfg.TypeHeaderRow, cfg.TypeHeaderRow - 1, cfg.TypeHeaderRow + 1 };
            foreach (int row in rows)
            {
                if (row < 1) continue;
                string norm = NormalizeTipo(ReadMergedCellText(ws, col, row));
                if (norm == "Realizado" || norm == "Previsto") return norm;
            }
            return InferTipoFromMonth(mes, offsetFromStart);
        }

        /// <summary>Jan–Abr = Realizado; Mai–Dez = Previsto (layout padrão D..O).</summary>
        private static string InferTipoFromMonth(string mes, int offsetFromStart)
        {
            string mesNorm = NormalizeMes(mes);
            if (!string.IsNullOrEmpty(mesNorm))
            {
                for (int i = 0; i < Meses.Length; i++)
                {
                    if (!MesMatches(NormalizeMes(Meses[i]), mesNorm)) continue;
                    return i <= 3 ? "Realizado" : "Previsto";
                }
            }
            return offsetFromStart <= 3 ? "Realizado" : "Previsto";
        }

        private static string ReadMergedCellText(Excel.Worksheet ws, string col, int row)
        {
            var range = (Excel.Range)ws.Range[col + row];
            object value;
            if (range.MergeCells is bool merged && merged)
                value = range.MergeArea.Cells[1, 1].Value2;
            else
                value = range.Value2;
            return ExcelHelper.CellStr(value);
        }

        /// <summary>Tipo é por coluna/mês — grava na linha de cabeçalho (ex.: H8).</summary>
        private static void SetColumnTipoHeader(Excel.Worksheet ws, SheetConfig cfg, string col, string tipo)
        {
            tipo = NormalizeTipo(tipo);
            if (tipo != "Realizado" && tipo != "Previsto") return;
            int row = cfg.TypeHeaderRow;
            var range = (Excel.Range)ws.Range[col + row];
            if (range.MergeCells is bool merged && merged)
                range.MergeArea.Cells[1, 1].Value2 = tipo;
            else
                ExcelHelper.SetCell(ws, col, row, tipo);
        }

        private static bool IsReadOnlyColumn(Excel.Worksheet ws, SheetConfig cfg, string col, string mes = null)
        {
            if (string.IsNullOrEmpty(mes))
                mes = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, col, cfg.MonthHeaderRow));
            if (string.IsNullOrWhiteSpace(mes) && cfg.MonthHeaderRow > 1)
                mes = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, col, cfg.MonthHeaderRow - 1));
            mes = (mes ?? "").Trim();
            if (mes.IndexOf("TOTAL", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            string tipo = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, col, cfg.TypeHeaderRow));
            if (tipo.IndexOf("TOTAL", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (tipo.IndexOf("REALIZADO + PREVISTO", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        private string GetSectionAtRow(Excel.Worksheet ws, SheetConfig cfg, int targetRow)
        {
            string sectionCol = string.IsNullOrEmpty(cfg.SectionCol) ? "A" : cfg.SectionCol;
            string current = "";
            for (int r = cfg.DataStartRowMatrix; r <= targetRow; r++)
            {
                string a = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, sectionCol, r));
                if (!string.IsNullOrWhiteSpace(a)) current = a.Trim();
            }
            return current;
        }

        public static string ClassifyNatureza(string rubrica, string sectionColA = "")
        {
            if (string.IsNullOrWhiteSpace(rubrica)) return "Outros";
            string r = rubrica.Trim();
            if (r.StartsWith("(+)", StringComparison.OrdinalIgnoreCase)) return "Receita";
            if (r.StartsWith("(-)", StringComparison.OrdinalIgnoreCase)) return "Despesa";
            if (r.StartsWith("(=)", StringComparison.OrdinalIgnoreCase)) return "Resultado";
            if (!string.IsNullOrWhiteSpace(sectionColA))
            {
                if (sectionColA.IndexOf("Receita", StringComparison.OrdinalIgnoreCase) >= 0) return "Receita";
                if (sectionColA.IndexOf("Despesa", StringComparison.OrdinalIgnoreCase) >= 0) return "Despesa";
            }
            return "Outros";
        }

        private sealed class MatrixColumn
        {
            public string Col { get; set; }
            public string Mes { get; set; }
            public string Tipo { get; set; }
            public bool ReadOnly { get; set; }
        }

        private static string NormalizeMes(string m)
        {
            if (string.IsNullOrEmpty(m)) return "";
            m = m.ToUpperInvariant().Trim();
            m = m.Replace("Ç", "C").Replace("Á", "A").Replace("Ã", "A").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Ú", "U");
            return m;
        }

        private static string NormalizeTipo(string t)
        {
            if (string.IsNullOrWhiteSpace(t)) return "";
            t = t.Trim();
            if (t.IndexOf("Prev", StringComparison.OrdinalIgnoreCase) >= 0) return "Previsto";
            if (t.IndexOf("Real", StringComparison.OrdinalIgnoreCase) >= 0) return "Realizado";
            return t;
        }

        private static bool MesMatches(string monthName, string mesNorm)
        {
            if (string.IsNullOrEmpty(monthName) || string.IsNullOrEmpty(mesNorm)) return false;
            return monthName.StartsWith(mesNorm, StringComparison.Ordinal)
                || mesNorm.StartsWith(monthName, StringComparison.Ordinal);
        }

        private static bool TipoMatches(string colTipo, string tipoNorm)
        {
            if (string.IsNullOrEmpty(tipoNorm)) return true;
            return NormalizeTipo(colTipo).Equals(tipoNorm, StringComparison.OrdinalIgnoreCase);
        }
    }
}
