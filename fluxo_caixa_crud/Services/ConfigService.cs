using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using fluxo_caixa_crud.Models;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    public class ConfigService
    {
        private const string LookupsSheetName = "_LOOKUPS";
        private const string DefaultUnitTemplate = "ALS";

        private readonly SheetRegistry _registry;
        private readonly MatrixCrudService _matrix;

        public ConfigService(SheetRegistry registry, MatrixCrudService matrix)
        {
            _registry = registry;
            _matrix = matrix;
        }

        public object DownloadWorkbook()
        {
            string savedPath = null;
            ExcelHelper.RunWithExcel(() =>
            {
                var wb = Globals.ThisWorkbook;
                string defaultName = Path.GetFileNameWithoutExtension(wb.Name) + "_backup.xlsx";
                using (var dlg = new SaveFileDialog())
                {
                    dlg.Filter = "Excel (*.xlsx)|*.xlsx";
                    dlg.FileName = defaultName;
                    dlg.Title = "Salvar cópia do Fluxo de Caixa";
                    if (dlg.ShowDialog() != DialogResult.OK)
                        throw new InvalidOperationException("Download cancelado.");
                    wb.SaveCopyAs(dlg.FileName);
                    savedPath = dlg.FileName;
                }
            });
            return new { path = savedPath, fileName = Path.GetFileName(savedPath) };
        }

        public List<Dictionary<string, object>> ListResponsaveis()
        {
            EnsureLookupsSheet();
            var fromConfig = new HashSet<string>(ReadLookupValues("responsavel"), StringComparer.OrdinalIgnoreCase);
            var fromSheet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var ws = ExcelHelper.GetSheet("ANOTAÇÕES DE PREMISSAS");
                int last = ExcelHelper.LastRow(ws, "D", 3);
                for (int r = 3; r <= last; r++)
                {
                    string v = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "D", r));
                    if (!string.IsNullOrWhiteSpace(v)) fromSheet.Add(v.Trim());
                }
            }
            catch { }

            var merged = new HashSet<string>(fromConfig, StringComparer.OrdinalIgnoreCase);
            foreach (var s in fromSheet) merged.Add(s);

            return merged.OrderBy(x => x).Select(n =>
            {
                bool inConfig = fromConfig.Contains(n);
                bool inSheet = fromSheet.Contains(n);
                string origem = inConfig && inSheet ? "ambos" : inConfig ? "config" : "planilha";
                return new Dictionary<string, object> { ["nome"] = n, ["origem"] = origem, ["removivel"] = inConfig };
            }).ToList();
        }

        public List<Dictionary<string, object>> ListUnidades()
        {
            var labels = UnitDisplayService.ReadAllUnitLabels(_registry);
            return _registry.MatrixSheets
                .Where(m => m.Entity != null && m.Entity.StartsWith("matrix_") &&
                            !m.Entity.Contains("premissas") && !m.Entity.Contains("capex0403"))
                .OrderBy(m => labels.ContainsKey(m.SheetName.Trim()) ? labels[m.SheetName.Trim()].Label : m.SheetName)
                .Select(m =>
                {
                    string sheet = m.SheetName.Trim();
                    UnitDisplayInfo info;
                    labels.TryGetValue(sheet, out info);
                    return new Dictionary<string, object>
                    {
                        ["nome"] = sheet,
                        ["label"] = info?.Label ?? sheet,
                        ["entity"] = m.Entity
                    };
                }).ToList();
        }

        public List<Dictionary<string, object>> ListRubricas(string sheet)
        {
            if (string.IsNullOrWhiteSpace(sheet))
                throw new InvalidOperationException("Informe payload.sheet.");
            return _matrix.GetRubricas(sheet);
        }

        public Dictionary<string, object> Create(string entity, JObject payload)
        {
            switch (entity.ToLowerInvariant())
            {
                case "config_responsaveis":
                    return AddResponsavel(payload["nome"]?.ToString());
                case "config_unidades":
                    return AddUnidade(payload["nome"]?.ToString(), payload["templateSheet"]?.ToString());
                case "config_rubricas":
                    return _matrix.AddRubrica(
                        payload["sheet"]?.ToString(),
                        payload["nome"]?.ToString(),
                        payload["secao"]?.ToString());
                default:
                    throw new InvalidOperationException("Entidade de configuração desconhecida: " + entity);
            }
        }

        public void Delete(string entity, JObject payload)
        {
            switch (entity.ToLowerInvariant())
            {
                case "config_responsaveis":
                    RemoveResponsavel(payload["nome"]?.ToString());
                    break;
                case "config_unidades":
                    RemoveUnidade(payload["nome"]?.ToString());
                    break;
                case "config_rubricas":
                    _matrix.RemoveRubrica(payload["sheet"]?.ToString(), payload["nome"]?.ToString());
                    break;
                default:
                    throw new InvalidOperationException("Entidade de configuração desconhecida: " + entity);
            }
        }

        private Dictionary<string, object> AddResponsavel(string nome)
        {
            nome = (nome ?? "").Trim();
            if (nome.Length == 0) throw new InvalidOperationException("Informe o nome do responsável.");
            if (nome.Length > 80) throw new InvalidOperationException("Nome muito longo (máx. 80 caracteres).");

            ExcelHelper.RunWithExcel(() =>
            {
                EnsureLookupsSheet();
                var values = ReadLookupValues("responsavel");
                if (values.Any(v => v.Equals(nome, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Responsável já cadastrado.");
                AppendLookup("responsavel", nome);
            });
            ReloadLookups();
            return new Dictionary<string, object> { ["nome"] = nome, ["origem"] = "config" };
        }

        private void RemoveResponsavel(string nome)
        {
            nome = (nome ?? "").Trim();
            if (nome.Length == 0) throw new InvalidOperationException("Informe o nome do responsável.");

            ExcelHelper.RunWithExcel(() =>
            {
                EnsureLookupsSheet();
                if (!RemoveLookup("responsavel", nome))
                    throw new InvalidOperationException("Responsável não encontrado na lista de configuração.");
            });
            ReloadLookups();
        }

        private Dictionary<string, object> AddUnidade(string nome, string templateSheet)
        {
            nome = (nome ?? "").Trim();
            if (nome.Length == 0) throw new InvalidOperationException("Informe o nome da unidade (aba).");
            if (nome.Length > 31) throw new InvalidOperationException("Nome da aba deve ter no máximo 31 caracteres.");
            if (nome.StartsWith("_", StringComparison.Ordinal))
                throw new InvalidOperationException("Nome de aba inválido.");

            templateSheet = string.IsNullOrWhiteSpace(templateSheet) ? DefaultUnitTemplate : templateSheet.Trim();
            string entity = SheetRegistry.EntityFromSheetName(nome);

            ExcelHelper.RunWithExcel(() =>
            {
                try { ExcelHelper.GetSheet(nome); throw new InvalidOperationException("Já existe uma aba com esse nome."); }
                catch (InvalidOperationException ex) when (ex.Message.Contains("Já existe")) { throw; }
                catch { }

                ExcelHelper.CopySheet(templateSheet, nome);
                AppendConfigRow(new SheetConfig
                {
                    SheetName = nome,
                    Entity = entity,
                    Type = SheetType.Matrix,
                    MonthHeaderRow = 7,
                    TypeHeaderRow = 8,
                    DataStartRowMatrix = 11,
                    RubricaCol = "B",
                    MonthStartCol = "D",
                    MonthEndCol = "P",
                    SectionCol = "A",
                    MonthCount = 12
                });
            });

            _registry.Initialize();
            return new Dictionary<string, object> { ["nome"] = nome, ["entity"] = entity };
        }

        private void RemoveUnidade(string nome)
        {
            nome = (nome ?? "").Trim();
            if (nome.Length == 0) throw new InvalidOperationException("Informe o nome da unidade.");
            if (string.Equals(nome, DefaultUnitTemplate, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Não é permitido remover a unidade modelo (" + DefaultUnitTemplate + ").");

            ExcelHelper.RunWithExcel(() =>
            {
                var cfg = _registry.GetBySheet(nome);
                if (cfg.Type != SheetType.Matrix)
                    throw new InvalidOperationException("A aba informada não é uma unidade matricial.");

                if (_matrix.HasLaunchValues(nome))
                    throw new InvalidOperationException("A unidade possui lançamentos com valor. Zere os valores antes de excluir.");

                RemoveConfigRow(nome);
                ExcelHelper.DeleteSheet(nome);
            });

            _registry.Initialize();
        }

        private void ReloadLookups()
        {
            _registry.Lookups.Load(_registry);
        }

        internal void EnsureLookupsSheet()
        {
            Excel.Worksheet ws;
            try
            {
                ws = ExcelHelper.GetSheet(LookupsSheetName);
                if (ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", 1)) == "Type") return;
            }
            catch
            {
                ws = ExcelHelper.GetOrCreateSheet(LookupsSheetName);
            }

            ws.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            ws.Cells.Clear();
            ExcelHelper.SetCell(ws, "A", 1, "Type");
            ExcelHelper.SetCell(ws, "B", 1, "Value");
        }

        internal static List<string> ReadLookupValues(string type)
        {
            var list = new List<string>();
            try
            {
                var ws = ExcelHelper.GetSheet(LookupsSheetName);
                int row = 2;
                while (true)
                {
                    string t = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", row));
                    if (string.IsNullOrEmpty(t)) break;
                    if (t.Equals(type, StringComparison.OrdinalIgnoreCase))
                    {
                        string v = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "B", row));
                        if (!string.IsNullOrWhiteSpace(v)) list.Add(v.Trim());
                    }
                    row++;
                }
            }
            catch { }
            return list;
        }

        private static void AppendLookup(string type, string value)
        {
            var ws = ExcelHelper.GetSheet(LookupsSheetName);
            int row = 2;
            while (!string.IsNullOrEmpty(ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", row))))
                row++;
            ExcelHelper.SetCell(ws, "A", row, type);
            ExcelHelper.SetCell(ws, "B", row, value);
        }

        private static bool RemoveLookup(string type, string value)
        {
            var ws = ExcelHelper.GetSheet(LookupsSheetName);
            int row = 2;
            while (true)
            {
                string t = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", row));
                if (string.IsNullOrEmpty(t)) break;
                string v = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "B", row));
                if (t.Equals(type, StringComparison.OrdinalIgnoreCase) &&
                    v.Equals(value, StringComparison.OrdinalIgnoreCase))
                {
                    ((Excel.Range)ws.Rows[row]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    return true;
                }
                row++;
            }
            return false;
        }

        private static void AppendConfigRow(SheetConfig cfg)
        {
            var ws = ExcelHelper.GetSheet("_CONFIG");
            int row = 2;
            while (!string.IsNullOrEmpty(ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", row))))
                row++;
            WriteConfigRow(ws, row, cfg);
        }

        private static void RemoveConfigRow(string sheetName)
        {
            var ws = ExcelHelper.GetSheet("_CONFIG");
            int row = 2;
            while (true)
            {
                string name = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "A", row));
                if (string.IsNullOrEmpty(name)) break;
                if (name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    ((Excel.Range)ws.Rows[row]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    return;
                }
                row++;
            }
        }

        private static void WriteConfigRow(Excel.Worksheet ws, int row, SheetConfig c)
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
        }
    }
}
