using System;
using System.Collections.Generic;
using System.Linq;
using fluxo_caixa_crud.Models;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    public class TabularCrudService
    {
        private readonly SheetRegistry _registry;

        public TabularCrudService(SheetRegistry registry)
        {
            _registry = registry;
        }

        public List<Dictionary<string, object>> List(string entity)
        {
            var cfg = _registry.GetByEntity(entity);
            var ws = ExcelHelper.GetSheet(cfg.SheetName);
            var result = new List<Dictionary<string, object>>();
            int last = ExcelHelper.LastRow(ws, GetScanColumn(cfg), cfg.DataStartRow);
            for (int r = cfg.DataStartRow; r <= last; r++)
            {
                var row = ReadRow(ws, cfg, r);
                if (IsEmptyRow(row, cfg)) continue;
                row["id"] = GetOrAssignId(ws, cfg, r);
                row["_row"] = r;
                result.Add(row);
            }
            return result;
        }

        public Dictionary<string, object> Get(string entity, string id)
        {
            var (ws, cfg, row) = FindById(entity, id);
            var data = ReadRow(ws, cfg, row);
            data["id"] = id;
            data["_row"] = row;
            return data;
        }

        public Dictionary<string, object> Create(string entity, JObject payload)
        {
            var cfg = _registry.GetByEntity(entity);
            Dictionary<string, object> created = null;
            ExcelHelper.RunWithExcel(() =>
            {
                var ws = ExcelHelper.GetSheet(cfg.SheetName);
                int row = Math.Max(cfg.DataStartRow, ExcelHelper.LastRow(ws, GetScanColumn(cfg), cfg.DataStartRow) + 1);
                string id = Guid.NewGuid().ToString("N");
                ExcelHelper.SetCell(ws, cfg.IdColumn, row, id);
                WriteRow(ws, cfg, row, payload);
                created = ReadRow(ws, cfg, row);
                created["id"] = id;
                created["_row"] = row;
                LogAudit("create", entity, id);
            });
            return created;
        }

        public Dictionary<string, object> Update(string entity, string id, JObject payload)
        {
            Dictionary<string, object> updated = null;
            ExcelHelper.RunWithExcel(() =>
            {
                var (ws, cfg, row) = FindById(entity, id);
                WriteRow(ws, cfg, row, payload);
                updated = ReadRow(ws, cfg, row);
                updated["id"] = id;
                updated["_row"] = row;
                LogAudit("update", entity, id);
            });
            return updated;
        }

        public void Delete(string entity, string id)
        {
            ExcelHelper.RunWithExcel(() =>
            {
                var (ws, cfg, row) = FindById(entity, id);
                ((Excel.Range)ws.Rows[row]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                LogAudit("delete", entity, id);
            });
        }

        private string GetScanColumn(SheetConfig cfg)
        {
            var f = cfg.Fields.FirstOrDefault(x => x.Key == "unidade" || x.Key == "nomeProjeto" || x.Key == "descricao");
            return f?.Column ?? cfg.Fields.FirstOrDefault(x => x.Key != "id")?.Column ?? "B";
        }

        private (Excel.Worksheet ws, SheetConfig cfg, int row) FindById(string entity, string id)
        {
            var cfg = _registry.GetByEntity(entity);
            var ws = ExcelHelper.GetSheet(cfg.SheetName);
            int last = ExcelHelper.LastRow(ws, cfg.IdColumn, cfg.DataStartRow);
            for (int r = cfg.DataStartRow; r <= last; r++)
            {
                if (ExcelHelper.CellStr(ExcelHelper.GetCell(ws, cfg.IdColumn, r)) == id)
                    return (ws, cfg, r);
            }
            throw new InvalidOperationException("Registro não encontrado: " + id);
        }

        private string GetOrAssignId(Excel.Worksheet ws, SheetConfig cfg, int row)
        {
            string id = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, cfg.IdColumn, row));
            if (string.IsNullOrEmpty(id))
            {
                id = Guid.NewGuid().ToString("N");
                ExcelHelper.SetCell(ws, cfg.IdColumn, row, id);
            }
            return id;
        }

        private Dictionary<string, object> ReadRow(Excel.Worksheet ws, SheetConfig cfg, int row)
        {
            var d = new Dictionary<string, object>();
            foreach (var f in cfg.Fields.Where(x => x.Key != "id"))
            {
                object raw = ExcelHelper.GetCell(ws, f.Column, row);
                if (f.Type == "money")
                    d[f.Key] = ExcelHelper.ToCents(raw);
                else if (f.Type == "number" || f.Type == "percent")
                    d[f.Key] = ExcelHelper.CellNum(raw);
                else
                    d[f.Key] = ExcelHelper.CellStr(raw);
            }
            return d;
        }

        private void WriteRow(Excel.Worksheet ws, SheetConfig cfg, int row, JObject payload)
        {
            foreach (var f in cfg.Fields.Where(x => x.Key != "id"))
            {
                if (payload[f.Key] == null) continue;
                if (f.Type == "money")
                    ExcelHelper.SetCell(ws, f.Column, row, ExcelHelper.FromCents(payload[f.Key].Value<long>()));
                else if (f.Type == "number" || f.Type == "percent")
                    ExcelHelper.SetCell(ws, f.Column, row, payload[f.Key].Value<double>());
                else
                    ExcelHelper.SetCell(ws, f.Column, row, payload[f.Key].ToString());
            }
        }

        private bool IsEmptyRow(Dictionary<string, object> row, SheetConfig cfg)
        {
            foreach (var f in cfg.Fields.Where(x => x.Key != "id"))
            {
                if (row.TryGetValue(f.Key, out var v))
                {
                    if (v is long l && l != 0) return false;
                    if (v is double d && Math.Abs(d) > 0.0001) return false;
                    if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) return false;
                }
            }
            return true;
        }

        private void LogAudit(string action, string entity, string id)
        {
            try
            {
                var ws = ExcelHelper.GetOrCreateSheet("_CRUD_LOG");
                ws.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                int row = ExcelHelper.LastRow(ws, "A", 1) + 1;
                if (row == 1) { ExcelHelper.SetCell(ws, "A", 1, "Timestamp"); ExcelHelper.SetCell(ws, "B", 1, "Action"); ExcelHelper.SetCell(ws, "C", 1, "Entity"); ExcelHelper.SetCell(ws, "D", 1, "Id"); row = 2; }
                ExcelHelper.SetCell(ws, "A", row, DateTime.Now.ToString("o"));
                ExcelHelper.SetCell(ws, "B", row, action);
                ExcelHelper.SetCell(ws, "C", row, entity);
                ExcelHelper.SetCell(ws, "D", row, id);
            }
            catch { }
        }
    }
}
