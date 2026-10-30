using System;
using System.Collections.Generic;
using System.Linq;

namespace fluxo_caixa_crud.Services
{
    public sealed class UnitDisplayInfo
    {
        public string Sheet { get; set; }
        public string Code { get; set; }
        public string Nome { get; set; }
        public string Label { get; set; }
    }

    public static class UnitDisplayService
    {
        public static bool IsUnitMatrix(string entity)
        {
            if (string.IsNullOrEmpty(entity) || !entity.StartsWith("matrix_", StringComparison.OrdinalIgnoreCase))
                return false;
            string e = entity.ToLowerInvariant();
            return !e.Contains("premissas") && !e.Contains("capex0403");
        }

        public static UnitDisplayInfo Read(string sheetName)
        {
            string sheet = (sheetName ?? "").Trim();
            if (string.IsNullOrEmpty(sheet))
                return new UnitDisplayInfo { Sheet = "", Code = "", Nome = "", Label = "" };

            try
            {
                var ws = ExcelHelper.GetSheet(sheet);
                var lines = new List<string>();
                for (int r = 4; r <= 9; r++)
                {
                    string t = ExcelHelper.CellStr(ExcelHelper.GetCell(ws, "B", r));
                    if (!string.IsNullOrWhiteSpace(t))
                        lines.Add(t.Trim());
                }
                return BuildLabel(sheet, lines);
            }
            catch
            {
                return new UnitDisplayInfo { Sheet = sheet, Code = sheet, Nome = sheet, Label = sheet };
            }
        }

        internal static UnitDisplayInfo BuildLabel(string sheet, List<string> lines)
        {
            string code = sheet;
            string nome = null;

            foreach (var line in lines)
            {
                if (line.IndexOf("Fluxo de Caixa", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string part = ExtractFluxoNome(line);
                    if (!string.IsNullOrWhiteSpace(part))
                        nome = part;
                }
            }

            if (string.IsNullOrWhiteSpace(nome))
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    var line = lines[i];
                    if (line.IndexOf("Unidade de Neg", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    int idx = line.IndexOf(':');
                    if (idx < 0) continue;
                    string v = line.Substring(idx + 1).Trim();
                    if (!string.IsNullOrWhiteSpace(v))
                    {
                        nome = v;
                        break;
                    }
                    if (i + 1 < lines.Count && !lines[i + 1].StartsWith("Fluxo de Caixa", StringComparison.OrdinalIgnoreCase))
                    {
                        nome = lines[i + 1].Trim();
                        break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(nome))
                nome = code;

            string label = FormatLabel(code, nome);
            return new UnitDisplayInfo { Sheet = sheet, Code = code, Nome = nome, Label = label };
        }

        public static string FormatLabel(string code, string nome)
        {
            code = (code ?? "").Trim();
            nome = (nome ?? "").Trim();
            if (string.IsNullOrEmpty(code)) return nome;
            if (string.IsNullOrEmpty(nome) || nome.Equals(code, StringComparison.OrdinalIgnoreCase))
                return code;
            if (nome.StartsWith(code, StringComparison.OrdinalIgnoreCase))
                return code + " - " + nome;
            return code + " - " + nome;
        }

        private static string ExtractFluxoNome(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return "";
            string s = line.Trim();
            int fc = s.IndexOf("Fluxo de Caixa", StringComparison.OrdinalIgnoreCase);
            if (fc >= 0)
                s = s.Substring(fc + "Fluxo de Caixa".Length).Trim();

            foreach (var sep in new[] { " – ", " - ", "–", "-" })
            {
                int i = s.IndexOf(sep, StringComparison.Ordinal);
                if (i < 0) continue;
                string tail = s.Substring(i + sep.Length).Trim();
                if (!string.IsNullOrWhiteSpace(tail))
                    return tail;
            }

            return s.Trim();
        }

        public static Dictionary<string, UnitDisplayInfo> ReadAllUnitLabels(SheetRegistry registry)
        {
            var map = new Dictionary<string, UnitDisplayInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var cfg in registry.MatrixSheets.Where(c => IsUnitMatrix(c.Entity)))
            {
                var info = Read(cfg.SheetName);
                map[cfg.SheetName.Trim()] = info;
            }
            return map;
        }
    }
}
