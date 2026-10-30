using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    internal static class ExcelInstanceHelper
    {
        public static bool IsHostWorkbook(Excel.Workbook wb)
        {
            if (wb == null) return false;

            return IsHostIdentity(wb, GetHostWorkbookPath(), GetHostWorkbookName());
        }

        /// <summary>
        /// Identidade da pasta de trabalho hospedeira: caminho completo. O nome é usado
        /// somente quando o host ainda não tem caminho salvo. Assim cópias do arquivo em
        /// outra pasta (ex.: Documents\FluxoCaixa, E:\PROJETOS...) nunca são confundidas
        /// com a planilha hospedeira do add-in.
        /// </summary>
        internal static bool IsHostIdentity(Excel.Workbook wb, string hostPath, string hostName)
        {
            if (wb == null) return false;

            if (!string.IsNullOrEmpty(hostPath)
                && string.Equals(wb.FullName, hostPath, StringComparison.OrdinalIgnoreCase))
                return true;

            if (string.IsNullOrEmpty(hostPath)
                && !string.IsNullOrEmpty(hostName)
                && string.Equals(wb.Name, hostName, StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private static string GetHostWorkbookPath()
        {
            try { return Globals.ThisWorkbook.FullName; }
            catch { return null; }
        }

        private static string GetHostWorkbookName()
        {
            try { return Globals.ThisWorkbook.Name; }
            catch { return null; }
        }

        public static bool IsIgnorableWorkbook(Excel.Workbook wb)
        {
            if (wb == null) return true;

            try
            {
                string name = wb.Name ?? "";
                if (string.Equals(name, "PERSONAL.XLSB", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "PERSONAL", StringComparison.OrdinalIgnoreCase))
                    return true;

                string path = wb.FullName ?? "";
                if (path.IndexOf("XLSTART", StringComparison.OrdinalIgnoreCase) >= 0
                    || path.IndexOf(@"\STARTUP\", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }

            return false;
        }
    }
}
