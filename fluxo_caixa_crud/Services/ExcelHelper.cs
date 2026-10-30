using System;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud.Services
{
    public static class ExcelHelper
    {
        public static Excel.Application App => Globals.ThisWorkbook.Application;

        public static Excel.Worksheet GetSheet(string name)
        {
            foreach (Excel.Worksheet ws in Globals.ThisWorkbook.Worksheets)
            {
                if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                    return ws;
            }
            throw new InvalidOperationException("Aba não encontrada: " + name);
        }

        public static Excel.Worksheet GetOrCreateSheet(string name)
        {
            try { return GetSheet(name); }
            catch
            {
                var ws = (Excel.Worksheet)Globals.ThisWorkbook.Worksheets.Add(
                    After: Globals.ThisWorkbook.Worksheets[Globals.ThisWorkbook.Worksheets.Count]);
                ws.Name = name;
                return ws;
            }
        }

        public static Excel.Worksheet CopySheet(string templateName, string newName)
        {
            var template = GetSheet(templateName);
            template.Copy(After: template);
            var wb = Globals.ThisWorkbook;
            var ws = (Excel.Worksheet)wb.Worksheets[template.Index + 1];
            ws.Name = newName;
            return ws;
        }

        public static void DeleteSheet(string name)
        {
            var ws = GetSheet(name);
            if (Globals.ThisWorkbook.Worksheets.Count <= 1)
                throw new InvalidOperationException("Não é possível remover a única aba da planilha.");
            ws.Delete();
        }

        public static string ColLetter(int col)
        {
            string s = "";
            while (col > 0)
            {
                col--;
                s = (char)('A' + col % 26) + s;
                col /= 26;
            }
            return s;
        }

        public static int ColIndex(string col)
        {
            int idx = 0;
            foreach (char c in col.ToUpperInvariant())
                idx = idx * 26 + (c - 'A' + 1);
            return idx;
        }

        public static object GetCell(Excel.Worksheet ws, string col, int row)
        {
            return ws.Range[col + row].Value2;
        }

        public static void SetCell(Excel.Worksheet ws, string col, int row, object value)
        {
            ws.Range[col + row].Value2 = value ?? "";
        }

        public static string CellStr(object v)
        {
            if (v == null) return "";
            return Convert.ToString(v).Trim();
        }

        public static double CellNum(object v)
        {
            if (v == null) return 0;
            if (v is double d) return d;
            if (v is float f) return f;
            if (v is decimal dec) return (double)dec;
            if (v is int i) return i;
            if (v is long l) return l;
            if (v is short s) return s;

            string str = Convert.ToString(v, CultureInfo.InvariantCulture)?.Trim();
            if (string.IsNullOrEmpty(str)) return 0;

            bool looksPtBr = str.IndexOf(',') >= 0;
            if (looksPtBr)
            {
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.GetCultureInfo("pt-BR"), out d))
                    return d;
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;
            }
            else
            {
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;
                if (double.TryParse(str, NumberStyles.Any, CultureInfo.GetCultureInfo("pt-BR"), out d))
                    return d;
            }
            if (double.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                return d;
            return 0;
        }

        public static long ToCents(object v)
        {
            return (long)Math.Round(CellNum(v) * 100);
        }

        public static double FromCents(long cents)
        {
            return cents / 100.0;
        }

        public static int LastRow(Excel.Worksheet ws, string col, int startRow)
        {
            var used = ws.UsedRange;
            int max = used.Row + used.Rows.Count - 1;
            for (int r = max; r >= startRow; r--)
            {
                if (!string.IsNullOrWhiteSpace(CellStr(GetCell(ws, col, r))))
                    return r;
            }
            return startRow - 1;
        }

        public static void RunWithExcel(Action action)
        {
            bool screen = App.ScreenUpdating;
            try
            {
                App.ScreenUpdating = false;
                action();
                App.CalculateFullRebuild();
            }
            finally
            {
                App.ScreenUpdating = screen;
            }
        }

        /// <summary>Recalcula fórmulas e grava o workbook hospedeiro. Nada roda após Save para não re-marcar como alterado.</summary>
        public static string SaveHostWorkbook()
        {
            var wb = Globals.ThisWorkbook;
            if (wb.ReadOnly)
                throw new InvalidOperationException("A planilha está somente leitura e não pode ser salva.");

            bool screen = App.ScreenUpdating;
            try
            {
                App.ScreenUpdating = false;
                App.CalculateFullRebuild();
                wb.Save();
            }
            finally
            {
                App.ScreenUpdating = screen;
            }
            return wb.FullName;
        }

        public static bool IsHostWorkbookDirty()
        {
            try
            {
                return !Globals.ThisWorkbook.Saved;
            }
            catch
            {
                return false;
            }
        }

        private const string NomeBotaoAbrirForm = "btnAbrirInterfaceFluxo";

        /// <summary>
        /// Cria na planilha ativa um botão para reabrir a interface do Fluxo de Caixa
        /// (útil quando o formulário some ou a planilha é aberta junto com outras).
        /// Criado no ThisWorkbook_Startup, pois o host item Sheet1 pode não existir na
        /// pasta de trabalho em produção.
        /// </summary>
        public static void CreateOpenFormButton()
        {
            try
            {
                Excel.Worksheet interop = App.ActiveSheet as Excel.Worksheet;
                if (interop == null)
                    interop = (Excel.Worksheet)Globals.ThisWorkbook.Worksheets[1];
                LogBotao("Start: sheet=" + interop.Name);

                // Container estendido com os metodos Add* ( ControlExtensions.AddButton )
                Microsoft.Office.Tools.Excel.Worksheet vsto = Globals.Factory.GetVstoObject(interop);
                LogBotao("GetVstoObject: " + (vsto == null ? "NULL" : "ok"));

                // Evita duplicar em reaberturas: remove shapes anteriores do botao.
                // O shape pode persistir com nome interno (_ActiveXWrapperN), entao procuramos
                // tambem pelo progID do host de WinForms do VSTO.
                try
                {
                    for (int i = interop.Shapes.Count; i >= 1; i--)
                    {
                        try
                        {
                            var sh = interop.Shapes.Item(i);
                            if (sh == null) continue;
                            string shName = Convert.ToString(sh.Name);
                            if (shName == NomeBotaoAbrirForm)
                            {
                                sh.Delete();
                                continue;
                            }
                            if (Convert.ToInt32(sh.Type) == 12) // msoOLEControlObject
                            {
                                string progId = Convert.ToString(sh.OLEFormat.progID) ?? "";
                                if (progId.StartsWith("WinForms.Control.Host"))
                                    sh.Delete();
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // Posição: primeira coluna à direita do intervalo usado, na linha 1 (fora dos dados)
                Excel.Range used = interop.UsedRange;
                int col = used.Column + used.Columns.Count;
                if (col < 3) col = 3;
                if (col >= interop.Columns.Count - 2) col = 2;
                Excel.Range ancora = (Excel.Range)interop.Cells[1, col];

                Microsoft.Office.Tools.Excel.ControlExtensions.AddButton(
                    vsto.Controls, (double)ancora.Left, (double)ancora.Top, 150.0, 30.0, NomeBotaoAbrirForm);
                var botao = vsto.Controls[NomeBotaoAbrirForm] as System.Windows.Forms.Button;
                botao.Text = "Abrir Fluxo de Caixa";
                botao.Click += (s, e) =>
                {
                    // hideExcel: false — a partir do botão, mantém a planilha visível junto com o form
                    FluxoCaixaForm.ShowMain(hideExcel: false);
                };
                LogBotao("Botao criado em " + interop.Name + " col=" + col);
            }
            catch (Exception ex)
            {
                LogBotao("ERRO: " + ex.GetType().Name + ": " + ex.Message + " | " + ex.StackTrace);
                System.Diagnostics.Debug.WriteLine("Botão de reabrir o form: " + ex.Message);
            }
        }

        private static void LogBotao(string msg)
        {
            try
            {
                string dir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "fluxo_caixa_crud");
                System.IO.Directory.CreateDirectory(dir);
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(dir, "botao.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + msg + Environment.NewLine);
            }
            catch { }
        }

        public static string GetHostWorkbookDisplayName()
        {
            try
            {
                var wb = Globals.ThisWorkbook;
                string name = wb.Name;
                if (!string.IsNullOrWhiteSpace(name)) return name;
                return System.IO.Path.GetFileName(wb.FullName) ?? "Fluxo de Caixa";
            }
            catch
            {
                return "Fluxo de Caixa";
            }
        }

        public static void EnsureApplicationReady()
        {
            try
            {
                App.ScreenUpdating = true;
                App.EnableEvents = true;
                App.DisplayAlerts = true;
            }
            catch { }
        }

        public static void HideApplication()
        {
            try
            {
                // Não esconder o Excel se o usuário tiver outras planilhas abertas na mesma instância,
                // pois elas desapareceriam junto com a janela.
                if (HasOtherOpenWorkbooks(App, Globals.ThisWorkbook.FullName, Globals.ThisWorkbook.Name))
                    return;

                App.ScreenUpdating = false;
                App.Visible = false;
            }
            catch { }
        }

        public static void ShowApplication()
        {
            try
            {
                EnsureApplicationReady();
                App.Visible = true;
                if (Globals.ThisWorkbook.Windows.Count > 0)
                {
                    Excel.Window win = App.ActiveWindow;
                    if (win == null)
                        win = (Excel.Window)Globals.ThisWorkbook.Windows[1];
                    win.WindowState = Excel.XlWindowState.xlNormal;
                    win.Activate();
                }

                // O Excel voltou a ficar visível: a aba "Fluxo de Caixa" deve continuar à vista.
                if (IsHostWorkbookActive())
                    RibbonTabActivator.AtivarAbaFluxo();
            }
            catch { }
        }

        /// <summary>True quando a Ribbon do Excel pode ser desenhada (janela não oculta/não minimizada).</summary>
        public static bool IsApplicationVisible()
        {
            try { return App.Visible; }
            catch { return true; }
        }

        /// <summary>True se a planilha hospedeira do Fluxo é a planilha ativa.</summary>
        public static bool IsHostWorkbookActive()
        {
            try { return ExcelInstanceHelper.IsHostWorkbook(App.ActiveWorkbook); }
            catch { return false; }
        }

        /// <summary>
        /// Coloca em foco a janela de um workbook (inclusive aberto pelo usuário enquanto a
        /// interface estava em primeiro plano e o Excel, oculto). Sem este passo o arquivo
        /// aberto via Explorer/DDE entrava na mesma instância oculta e o usuário via o
        /// Excel "abrir" sem a planilha aparecer.
        /// </summary>
        public static void ShowWorkbook(Excel.Workbook wb, Excel.Window win = null)
        {
            try
            {
                EnsureApplicationReady();
                App.Visible = true;

                if (wb != null)
                {
                    try { wb.Activate(); } catch { }
                }

                Excel.Window alvo = win;
                if (alvo == null && wb != null && wb.Windows.Count > 0)
                    alvo = (Excel.Window)wb.Windows[1];

                if (alvo != null)
                {
                    try
                    {
                        if (alvo.WindowState == Excel.XlWindowState.xlMinimized)
                            alvo.WindowState = Excel.XlWindowState.xlNormal;
                    }
                    catch { }

                    try { alvo.Activate(); } catch { }
                }

                // Só mexe no estado da janela quando ela estava minimizada: chamar Activate a
                // cada evento de foco empurraria o Excel para a frente indevidamente. O foco em
                // si já é garantido por alvo.Activate() acima.
                try
                {
                    if (App.WindowState == Excel.XlWindowState.xlMinimized)
                        App.WindowState = Excel.XlWindowState.xlNormal;
                }
                catch { }
            }
            catch { }
        }

        /// <summary>True quando a janela informada está minimizada.</summary>
        public static bool IsWindowMinimized(Excel.Window win)
        {
            if (win == null) return false;
            try { return win.WindowState == Excel.XlWindowState.xlMinimized; }
            catch { return false; }
        }

        private static System.Windows.Forms.Timer _closeTimer;

        public static void CancelScheduledClose()
        {
            if (_closeTimer == null) return;
            _closeTimer.Stop();
            _closeTimer.Dispose();
            _closeTimer = null;
        }

        /// <summary>Agenda o fechamento da planilha Fluxo fora do ciclo FormClosing/FormClosed (evita deadlock COM).</summary>
        public static void ScheduleCloseHostWorkbook(int delayMs = 300)
        {
            CancelScheduledClose();

            _closeTimer = new System.Windows.Forms.Timer { Interval = delayMs };
            _closeTimer.Tick += (s, e) =>
            {
                CancelScheduledClose();
                CloseHostWorkbook();
            };
            _closeTimer.Start();
        }

        /// <summary>Fecha somente o workbook Fluxo de Caixa. Outros arquivos Excel abertos permanecem.</summary>
        public static void CloseHostWorkbook()
        {
            Excel.Application app = null;
            string hostPath = null;
            string hostName = null;

            try
            {
                app = Globals.ThisWorkbook.Application;
                hostPath = Globals.ThisWorkbook.FullName;
                hostName = Globals.ThisWorkbook.Name;
            }
            catch
            {
                return;
            }

            if (app == null)
                return;

            bool otherWorkbooksOpen = HasOtherOpenWorkbooks(app, hostPath, hostName);
            bool appWasVisible = false;
            try { appWasVisible = app.Visible; } catch { }

            bool displayAlerts = true;
            bool enableEvents = true;
            try { displayAlerts = app.DisplayAlerts; } catch { }
            try { enableEvents = app.EnableEvents; } catch { }

            try
            {
                app.DisplayAlerts = false;
                app.EnableEvents = false;
                app.ScreenUpdating = true;

                if (!TryCloseVstoHostWorkbook())
                    TryCloseWorkbookByIdentity(app, hostPath, hostName);

                if (otherWorkbooksOpen)
                {
                    EnsureApplicationReady();
                    ActivateAnotherWorkbook(app, hostPath, hostName);
                    if (!appWasVisible)
                    {
                        try { app.Visible = true; } catch { }
                    }
                }
                else if (CountMeaningfulWorkbooks(app) == 0)
                {
                    try { app.Quit(); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("CloseHostWorkbook: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (CountOpenWorkbooks(app) > 0)
                    {
                        app.DisplayAlerts = displayAlerts;
                        app.EnableEvents = enableEvents;
                        EnsureApplicationReady();
                    }
                }
                catch { }
            }
        }

        private static bool TryCloseVstoHostWorkbook()
        {
            try
            {
                Globals.ThisWorkbook.Close(false, Type.Missing, Type.Missing);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryCloseWorkbook(Excel.Workbook wb)
        {
            if (wb == null) return false;
            try
            {
                wb.Close(false, Type.Missing, Type.Missing);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void TryCloseWorkbookByIdentity(Excel.Application app, string fullPath, string name)
        {
            if (app == null) return;

            Excel.Workbook match = FindHostWorkbook(app, fullPath, name);
            if (match != null)
                TryCloseWorkbook(match);
        }

        private static Excel.Workbook FindHostWorkbook(Excel.Application app, string fullPath, string name)
        {
            try
            {
                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    // Identidade por caminho completo. Sem o antigo critério por nome
                    // ("Fluxo"), uma cópia do arquivo em outra pasta nunca é fechada
                    // por engano no lugar da planilha hospedeira.
                    if (ExcelInstanceHelper.IsHostIdentity(wb, fullPath, name))
                        return wb;
                }
            }
            catch { }
            return null;
        }

        private static bool HasOtherOpenWorkbooks(Excel.Application app, string hostPath, string hostName)
        {
            try
            {
                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    if (ExcelInstanceHelper.IsIgnorableWorkbook(wb))
                        continue;

                    if (ExcelInstanceHelper.IsHostIdentity(wb, hostPath, hostName))
                        continue;

                    return true;
                }
            }
            catch { }
            return false;
        }

        private static int CountOpenWorkbooks(Excel.Application app)
        {
            try { return app.Workbooks.Count; }
            catch { return 0; }
        }

        private static int CountMeaningfulWorkbooks(Excel.Application app)
        {
            try
            {
                int count = 0;
                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    if (ExcelInstanceHelper.IsIgnorableWorkbook(wb)) continue;
                    count++;
                }
                return count;
            }
            catch { return 0; }
        }

        private static void ActivateAnotherWorkbook(Excel.Application app, string hostPath, string hostName)
        {
            try
            {
                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    if (ExcelInstanceHelper.IsIgnorableWorkbook(wb))
                        continue;

                    if (ExcelInstanceHelper.IsHostIdentity(wb, hostPath, hostName))
                        continue;

                    wb.Activate();
                    if (wb.Windows.Count > 0)
                        ((Excel.Window)wb.Windows[1]).Activate();
                    return;
                }
            }
            catch { }
        }
    }
}
