using System;
using System.Windows.Forms;
using fluxo_caixa_crud.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace fluxo_caixa_crud
{
    public partial class ThisWorkbook
    {
        private bool _eventosDoExcelConectados;

        private void ThisWorkbook_Startup(object sender, EventArgs e)
        {
            ExcelHelper.CancelScheduledClose();
            ExcelHelper.EnsureApplicationReady();

            // A interface esconde o Excel (Visible = false). Sem estes eventos, um arquivo
            // aberto pelo usuário (Explorer/DDE) entrava nesta mesma instância oculta e o
            // usuário via o Excel "abrir" sem nenhuma planilha aparecer.
            ConectarEventosDoExcel();

            // Botão na planilha para reabrir o form. Criado aqui (e não no Sheet1) porque
            // o host item Sheet1 pode não existir nesta pasta de trabalho.
            try
            {
                ExcelHelper.CreateOpenFormButton();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Botão de reabrir o form: " + ex.Message);
            }

            // Abre normalmente, junto com as demais planilhas abertas.
            // Sem relançamento em instância dedicada e sem mover/fechar arquivos do usuário.

            try
            {
                var bridge = new WebViewBridge();
                bridge.EnsureInitialized();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Init _CONFIG: " + ex.Message);
            }

            var timer = new Timer { Interval = 400 };
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                timer.Dispose();
                try
                {
                    FluxoCaixaForm.ShowMain(hideExcel: true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Auto-open form: " + ex.Message);
                    MessageBox.Show(
                        "Não foi possível abrir a interface do Fluxo de Caixa.\n\n" + ex.Message,
                        "Fluxo de Caixa",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            };
            timer.Start();
        }

        private void ThisWorkbook_Shutdown(object sender, EventArgs e)
        {
            DesconectarEventosDoExcel();
        }

        #region Eventos do aplicativo Excel

        /// <summary>
        /// Conecta os eventos do aplicativo. Sem eles, o comportamento anormal era:
        /// com o Excel oculto pela interface, o usuário abria outro arquivo e ele caía
        /// nesta instância invisível (sem janela à vista).
        /// </summary>
        private void ConectarEventosDoExcel()
        {
            if (_eventosDoExcelConectados) return;

            try
            {
                // Cast explícito: Application expõe NewWorkbook/WorkbookOpen/WindowActivate ao
                // mesmo tempo como propriedade e como evento (AppEvents_Event); sem o cast o
                // compilador acusa ambiguidade.
                Excel.AppEvents_Event eventos = (Excel.AppEvents_Event)ExcelHelper.App;
                eventos.WorkbookOpen += App_WorkbookOpen;
                eventos.NewWorkbook += App_NewWorkbook;
                eventos.WindowActivate += App_WindowActivate;
                _eventosDoExcelConectados = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ConectarEventosDoExcel: " + ex.Message);
            }
        }

        private void DesconectarEventosDoExcel()
        {
            if (!_eventosDoExcelConectados) return;
            _eventosDoExcelConectados = false;

            try
            {
                Excel.AppEvents_Event eventos = (Excel.AppEvents_Event)ExcelHelper.App;
                eventos.WorkbookOpen -= App_WorkbookOpen;
                eventos.NewWorkbook -= App_NewWorkbook;
                eventos.WindowActivate -= App_WindowActivate;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("DesconectarEventosDoExcel: " + ex.Message);
            }
        }

        /// <summary>Outro arquivo aberto: se a planilha do Fluxo não for a origem, mostra e foca a nova.</summary>
        private void App_WorkbookOpen(object wb)
        {
            TratarWorkbookAberto(wb);
        }

        /// <summary>Nova pasta de trabalho criada pelo usuário enquanto a interface está em primeiro plano.</summary>
        private void App_NewWorkbook(object wb)
        {
            TratarWorkbookAberto(wb);
        }

        /// <summary>Normaliza o tratamento de WorkbookOpen/NewWorkbook (assinaturas variam por versão do Office).</summary>
        private void TratarWorkbookAberto(object wbObj)
        {
            try
            {
                var wb = wbObj as Excel.Workbook;
                if (wb == null) return;

                if (ExcelInstanceHelper.IsHostWorkbook(wb))
                {
                    RibbonTabActivator.AtivarAbaFluxo();
                    return;
                }

                ExcelHelper.ShowWorkbook(wb);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("TratarWorkbookAberto: " + ex.Message);
            }
        }

        /// <summary>
        /// Janela em foco mudou. Rede de segurança para os casos em que o arquivo novo já
        /// existia na instância oculta: se o usuário ativou outra planilha, o Excel precisa
        /// estar visível para que ela apareça.
        /// </summary>
        private void App_WindowActivate(object wb, Excel.Window wnd)
        {
            try
            {
                var workbook = wb as Excel.Workbook;
                if (workbook != null && ExcelInstanceHelper.IsHostWorkbook(workbook))
                {
                    RibbonTabActivator.AtivarAbaFluxo();
                    return;
                }

                if (!ExcelHelper.IsApplicationVisible() || ExcelHelper.IsWindowMinimized(wnd))
                    ExcelHelper.ShowWorkbook(workbook, wnd);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("App_WindowActivate: " + ex.Message);
            }
        }

        #endregion

        #region VSTO Designer generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion
    }
}
