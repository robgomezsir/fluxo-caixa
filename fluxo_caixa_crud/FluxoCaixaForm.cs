using Microsoft.Web.WebView2.Core;
using System;
using System.IO;
using System.Windows.Forms;
using fluxo_caixa_crud.Models;
using fluxo_caixa_crud.Services;

namespace fluxo_caixa_crud
{
    public partial class FluxoCaixaForm : Form
    {
        private static FluxoCaixaForm _instance;
        private WebViewBridge _bridge;
        private bool _ready;
        private bool _allowClose;
        private bool _promptPending;

        public FluxoCaixaForm()
        {
            InitializeComponent();
            MinimumSize = new System.Drawing.Size(900, 600);
            Text = "Fluxo de Caixa";
            StartPosition = FormStartPosition.CenterScreen;
            Size = new System.Drawing.Size(1280, 800);
            FormClosed += FluxoCaixaForm_FormClosed;
            FormClosing += FluxoCaixaForm_FormClosing;
            IniciarWebview();
        }

        private void FluxoCaixaForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_allowClose)
            {
                PrepareForShutdown();
                return;
            }

            if (_promptPending)
            {
                e.Cancel = true;
                return;
            }

            if (!ExcelHelper.IsHostWorkbookDirty())
            {
                _allowClose = true;
                return;
            }

            e.Cancel = true;
            _promptPending = true;
            BeginInvoke(new Action(PromptSaveAndClose));
        }

        private void PromptSaveAndClose()
        {
            _promptPending = false;

            string fileName = ExcelHelper.GetHostWorkbookDisplayName();
            var result = MessageBox.Show(
                this,
                $"Deseja salvar as alterações em \"{fileName}\" antes de fechar?",
                "Fluxo de Caixa",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (result == DialogResult.Cancel)
                return;

            if (result == DialogResult.Yes)
            {
                try
                {
                    ExcelHelper.SaveHostWorkbook();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        this,
                        "Não foi possível salvar a planilha.\n\n" + ex.Message,
                        "Fluxo de Caixa",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }
            }

            _allowClose = true;
            Close();
        }

        private void PrepareForShutdown()
        {
            _ready = false;
            try
            {
                if (webView21?.CoreWebView2 != null)
                    webView21.CoreWebView2.WebMessageReceived -= CoreWebView2_WebMessageReceived;
            }
            catch { }
        }

        public static void ShowMain(bool hideExcel = true)
        {
            ExcelHelper.EnsureApplicationReady();

            if (_instance == null || _instance.IsDisposed)
                _instance = new FluxoCaixaForm();

            if (!_instance.Visible)
                _instance.Show();
            else if (_instance.WindowState == FormWindowState.Minimized)
                _instance.WindowState = FormWindowState.Normal;

            _instance.BringToFront();
            _instance.Activate();

            if (hideExcel)
                ExcelHelper.HideApplication();
        }

        private void FluxoCaixaForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _instance = null;
            _allowClose = false;
            ExcelHelper.ScheduleCloseHostWorkbook();
        }

        private async void IniciarWebview()
        {
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "fluxo_caixa_crud");

                var environment = await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: userDataFolder);

                await webView21.EnsureCoreWebView2Async(environment);

                _bridge = new WebViewBridge();
                _bridge.EnsureInitialized();

                webView21.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
                webView21.CoreWebView2.Settings.IsStatusBarEnabled = false;
                _ready = true;

                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "frontend", "index.html");
                if (File.Exists(filePath))
                    webView21.CoreWebView2.Navigate(new Uri(filePath).AbsoluteUri);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao iniciar interface: " + ex.Message, "Fluxo de Caixa",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            string json = e.TryGetWebMessageAsString();
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => ProcessMessage(json)));
                return;
            }
            ProcessMessage(json);
        }

        private void ProcessMessage(string json)
        {
            if (!_ready || _bridge == null)
            {
                PostResponse(ApiJson.Serialize(ApiResponse.Fail("", "Interface ainda não está pronta.")));
                return;
            }

            try
            {
                PostResponse(_bridge.HandleMessage(json));
            }
            catch (Exception ex)
            {
                PostResponse(ApiJson.Serialize(ApiResponse.Fail("", ex.Message)));
            }
        }

        private void PostResponse(string response)
        {
            if (webView21?.CoreWebView2 != null)
                webView21.CoreWebView2.PostWebMessageAsString(response);
        }
    }
}
