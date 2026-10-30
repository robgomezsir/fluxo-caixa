using System;
using System.Diagnostics;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace fluxo_caixa_crud.Services
{
    /// <summary>
    /// Mantém a aba "Fluxo de Caixa" sempre visível e selecionada na Ribbon do Excel.
    ///
    /// O acesso aos membros da aba é feito por reflexão porque RibbonTab.ActivateTab e
    /// RibbonTab.Visible só existem no runtime VSTO 2010 ou superior, enquanto o projeto
    /// compila contra a ref-assembly do Visual Studio, que nem sempre os expõe. Assim o
    /// add-in compila e roda em qualquer versão instalada: o passo inexistente é simplesmente
    /// ignorado, sem exceção em tempo de execução.
    /// </summary>
    internal static class RibbonTabActivator
    {
        private static RibbonTab _aba;
        private static bool _resolvido;
        private static MethodInfo _metodoAtivarAba;

        /// <summary>Guarda a aba customizada assim que a Ribbon carrega (chamado no Ribbon1_Load).</summary>
        public static void Registrar(RibbonTab aba, bool global, bool startFromScratch)
        {
            _aba = aba;
            _resolvido = false;
            _metodoAtivarAba = null;

            bool temVisivel = aba != null && BuscarPropriedade(aba.GetType(), "Visible") != null;

            RegistrarLog("Ribbon carregada. aba=" + (aba == null ? "null" : aba.Name)
                + " global=" + global
                + " startFromScratch=" + startFromScratch
                + " ActivateTab=" + (aba != null && aba.GetType().GetMethod("ActivateTab", Type.EmptyTypes) != null)
                + " propVisible=" + temVisivel);
        }

        /// <summary>
        /// Torna a aba visível e a traz para frente. Nunca lança exceção: a Ribbon do Excel
        /// não pode ser derrubada por uma falha de foco.
        /// </summary>
        public static void AtivarAbaFluxo()
        {
            try
            {
                RibbonTab aba = _aba;
                if (aba == null)
                    return; // Ribbon ainda não carregou; o próximo Ribbon1_Load resolve.

                // O Excel reconstrói a Ribbon ao trocar a janela ativa. Sem invalidar, a
                // aba pode continuar "desenhada" como ausente até o Excel ser forçado a
                // redesenhar a UI.
                Invalidar();

                if (!_resolvido)
                {
                    _metodoAtivarAba = aba.GetType().GetMethod("ActivateTab", Type.EmptyTypes);
                    _resolvido = true;
                }

                TornarVisivel(aba);

                if (_metodoAtivarAba != null)
                {
                    _metodoAtivarAba.Invoke(aba, null);
                    return;
                }

                AtivarViaRibbonUi(aba);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("AtivarAbaFluxo: " + ex.Message);
            }
        }

        /// <summary>Força o Excel a redesenhar a Ribbon personalizada (IRibbonUI.Invalidate).</summary>
        public static void Invalidar()
        {
            try
            {
                RibbonTab aba = _aba;
                if (aba == null)
                    return;

                IRibbonUI ui = aba.RibbonUI;
                if (ui != null)
                    ui.Invalidate();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Invalidar: " + ex.Message);
            }
        }

        /// <summary>Desfaz o estado "oculto pelo usuário" na personalização da Ribbon.</summary>
        private static void TornarVisivel(RibbonTab aba)
        {
            try
            {
                PropertyInfo propriedade = BuscarPropriedade(aba.GetType(), "Visible");
                if (propriedade == null || !propriedade.CanWrite)
                    return;

                if (!(bool)propriedade.GetValue(aba, null))
                    propriedade.SetValue(aba, true, null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("TornarVisivel: " + ex.Message);
            }
        }

        /// <summary>Caminho alternativo (Office 2010+): IRibbonUI.ActivateTab(id da aba).</summary>
        private static void AtivarViaRibbonUi(RibbonTab aba)
        {
            try
            {
                IRibbonUI ribbonUi = aba.RibbonUI;
                if (ribbonUi == null)
                    return;

                string id = ObterIdDaAba(aba);
                if (string.IsNullOrEmpty(id))
                    return;

                ribbonUi.ActivateTab(id);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("AtivarViaRibbonUi: " + ex.Message);
            }
        }

        /// <summary>RibbonFactory.GetRibbonId também é reflexão: só existe no VSTO 2010+.</summary>
        private static string ObterIdDaAba(RibbonTab aba)
        {
            try
            {
                var factory = Globals.Factory;
                if (factory == null)
                    return null;

                MethodInfo metodo = factory.GetType().GetMethod("GetRibbonId", new[] { typeof(object) });
                if (metodo == null)
                    return null;

                return metodo.Invoke(factory, new object[] { aba }) as string;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ObterIdDaAba: " + ex.Message);
                return null;
            }
        }

        private static PropertyInfo BuscarPropriedade(Type tipo, string nome)
        {
            for (Type atual = tipo; atual != null; atual = atual.BaseType)
            {
                PropertyInfo propriedade = atual.GetProperty(
                    nome,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

                if (propriedade != null)
                    return propriedade;
            }
            return null;
        }

        /// <summary>
        /// Diagnóstico do carregamento da Ribbon em %LOCALAPPDATA%\fluxo_caixa_crud\ribbon.log.
        /// Se a aba não aparecer, este arquivo diz se a Ribbon chegou a ser carregada.
        /// </summary>
        private static void RegistrarLog(string mensagem)
        {
            try
            {
                string dir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "fluxo_caixa_crud");
                System.IO.Directory.CreateDirectory(dir);
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(dir, "ribbon.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + mensagem + Environment.NewLine);
            }
            catch { /* log nunca pode derrubar a Ribbon */ }
        }
    }
}
