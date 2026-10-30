using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using fluxo_caixa_crud.Services;

namespace fluxo_caixa_crud
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // A aba "Fluxo de Caixa" fica registrada para ser mantida visível/selecionada
            // sempre que o Excel voltar a ficar em primeiro plano.
            RibbonTabActivator.Registrar(tab1, Global, StartFromScratch);
            RibbonTabActivator.AtivarAbaFluxo();

            ApplyFluxoIcon();
        }

        private void ApplyFluxoIcon()
        {
            try
            {
                bool dark = RibbonThemeHelper.IsDarkOfficeTheme();
                button1.Image = dark
                    ? Properties.Resources.RibbonFluxoEscuro
                    : Properties.Resources.RibbonFluxoClaro;
            }
            catch { /* mantém botão sem imagem se recurso ausente */ }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            FluxoCaixaForm.ShowMain(hideExcel: true);
        }

        private void btnMostrarPlanilha_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelper.ShowApplication();
        }

        private void btnSalvar_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string caminho = ExcelHelper.SaveHostWorkbook();
                MessageBox.Show(
                    "Planilha salva com sucesso em:\n\n" + caminho,
                    "Fluxo de Caixa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Não foi possível salvar a planilha.\n\n" + ex.Message,
                    "Fluxo de Caixa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnRecalcular_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ExcelHelper.RunWithExcel(() => { });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Não foi possível recalcular a planilha.\n\n" + ex.Message,
                    "Fluxo de Caixa",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnAjuda_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "Interface do Fluxo de Caixa\n\n" +
                "• Abrir Formulário: abre a tela de gestão do Fluxo de Caixa.\n" +
                "• Mostrar Planilha: traz a janela do Excel de volta (útil quando o formulário cobre a base).\n" +
                "• Salvar: recalcula as fórmulas e salva a planilha que hospeda o sistema.\n" +
                "• Recalcular: recalcula todas as fórmulas da pasta de trabalho.\n\n" +
                "Dica: também existe um botão \"Abrir Fluxo de Caixa\" dentro da planilha, " +
                "na primeira linha, à direita dos dados — use-o caso o formulário não abra sozinho.",
                "Ajuda — Fluxo de Caixa",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
