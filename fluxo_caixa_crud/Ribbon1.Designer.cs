namespace fluxo_caixa_crud
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnMostrarPlanilha = this.Factory.CreateRibbonButton();
            this.btnAjuda = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnSalvar = this.Factory.CreateRibbonButton();
            this.btnRecalcular = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Fluxo de Caixa";
            this.tab1.Name = "tab1";
            this.tab1.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.btnMostrarPlanilha);
            this.group1.Items.Add(this.btnAjuda);
            this.group1.Label = "Interface";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Abrir Formulário";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btnMostrarPlanilha
            // 
            this.btnMostrarPlanilha.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMostrarPlanilha.Label = "Mostrar Planilha";
            this.btnMostrarPlanilha.OfficeImageId = "WindowSwitchToWindow";
            this.btnMostrarPlanilha.Name = "btnMostrarPlanilha";
            this.btnMostrarPlanilha.ShowImage = true;
            this.btnMostrarPlanilha.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMostrarPlanilha_Click);
            // 
            // btnAjuda
            // 
            this.btnAjuda.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnAjuda.Label = "Ajuda";
            this.btnAjuda.OfficeImageId = "Help";
            this.btnAjuda.Name = "btnAjuda";
            this.btnAjuda.ShowImage = true;
            this.btnAjuda.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAjuda_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnSalvar);
            this.group2.Items.Add(this.btnRecalcular);
            this.group2.Label = "Planilha";
            this.group2.Name = "group2";
            // 
            // btnSalvar
            // 
            this.btnSalvar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSalvar.Label = "Salvar";
            this.btnSalvar.OfficeImageId = "FileSave";
            this.btnSalvar.Name = "btnSalvar";
            this.btnSalvar.ShowImage = true;
            this.btnSalvar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSalvar_Click);
            // 
            // btnRecalcular
            // 
            this.btnRecalcular.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRecalcular.Label = "Recalcular";
            this.btnRecalcular.OfficeImageId = "CalculateSheet";
            this.btnRecalcular.Name = "btnRecalcular";
            this.btnRecalcular.ShowImage = true;
            this.btnRecalcular.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRecalcular_Click);
            // 
            // Ribbon1
            // 
            // Global = true: a aba "Fluxo de Caixa" fica sempre visível na Ribbon, mesmo
            // quando outra pasta de trabalho (não a hospedeira) está ativa. Sem isso o VSTO
            // document-level mostra a aba apenas enquanto a Fluxo_Caixa.xlsx estiver em foco.
            this.Global = true;
            // StartFromScratch = false: preserva as abas nativas do Excel e ADICIONA a nossa
            // antes de "Página Inicial", em vez de substituir a Ribbon inteira.
            this.StartFromScratch = false;
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMostrarPlanilha;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAjuda;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalvar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRecalcular;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
