namespace GKL_Analise
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.GKL_Analise = this.Factory.CreateRibbonGroup();
            this.bScan = this.Factory.CreateRibbonButton();
            this.bFill1 = this.Factory.CreateRibbonButton();
            this.bFill2 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.GKL_Analise.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.GKL_Analise);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // GKL_Analise
            // 
            this.GKL_Analise.Items.Add(this.bScan);
            this.GKL_Analise.Items.Add(this.bFill1);
            this.GKL_Analise.Items.Add(this.bFill2);
            this.GKL_Analise.Label = "GKL Analise";
            this.GKL_Analise.Name = "GKL_Analise";
            // 
            // bScan
            // 
            this.bScan.Label = "Сканировать";
            this.bScan.Name = "bScan";
            this.bScan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bScan_Click);
            // 
            // bFill1
            // 
            this.bFill1.Label = "Заполнить Л1";
            this.bFill1.Name = "bFill1";
            this.bFill1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bFill1_Click);
            // 
            // bFill2
            // 
            this.bFill2.Label = "Заполнить Л2";
            this.bFill2.Name = "bFill2";
            this.bFill2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.bFill2_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.GKL_Analise.ResumeLayout(false);
            this.GKL_Analise.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GKL_Analise;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bScan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bFill1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton bFill2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
