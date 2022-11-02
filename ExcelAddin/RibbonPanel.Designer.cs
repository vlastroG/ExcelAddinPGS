namespace ExcelAddin
{
    partial class RibbonPanel : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonPanel()
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
            this.group = this.Factory.CreateRibbonGroup();
            this.btnMultipleCSV = this.Factory.CreateRibbonButton();
            this.btnMultipleCsvSettings = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group
            // 
            this.group.Items.Add(this.btnMultipleCSV);
            this.group.Items.Add(this.btnMultipleCsvSettings);
            this.group.Label = "PGS-Excel";
            this.group.Name = "group";
            // 
            // btnMultipleCSV
            // 
            this.btnMultipleCSV.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMultipleCSV.Image = global::ExcelAddin.Properties.Resources.Info;
            this.btnMultipleCSV.Label = "CSV import";
            this.btnMultipleCSV.Name = "btnMultipleCSV";
            this.btnMultipleCSV.ScreenTip = "Импорт нескольких CSV файлов на лист";
            this.btnMultipleCSV.ShowImage = true;
            this.btnMultipleCSV.SuperTip = "По выбранной ячейке, которая будет левым верхним углом вставки, импортируются на " +
    "лист все выбранные в папке CSV файлы друг под другом";
            this.btnMultipleCSV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMultipleCSV_Click);
            // 
            // btnMultipleCsvSettings
            // 
            this.btnMultipleCsvSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMultipleCsvSettings.Image = global::ExcelAddin.Properties.Resources.Info;
            this.btnMultipleCsvSettings.Label = "Настройки разделителя";
            this.btnMultipleCsvSettings.Name = "btnMultipleCsvSettings";
            this.btnMultipleCsvSettings.ShowImage = true;
            this.btnMultipleCsvSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMultipleCsvSettings_Click);
            // 
            // RibbonPanel
            // 
            this.Name = "RibbonPanel";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonPanel_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group.ResumeLayout(false);
            this.group.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMultipleCSV;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMultipleCsvSettings;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonPanel RibbonPanel
        {
            get { return this.GetRibbon<RibbonPanel>(); }
        }
    }
}
