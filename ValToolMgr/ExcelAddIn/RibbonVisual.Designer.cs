namespace ExcelAddIn
{
    partial class RibbonVisual : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonVisual()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonVisual));
            this.TabValToolMgr = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.LayoutVersion_DD = this.Factory.CreateRibbonDropDown();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button_NewPR = this.Factory.CreateRibbonButton();
            this.addCategory = this.Factory.CreateRibbonButton();
            this.plan2Tests = this.Factory.CreateRibbonButton();
            this.TestEditGroup = this.Factory.CreateRibbonGroup();
            this.AddStep = this.Factory.CreateRibbonButton();
            this.RemoveStep = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.AddActionVar = this.Factory.CreateRibbonButton();
            this.RemoveActionVar = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.AddCheckVar = this.Factory.CreateRibbonButton();
            this.RemoveCheckVar = this.Factory.CreateRibbonButton();
            this.checks = this.Factory.CreateRibbonGroup();
            this.testCheck = this.Factory.CreateRibbonButton();
            this.Outputs = this.Factory.CreateRibbonGroup();
            this.toSwVTD = this.Factory.CreateRibbonButton();
            this.testStand = this.Factory.CreateRibbonButton();
            this.toSwVTDR = this.Factory.CreateRibbonButton();
            this.macroInfos = this.Factory.CreateRibbonGroup();
            this.help = this.Factory.CreateRibbonButton();
            this.macroVersion = this.Factory.CreateRibbonLabel();
            this.UpdateDate = this.Factory.CreateRibbonLabel();
            this.autoUpdate = this.Factory.CreateRibbonCheckBox();
            this.TabValToolMgr.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.TestEditGroup.SuspendLayout();
            this.checks.SuspendLayout();
            this.Outputs.SuspendLayout();
            this.macroInfos.SuspendLayout();
            // 
            // TabValToolMgr
            // 
            this.TabValToolMgr.Groups.Add(this.group1);
            this.TabValToolMgr.Groups.Add(this.group2);
            this.TabValToolMgr.Groups.Add(this.TestEditGroup);
            this.TabValToolMgr.Groups.Add(this.checks);
            this.TabValToolMgr.Groups.Add(this.Outputs);
            this.TabValToolMgr.Groups.Add(this.macroInfos);
            this.TabValToolMgr.Label = "Val tool Mgr";
            this.TabValToolMgr.Name = "TabValToolMgr";
            // 
            // group1
            // 
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.LayoutVersion_DD);
            this.group1.Label = "Layout Version";
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // label1
            // 
            this.label1.Label = "Choose a version";
            this.label1.Name = "label1";
            // 
            // LayoutVersion_DD
            // 
            this.LayoutVersion_DD.Enabled = false;
            this.LayoutVersion_DD.Label = " ";
            this.LayoutVersion_DD.Name = "LayoutVersion_DD";
            this.LayoutVersion_DD.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LayoutVersion_DD_SelectionChanged);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button_NewPR);
            this.group2.Items.Add(this.addCategory);
            this.group2.Items.Add(this.plan2Tests);
            this.group2.Label = "Editing";
            this.group2.Name = "group2";
            // 
            // button_NewPR
            // 
            this.button_NewPR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_NewPR.Image = global::ExcelAddIn.Properties.Resources.NewPR;
            this.button_NewPR.Label = "New PR";
            this.button_NewPR.Name = "button_NewPR";
            this.button_NewPR.ShowImage = true;
            this.button_NewPR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_NewPR_Click);
            // 
            // addCategory
            // 
            this.addCategory.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.addCategory.Label = "Add a Category";
            this.addCategory.Name = "addCategory";
            this.addCategory.ShowImage = true;
            this.addCategory.Visible = false;
            this.addCategory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addCategory_Click);
            // 
            // plan2Tests
            // 
            this.plan2Tests.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.plan2Tests.Image = ((System.Drawing.Image)(resources.GetObject("plan2Tests.Image")));
            this.plan2Tests.Label = "Plan to Tests";
            this.plan2Tests.Name = "plan2Tests";
            this.plan2Tests.ShowImage = true;
            this.plan2Tests.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.plan2Tests_Click);
            // 
            // TestEditGroup
            // 
            this.TestEditGroup.Items.Add(this.AddStep);
            this.TestEditGroup.Items.Add(this.RemoveStep);
            this.TestEditGroup.Items.Add(this.separator1);
            this.TestEditGroup.Items.Add(this.AddActionVar);
            this.TestEditGroup.Items.Add(this.RemoveActionVar);
            this.TestEditGroup.Items.Add(this.separator2);
            this.TestEditGroup.Items.Add(this.AddCheckVar);
            this.TestEditGroup.Items.Add(this.RemoveCheckVar);
            this.TestEditGroup.Label = "Test Editing";
            this.TestEditGroup.Name = "TestEditGroup";
            // 
            // AddStep
            // 
            this.AddStep.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddStep.Image = global::ExcelAddIn.Properties.Resources.AddStep;
            this.AddStep.Label = "Add Step";
            this.AddStep.Name = "AddStep";
            this.AddStep.ShowImage = true;
            this.AddStep.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddStep_Click);
            // 
            // RemoveStep
            // 
            this.RemoveStep.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RemoveStep.Image = ((System.Drawing.Image)(resources.GetObject("RemoveStep.Image")));
            this.RemoveStep.Label = "Remove Step";
            this.RemoveStep.Name = "RemoveStep";
            this.RemoveStep.ShowImage = true;
            this.RemoveStep.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveStep_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // AddActionVar
            // 
            this.AddActionVar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddActionVar.Image = ((System.Drawing.Image)(resources.GetObject("AddActionVar.Image")));
            this.AddActionVar.Label = "Add Action Var";
            this.AddActionVar.Name = "AddActionVar";
            this.AddActionVar.ShowImage = true;
            this.AddActionVar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddActionVar_Click);
            // 
            // RemoveActionVar
            // 
            this.RemoveActionVar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RemoveActionVar.Image = ((System.Drawing.Image)(resources.GetObject("RemoveActionVar.Image")));
            this.RemoveActionVar.Label = "Remove Action Var";
            this.RemoveActionVar.Name = "RemoveActionVar";
            this.RemoveActionVar.ShowImage = true;
            this.RemoveActionVar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveActionVar_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // AddCheckVar
            // 
            this.AddCheckVar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddCheckVar.Image = ((System.Drawing.Image)(resources.GetObject("AddCheckVar.Image")));
            this.AddCheckVar.Label = "Add Check Var";
            this.AddCheckVar.Name = "AddCheckVar";
            this.AddCheckVar.ShowImage = true;
            this.AddCheckVar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddCheckVar_Click);
            // 
            // RemoveCheckVar
            // 
            this.RemoveCheckVar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RemoveCheckVar.Image = ((System.Drawing.Image)(resources.GetObject("RemoveCheckVar.Image")));
            this.RemoveCheckVar.Label = "Remove Check Var";
            this.RemoveCheckVar.Name = "RemoveCheckVar";
            this.RemoveCheckVar.ShowImage = true;
            this.RemoveCheckVar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveCheckVar_Click);
            // 
            // checks
            // 
            this.checks.Items.Add(this.testCheck);
            this.checks.Label = "Checks";
            this.checks.Name = "checks";
            // 
            // testCheck
            // 
            this.testCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.testCheck.Enabled = false;
            this.testCheck.Image = ((System.Drawing.Image)(resources.GetObject("testCheck.Image")));
            this.testCheck.Label = "Check Test";
            this.testCheck.Name = "testCheck";
            this.testCheck.ShowImage = true;
            // 
            // Outputs
            // 
            this.Outputs.Items.Add(this.toSwVTD);
            this.Outputs.Items.Add(this.testStand);
            this.Outputs.Items.Add(this.toSwVTDR);
            this.Outputs.Label = "Outputs";
            this.Outputs.Name = "Outputs";
            // 
            // toSwVTD
            // 
            this.toSwVTD.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toSwVTD.Image = global::ExcelAddIn.Properties.Resources.Tests2SwVTD;
            this.toSwVTD.Label = "To SwVTD";
            this.toSwVTD.Name = "toSwVTD";
            this.toSwVTD.ShowImage = true;
            this.toSwVTD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toSwVTD_Click);
            // 
            // testStand
            // 
            this.testStand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.testStand.Enabled = false;
            this.testStand.Image = global::ExcelAddIn.Properties.Resources._2TestStand;
            this.testStand.Label = "To TestStand";
            this.testStand.Name = "testStand";
            this.testStand.ShowImage = true;
            // 
            // toSwVTDR
            // 
            this.toSwVTDR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toSwVTDR.Enabled = false;
            this.toSwVTDR.Image = global::ExcelAddIn.Properties.Resources.TestStand2swVTDR;
            this.toSwVTDR.Label = "To SwVTDR";
            this.toSwVTDR.Name = "toSwVTDR";
            this.toSwVTDR.ShowImage = true;
            this.toSwVTDR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toSwVTDR_Click);
            // 
            // macroInfos
            // 
            this.macroInfos.Items.Add(this.help);
            this.macroInfos.Items.Add(this.macroVersion);
            this.macroInfos.Items.Add(this.UpdateDate);
            this.macroInfos.Items.Add(this.autoUpdate);
            this.macroInfos.Label = "Informations";
            this.macroInfos.Name = "macroInfos";
            // 
            // help
            // 
            this.help.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.help.Enabled = false;
            this.help.Image = global::ExcelAddIn.Properties.Resources.Help;
            this.help.Label = "Help";
            this.help.Name = "help";
            this.help.ShowImage = true;
            // 
            // macroVersion
            // 
            this.macroVersion.Label = "Version : 1.0.0.0";
            this.macroVersion.Name = "macroVersion";
            // 
            // UpdateDate
            // 
            this.UpdateDate.Label = "Update date: 25/01/2013";
            this.UpdateDate.Name = "UpdateDate";
            this.UpdateDate.Visible = false;
            // 
            // autoUpdate
            // 
            this.autoUpdate.Checked = true;
            this.autoUpdate.Enabled = false;
            this.autoUpdate.Label = "Auto update";
            this.autoUpdate.Name = "autoUpdate";
            this.autoUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.autoUpdate_Click);
            // 
            // RibbonVisual
            // 
            this.Name = "RibbonVisual";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TabValToolMgr);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonVisual_Load);
            this.TabValToolMgr.ResumeLayout(false);
            this.TabValToolMgr.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.TestEditGroup.ResumeLayout(false);
            this.TestEditGroup.PerformLayout();
            this.checks.ResumeLayout(false);
            this.checks.PerformLayout();
            this.Outputs.ResumeLayout(false);
            this.Outputs.PerformLayout();
            this.macroInfos.ResumeLayout(false);
            this.macroInfos.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabValToolMgr;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown LayoutVersion_DD;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_NewPR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton plan2Tests;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddStep;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup checks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton testCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Outputs;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton testStand;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup macroInfos;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel macroVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel UpdateDate;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox autoUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveStep;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TestEditGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddActionVar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddCheckVar;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveActionVar;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveCheckVar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toSwVTD;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toSwVTDR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton help;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addCategory;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonVisual RibbonVisual
        {
            get { return this.GetRibbon<RibbonVisual>(); }
        }
    }
}
