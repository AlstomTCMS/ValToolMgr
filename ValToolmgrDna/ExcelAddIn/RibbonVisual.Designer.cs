﻿namespace ExcelAddIn
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonVisual));
            this.ValToolMgrTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.LayoutVersion_DD = this.Factory.CreateRibbonDropDown();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button_NewPR = this.Factory.CreateRibbonButton();
            this.plan2Tests = this.Factory.CreateRibbonButton();
            this.AddStep = this.Factory.CreateRibbonButton();
            this.checks = this.Factory.CreateRibbonGroup();
            this.testCheck = this.Factory.CreateRibbonButton();
            this.Outputs = this.Factory.CreateRibbonGroup();
            this.testStand = this.Factory.CreateRibbonButton();
            this.macroInfos = this.Factory.CreateRibbonGroup();
            this.macroVersion = this.Factory.CreateRibbonLabel();
            this.UpdateDate = this.Factory.CreateRibbonLabel();
            this.autoUpdate = this.Factory.CreateRibbonCheckBox();
            this.ValToolMgrTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.checks.SuspendLayout();
            this.Outputs.SuspendLayout();
            this.macroInfos.SuspendLayout();
            // 
            // ValToolMgrTab
            // 
            this.ValToolMgrTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ValToolMgrTab.Groups.Add(this.group1);
            this.ValToolMgrTab.Groups.Add(this.group2);
            this.ValToolMgrTab.Groups.Add(this.checks);
            this.ValToolMgrTab.Groups.Add(this.Outputs);
            this.ValToolMgrTab.Groups.Add(this.macroInfos);
            this.ValToolMgrTab.Label = "Val tool Mgr";
            this.ValToolMgrTab.Name = "ValToolMgrTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.LayoutVersion_DD);
            this.group1.Label = "Layout Version";
            this.group1.Name = "group1";
            // 
            // label1
            // 
            this.label1.Label = "Choose a version";
            this.label1.Name = "label1";
            // 
            // LayoutVersion_DD
            // 
            this.LayoutVersion_DD.Enabled = false;
            ribbonDropDownItemImpl1.Label = "2013";
            ribbonDropDownItemImpl2.Label = "2012";
            this.LayoutVersion_DD.Items.Add(ribbonDropDownItemImpl1);
            this.LayoutVersion_DD.Items.Add(ribbonDropDownItemImpl2);
            this.LayoutVersion_DD.Label = " ";
            this.LayoutVersion_DD.Name = "LayoutVersion_DD";
            this.LayoutVersion_DD.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LayoutVersion_DD_SelectionChanged);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button_NewPR);
            this.group2.Items.Add(this.plan2Tests);
            this.group2.Items.Add(this.AddStep);
            this.group2.Label = "Editing";
            this.group2.Name = "group2";
            // 
            // button_NewPR
            // 
            this.button_NewPR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_NewPR.Image = ((System.Drawing.Image)(resources.GetObject("button_NewPR.Image")));
            this.button_NewPR.Label = "New PR";
            this.button_NewPR.Name = "button_NewPR";
            this.button_NewPR.ShowImage = true;
            this.button_NewPR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_NewPR_Click);
            // 
            // plan2Tests
            // 
            this.plan2Tests.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.plan2Tests.Enabled = false;
            this.plan2Tests.Image = ((System.Drawing.Image)(resources.GetObject("plan2Tests.Image")));
            this.plan2Tests.Label = "Plan to Tests";
            this.plan2Tests.Name = "plan2Tests";
            this.plan2Tests.ShowImage = true;
            // 
            // AddStep
            // 
            this.AddStep.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddStep.Enabled = false;
            this.AddStep.Image = ((System.Drawing.Image)(resources.GetObject("AddStep.Image")));
            this.AddStep.Label = "Add Step";
            this.AddStep.Name = "AddStep";
            this.AddStep.ShowImage = true;
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
            this.Outputs.Items.Add(this.testStand);
            this.Outputs.Label = "Outputs";
            this.Outputs.Name = "Outputs";
            // 
            // testStand
            // 
            this.testStand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.testStand.Enabled = false;
            this.testStand.Image = ((System.Drawing.Image)(resources.GetObject("testStand.Image")));
            this.testStand.Label = "To TestStand";
            this.testStand.Name = "testStand";
            this.testStand.ShowImage = true;
            // 
            // macroInfos
            // 
            this.macroInfos.Items.Add(this.macroVersion);
            this.macroInfos.Items.Add(this.UpdateDate);
            this.macroInfos.Items.Add(this.autoUpdate);
            this.macroInfos.Label = "Informations";
            this.macroInfos.Name = "macroInfos";
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
            // 
            // autoUpdate
            // 
            this.autoUpdate.Checked = true;
            this.autoUpdate.Label = "Auto update";
            this.autoUpdate.Name = "autoUpdate";
            this.autoUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.autoUpdate_Click);
            // 
            // RibbonVisual
            // 
            this.Name = "RibbonVisual";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ValToolMgrTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonVisual_Load);
            this.ValToolMgrTab.ResumeLayout(false);
            this.ValToolMgrTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.checks.ResumeLayout(false);
            this.checks.PerformLayout();
            this.Outputs.ResumeLayout(false);
            this.Outputs.PerformLayout();
            this.macroInfos.ResumeLayout(false);
            this.macroInfos.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ValToolMgrTab;
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
    }

    partial class ThisRibbonCollection
    {
        internal RibbonVisual RibbonVisual
        {
            get { return this.GetRibbon<RibbonVisual>(); }
        }
    }
}
