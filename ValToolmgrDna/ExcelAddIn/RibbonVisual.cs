﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ValToolFunctions_2013;

namespace ExcelAddIn
{
    public partial class RibbonVisual
    {
        private void RibbonVisual_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button_NewPR_Click(object sender, RibbonControlEventArgs e)
        {
            CreateTest.NewPR(Globals.ThisAddIn.Application);
        }

        private void LayoutVersion_DD_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void autoUpdate_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void AddStep_Click(object sender, RibbonControlEventArgs e)
        {
            TestManager.AddNewStep(Globals.ThisAddIn.Application);
        }

        private void RemoveStep_Click(object sender, RibbonControlEventArgs e)
        {
            TestManager.RemoveStep(Globals.ThisAddIn.Application,EditingZone.NONE);
        }

        private void AddActionVar_Click(object sender, RibbonControlEventArgs e)
        {
            TestManager.AddVariable(Globals.ThisAddIn.Application, TEST.TABLE.TYPE.ACTION, EditingZone.NONE);
        }

        private void AddCheckVar_Click(object sender, RibbonControlEventArgs e)
        {
            TestManager.AddVariable(Globals.ThisAddIn.Application, TEST.TABLE.TYPE.CHECK, EditingZone.NONE);
        }

        private void RemoveActionVar_Click(object sender, RibbonControlEventArgs e)
        {
            TestManager.RemoveVariable(Globals.ThisAddIn.Application, TEST.TABLE.TYPE.ACTION, EditingZone.NONE);
        }

        private void RemoveCheckVar_Click(object sender, RibbonControlEventArgs e)
        {
            TestManager.RemoveVariable(Globals.ThisAddIn.Application, TEST.TABLE.TYPE.CHECK, EditingZone.NONE);
        }
    }
}