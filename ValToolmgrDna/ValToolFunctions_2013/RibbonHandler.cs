﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ValToolFunctionsStub;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ValToolFunctions_2013
{
    public class RibbonHandler : RibbonHandlerInterface
    {
        private static bool _isAFunctionalityRunning;
        private static Excel.Application _excelApplication;

        private void SetExcelApplication(Excel.Application exelApp)
        {
            if (!IsAFunctionalityRunning)
            {
                _excelApplication= exelApp;
            }
            else
            {
                throw new ExcelApplicationNotAvailableException();
            }

        }

        internal static Excel.Application ExcelApplication
        {
            get {
                if (!IsAFunctionalityRunning)
                {
                    if (_excelApplication == null)
                    {
                        throw new ExcelApplicationMissingException();
                    }
                    else
                    {
                        return _excelApplication;
                    }
                }
                else
                {
                    //System.Windows.Forms.throw new NotImplementedException();
                    throw new ExcelApplicationNotAvailableException();
                }
            }
        }

        private static bool IsAFunctionalityRunning
        {
            get
            {
                return _isAFunctionalityRunning;
            }
            set
            {
                _isAFunctionalityRunning = value;
            }
        }

        #region SwVTP

        public void NewPR(Excel.Application exelApp)
        {
            SetExcelApplication(exelApp);
            CreateTest.NewPR();
            IsAFunctionalityRunning = false;
        }

        public void AddCategory(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            throw new NotImplementedException();
        }

        public void RemoveCategory(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            throw new NotImplementedException();
        }

        public void AddTest(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            throw new NotImplementedException();
        }

        public void RemoveTest(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            throw new NotImplementedException();
        }

        public void CutTest(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            throw new NotImplementedException();
        }

        public void PasteTest(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            throw new NotImplementedException();
        }

        #endregion

        public void PlanToTests(Excel.Application exelApp)
        {
            throw new NotImplementedException();
        }


        #region Test

        public void AddStep(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            SetExcelApplication(exelApp);
            TestManager.AddNewStep(editingMode);
            IsAFunctionalityRunning = false;
        }

        public void RemoveStep(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            SetExcelApplication(exelApp);
            TestManager.RemoveStep(editingMode);
            IsAFunctionalityRunning = false;
        }

        public void AddActionVar(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            SetExcelApplication(exelApp);
            TestManager.AddVariable(TEST.TABLE.TYPE.ACTION, editingMode);
            IsAFunctionalityRunning = false;
        }

        public void RemoveActionVar(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            SetExcelApplication(exelApp);
            TestManager.RemoveVariable(TEST.TABLE.TYPE.ACTION, editingMode);
            IsAFunctionalityRunning = false;
        }

        public void AddCheckVar(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            SetExcelApplication(exelApp);
            TestManager.AddVariable(TEST.TABLE.TYPE.CHECK, editingMode);
            IsAFunctionalityRunning = false;
        }

        public void RemoveCheckVar(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            SetExcelApplication(exelApp);
            TestManager.RemoveVariable(TEST.TABLE.TYPE.CHECK, editingMode);
            IsAFunctionalityRunning = false;
        }

        #endregion
    }
}