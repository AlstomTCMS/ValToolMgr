using System;
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
        private static Microsoft.Office.Tools.Excel.ApplicationFactory _factory;

        private void SetExcelApplication(Excel.Application exelApp)
        {
            if (!IsAFunctionalityRunning)
            {
                _excelApplication= exelApp;
                IsAFunctionalityRunning = true;
            }
            else
            {
                throw new ExcelApplicationNotAvailableException();
            }

        }

        internal static Excel.Application ExcelApplication
        {
            get {
                if (_excelApplication == null)
                {
                    throw new ExcelApplicationMissingException();
                }
                else
                {
                    return _excelApplication;
                }
            }
        }


        private void SetFactory(Microsoft.Office.Tools.Excel.ApplicationFactory factory)
        {
            _factory = factory;
        }

        internal static Microsoft.Office.Tools.Excel.ApplicationFactory Factory
        {
            get
            {
                if (_factory == null)
                {
                    throw new ExcelApplicationMissingException();
                }
                else
                {
                    return _factory;
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

        [System.Obsolete("Use createWholeTestFormat instead", true)]
        public void NewPR(Excel.Application exelApp)
        {
            SetExcelApplication(exelApp);
            SwVTP_Creation.NewPR();
            IsAFunctionalityRunning = false;
        }

        public void NewPR(Excel.Application exelApp, Microsoft.Office.Tools.Excel.ApplicationFactory factory)
        {
            SetFactory(factory);
            SetExcelApplication(exelApp);
            SwVTP_Creation.NewPR();
            IsAFunctionalityRunning = false;
        }

        public void AddCategory(Excel.Application exelApp, EditingZone editingMode = EditingZone.NONE)
        {
            SetExcelApplication(exelApp);
            SwVTPManager.AddCategory();
            IsAFunctionalityRunning = false;
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
            SetExcelApplication(exelApp);
            TestsGenerator.FromSwVTP2Tests();
            IsAFunctionalityRunning = false;
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


        public void extractTests2SwVTD(Excel.Application exelApp)
        {
            SetExcelApplication(exelApp);
            SwVTD.GenerateSwVTD();
            IsAFunctionalityRunning = false;
        }
    }
}
