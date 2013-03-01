using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ValToolFunctionsStub
{
    /// <summary>
    /// This interface define the use case between the Validation Tool Manager 
    /// and functionnal librairies which implement this interface.
    /// </summary>
    public interface RibbonHandlerInterface
    {
        //internal Application ExcelApplication { get; }

        ///// <summary>
        ///// This property is to use as a singleton to prevent simultaneous multiple user actions on a sheet.
        ///// </summary>
        //private bool IsAFunctionalityRunning { get; set; }


        /// <summary>
        /// Create a new PR from scratch
        /// </summary>
        void NewPR(Application exelApp);

        #region SwVTP

        /// <summary>
        /// Add a category at the end of SwVTP
        /// </summary>
        void AddCategory(Application exelApp, EditingZone editingMode = EditingZone.NONE);

        /// <summary>
        /// remove the category at the end of SwVTP
        /// </summary>
        void RemoveCategory(Application exelApp, EditingZone editingMode = EditingZone.NONE);

        /// <summary>
        /// Add a test at the end of the zone if category selected, the line after if a test selected.
        /// </summary>
        void AddTest(Application exelApp, EditingZone editingMode = EditingZone.NONE);

        /// <summary>
        /// Remove the test at the end of the zone if category selected, the current line if a test selected.
        /// </summary>
        void RemoveTest(Application exelApp, EditingZone editingMode = EditingZone.NONE);

        /// <summary>
        /// Cut the selected test and keep it in memory.
        /// </summary>
        void CutTest(Application exelApp, EditingZone editingMode = EditingZone.NONE);

        /// <summary>
        /// Paste the test in memory at the end of the zone if category selected, the line after if a test selected.
        /// </summary>
        void PasteTest(Application exelApp, EditingZone editingMode = EditingZone.NONE);

        #endregion

        /// <summary>
        /// Create tests sheets from the SwVTP or "Synthèse" sheet
        /// </summary>
        void PlanToTests(Application exelApp);

        #region Test's actions

        void AddStep(Application exelApp, EditingZone editingMode = EditingZone.NONE);
        void RemoveStep(Application exelApp, EditingZone editingMode = EditingZone.NONE);
        void AddActionVar(Application exelApp, EditingZone editingMode = EditingZone.NONE);
        void RemoveActionVar(Application exelApp, EditingZone editingMode = EditingZone.NONE);
        void AddCheckVar(Application exelApp, EditingZone editingMode = EditingZone.NONE);
        void RemoveCheckVar(Application exelApp, EditingZone editingMode = EditingZone.NONE);

        #endregion
    }
}
