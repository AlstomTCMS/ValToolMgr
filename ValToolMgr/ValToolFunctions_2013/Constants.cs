﻿using System;
using System.Collections.Generic;
using System.Text;
using ValToolFunctionsStub;

namespace ValToolFunctions_2013
{
    public enum ERROR:int
    {
        [StringValue("Function not implemented")]
        NOT_IMPLEMENTED_FUNCTION = 1,
        [StringValue("{0} est en doublon.")]
        DOUBLON = 2,
        [StringValue("Les colonnes {0} ne sont pas entièrement remplies.")]
        EMPTY = 3,
        [StringValue("L'ordre des types de variables (ACc, AEn, CCc, CEn) est non respecté.")]
        ORDER = 4,
        [StringValue("Chemin {0} incorrect.")]
        TARGET = 5        
    }

    public enum Type_Var:int
    {
        [StringValue("ACc")]
        TYPE_VAR_ACTION_EMB = 1,
        [StringValue("AEn")]
        TYPE_VAR_ACTION_ENV = 2,
        [StringValue("CCc")]
        TYPE_VAR_CHECK_EMB = 3,
        [StringValue("CEn")]
        TYPE_VAR_CHECK_ENV = 4,
        [StringValue("PGM")]
        TYPE_VAR_PGM  = 5
    }

    public enum SheetsNames:int
    {
        [StringValue("Endpaper_PR")]
        ENDPAPER_PR,
        [StringValue("Endpaper_PV")]
        ENDPAPER_PV,
        [StringValue("Evol")]
        EVOLUTION,
        //[StringValue("Bench Conf")]
        //BENCH_CONF,
        [StringValue("SwVTP")]
        SW_VTP,
        [StringValue("SwVTD")]
        SW_VTD,
        [StringValue("SwVTDR")]
        SW_VTDR,
        [StringValue("Data Validation")]
        VALID,
        [StringValue("Errors")]
        ERROR
    }

    public enum SwVTx_Columns:int
    {
        [StringValue("Category")]
        CATEGORY = 1,
        [StringValue("Test")]
        TEST = 2,
        [StringValue("Bench Conf")]
        BENCH_CONF = 3,
        [StringValue("Associated requirement(s)")]
        REQUIREMENT = 4,
        [StringValue("Step")]
        STEP = 5,
        [StringValue("Title")]
        TITLE = 6,
        [StringValue("Action")]
        ACTION = 7,
        [StringValue("Check")]
        CHECK = 8,
        [StringValue("Status")]
        STATUS = 9,
        [StringValue("Justification")]
        JUSTIFICATION = 10,
        [StringValue("Comment")]
        R_COMMENT = 11
    }

    public enum LAYOUT
    {
        [StringValue("None")]
        NONE = 0,
        [StringValue("2012")]
        L_2012 = 1,
        [StringValue("2013")]
        L_2013 = 2
    }

    public enum TEST_COLUMN:int
    {
        TYPE_ACTION = 7,
        DOUBLON_COMPARE = 12
    }

    public static class TEST
    {
        public static class TABLE
        {
            public enum TYPE:int
            {
                [StringValue("Action")]
                ACTION,
                [StringValue("Check")]
                CHECK,
                [StringValue("Desc")]
                DESC
            }

            public static class PREFIX
            {
                /// <summary>
                /// Test_
                /// </summary>
                public const string TEST = "Test_";

                /// <summary>
                /// TS_
                /// </summary>
                public const string SCENARIO = "TS_";

                /// <summary>
                /// Table_Action_
                /// </summary>
                public const string ACTION = TABLE_PREFIX + "Action_";

                /// <summary>
                /// Table_Check_
                /// </summary>
                public const string CHECK = TABLE_PREFIX + "Check_";

                /// <summary>
                /// Table_Desc_
                /// </summary>
                public const string DESC = TABLE_PREFIX + "Desc_";
            }
            
        }

        /// <summary>
        /// STEP 1
        /// </summary>
        public const string STEP_PATERN = "STEP 1";

        /// <summary>
        /// Table_
        /// </summary>
        public const string TABLE_PREFIX = "Table_";
        
        /// <summary>
        /// "Description table"
        /// </summary>
        public const string DESCRIPTION_TABLE_STYLE = "Description table";
        /// <summary>
        /// V.02
        /// </summary>
        public const string DESCRIPTION_TABLE_STYLE_VERSION = "V.02";
    }
        
    public static class Constants
    {
        /// <summary>
        /// Function not implemented
        /// </summary>
        public const String ERROR_NOT_IMPLEMENTED_FUNCTION = "Function not implemented";

    }
}