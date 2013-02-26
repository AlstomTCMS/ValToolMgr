using System;
using System.Collections.Generic;
using System.Text;

namespace ValToolFunctions_2013
{
    public enum ERROR
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

    public enum Type_Var
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

    public enum SheetsNames
    {
        [StringValue("PR In")]
        PR_IN_NAME = 1,
        [StringValue("PR Out")]
        PR_OUT_NAME = 2,
        [StringValue("PR Model")]
        PR_MODEL_NAME = 3,
        [StringValue("Synthèse Model")]
        SYNTHESE_MODEL_NAME = 4,
        [StringValue("Synthèse")]
        SYNTHESE_NAME = 5,
        [StringValue("Data Validation")]
        VALID_NAME = 6,
        [StringValue("Erreurs")]
        ERROR_NAME  = 7
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

    public enum TEST_COLUMN
    {
        TYPE_ACTION = 7,
        DOUBLON_COMPARE = 12
    }

    public static class TEST
    {
        public static class TABLE
        {
            public const string ACTION="Action";
            public const string CHECK="Check";
            public const string DESCRIPTION="Desc";

            public static class PREFIX
            {
                public const string TEST = "Test_";
                public const string SCENARIO = "TS_";
                public const string ACTION= TABLE_PREFIX + TABLE.ACTION + "_";
                public const string CHECK = TABLE_PREFIX + TABLE.CHECK + "_";
                public const string DESC = TABLE_PREFIX + TABLE.DESCRIPTION + "_";
            }
            
        }

        public const string STEP_PATERN = "STEP 1";
        public const string TABLE_PREFIX = "Table_";
        

        public const string DESCRIPTION_TABLE_STYLE = "Description table";
        public const string DESCRIPTION_TABLE_STYLE_VERSION = "V.01";
    }
        
    public static class Constants
    {
        public const String SETTING_FILE_NAME  = "Application_Settings_File.MIESET";
        public const String MacroPath = "C:\\macros_alstom";
        public const String exportFolder  = "\\export\\ValToolMgr\\";

        public const String macroVersion = "A0";
        public const String macroUpdateDate  = "29/01/2013";
        public const String ERROR_NOT_IMPLEMENTED_FUNCTION = "Function not implemented";

    }
}
