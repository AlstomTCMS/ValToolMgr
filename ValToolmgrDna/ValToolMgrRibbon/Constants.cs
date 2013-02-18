using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrRibbon
{
    public static class Constants
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
        

        public const String SETTING_FILE_NAME  = "Application_Settings_File.MIESET";
        public const String MacroPath = "C:\\macros_alstom";
        public const String exportFolder  = "\\export\\ValToolMgr\\";

        public const String macroVersion = "A0";
        public const String macroUpdateDate  = "29/01/2013";
    }
}
