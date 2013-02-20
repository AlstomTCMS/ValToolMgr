Attribute VB_Name = "Constantes"
Public Const SETTING_FILE_NAME As String = "Application_Settings_File.MIESET"
Public Const MacroPath As String = "C:\macros_alstom"
Public Const serverPath As String = "\\dom2.ad.sys\dfsbor1root\BOR1_FLO\DEP_Etudes\Tsysteme\Affaires\PRIMA EL2\Ctrl-cmd\Banc de Test\13_Macros"
Public Const exportFolder As String = "\export\Functions_PrimaELII_2-A0\"

Public Const macroVersion As String = "A9"
Public Const refVersion As String = "A3"
Public Const macroUpdateDate As String = "13/02/2013"

Public Const PR_IN_NAME As String = "PR In"
Public Const PR_OUT_NAME As String = "PR Out"
Public Const PR_MODEL_NAME As String = "PR Model"
Public Const SYNTHESE_MODEL_NAME As String = "Synthèse Model"
Public Const SYNTHESE_NAME As String = "Synthèse"
Public Const VALID_NAME As String = "Data Validation"
Public Const ERROR_NAME As String = "Erreurs"

Public Const ERROR_TYPE_PRIMA_VEHICULS As String = "Seuls les véhicules 1 et 2 sont permis pour PRIMA."
Public Const ERROR_TYPE_DOUBLON As String = "{0} est en doublon."
Public Const ERROR_TYPE_COLUMNS_EMPTY As String = "Les colonnes {0} ne sont pas entièrement remplies."
Public Const ERROR_TYPE_NOTESTSHEET As String = "La feuille de test n'existe pas."
Public Const ERROR_TYPE_ORDER As String = "L'ordre des types de variables (ACc, AEn, CCc, CEn) est non respecté."
Public Const ERROR_TYPE_TARGET As String = "Chemin {0} incorrect."

Public Const TEST_COLUMN_TYPE_ACTION As Integer = 7
Public Const TEST_COLUMN_DOUBLON_COMPARE As Integer = 12

Public Const TYPE_VAR_ACTION_EMB As String = "ACc"
Public Const TYPE_VAR_ACTION_ENV As String = "AEn"
Public Const TYPE_VAR_CHECK_EMB As String = "CCc"
Public Const TYPE_VAR_CHECK_ENV As String = "CEn"
Public Const TYPE_VAR_PGM As String = "PGM"



Public Const Anc_Num_Test = "A", Anc_Des_Etape = "B", Anc_Com_Test = "C", _
Anc_Num_Etape = "D", Anc_Com_Etape = "E", Anc_Com_act = "F", Anc_Com_chk = "G", Anc_Pause = "H", _
Anc_Type_Var = "I", Anc_Vehicule = "J", Anc_Variable = "K", Anc_Chemin = "L", Anc_Valeur = "M"

Public Const Nvo_Num_Test = "B", Nvo_Exigence_Associee = "C", Nvo_Description_Test = "D", _
Nvo_Etape = "E", Nvo_Commentaires_Etapes = "F", Nvo_Description_Actions = "G", Nvo_Description_Verification = "H"

Public Const Nvo_Test_Num_Etape = "A"
