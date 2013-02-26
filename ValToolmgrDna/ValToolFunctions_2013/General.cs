using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel =Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ValToolFunctions_2013
{
    class General
    {
        /// <summary>
        /// Init a sheet by his name
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="eraseContent"></param>
        /// <param name="visible"></param>
        /// <param name="sheetAlreadyExist"></param>
        /// <param name="titles"></param>
        /// <returns></returns>
        public static Excel.Worksheet InitSheet(string sheetName, Boolean eraseContent = false, Boolean visible = true, Boolean sheetAlreadyExist = false , Array titles  = null)
        {
            Excel.Sheets sheets = (Excel.Sheets)ExcelApplication.getInstance().ActiveWorkbook.Worksheets;

            //Si la feuille n'existe pas, on l'ajoute
            if (! WsExist(sheetName)){
                sheets.Add();
                sheets[sheets.Count].name = sheetName;
            }else{
                sheetAlreadyExist = true;
            }
    
            //On Error Resume Next
            // On efface le contenu de la feuille
            //Sheets(sheetName).Cells.ClearContents
            if (eraseContent)
            {
                sheets[sheetName].Cells.ClearContents();
            }
    
            //On ajoute les titres s//il y en a
            if (titles !=null){
                Excel.Range endTitle = sheets[sheetName].Cells(1, titles.Length + 1);
                Excel.Range titleRange = sheets[sheetName].Range("A1", endTitle);
                titleRange.Value=titles;
                string tableLiens = "Tableau" + sheetName;

                sheets[sheetName].ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, titleRange, Excel.XlYesNoGuess.xlYes).name = tableLiens;
                sheets[sheetName].ListObjects(tableLiens).TableStyle = "tableau de test";
            
                //enlève l'affichage grille
                sheets[sheetName].Activate();
                ExcelApplication.getInstance().ActiveWindow.DisplayGridlines = false;

            }
            //On Error GoTo 0
           
            if(! visible){
                sheets[sheetName].visible = Excel.XlSheetVisibility.xlSheetHidden;
            }else{
                sheets[sheetName].visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
    
            //feuille renvoyée
            return sheets[sheetName];
    
        }

        /// <summary>
        /// Dit si une feuille existe dans le fichier
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Boolean WsExist(string name)
        {            
            //Nous dit si la feuille mis en paramètre existe
            try{
                int exist = ExcelApplication.getInstance().ActiveWorkbook.Worksheets[name].index;
                if (exist > 0){
                    return true;
                }
            }
            catch { }
            return false; 
        }



        ////Fonction à appeler depuis toute macro appelée par un bouton de barre de macro externe
        ////return vrai si il y a un fichier d//ouvert
        public static Boolean HasActiveBook(Boolean displayMsg = true)
        {
            Boolean hasActiveBook = false;
            try{
                //Si on a un nouveau classeur vide
                //If ActiveWorkbook.Name Like "Classeur*" Or ActiveWorkbook.Name Like "Book*" Then
                    //GoTo NoActiveWorkBook
                //End If
                if (ExcelApplication.getInstance().Workbooks.Count > 0)
                {
                    hasActiveBook = true;
                }
            }    
            catch{
                hasActiveBook = false;
                if (displayMsg){
                    MessageBox.Show("Alerte", "Please open a PR file to use this feature.");
                }
            }
            return hasActiveBook;
        }


        ////Réecri un String avec des parametres entre crochet {} remplacés par la liste de paramètres mis en argument
        //Public Function StringFormat(ByVal forFormat As String, ParamArray params() As Variant) As String
        //    Dim i As Integer
        //    Dim formatted As String

        //    formatted = forFormat
        //    For i = LBound(params()) To UBound(params())
        //        formatted = Replace(formatted, "{" & CStr(i) & "}", CStr(params(i)))
        //    Next
        //    StringFormat = formatted
        //End Function

    }
}
