using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel =Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Reflection;
using System.IO;

namespace ValToolFunctions_2013
{
    internal class General
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
        internal static Worksheet InitSheet(string sheetName, Boolean eraseContent = false, Boolean visible = true, Boolean sheetAlreadyExist = false , Array titles  = null)
        {
            Sheets sheets = (Sheets)RibbonHandler.ExcelApplication.ActiveWorkbook.Worksheets;

            //Si la feuille n'existe pas, on l'ajoute
            if (! WsExist(sheetName)){
                sheets.Add(After: sheets[sheets.Count]).Name = sheetName;
            }else{
                sheetAlreadyExist = true;
            }
            Worksheet newSheet = sheets[sheetName];

            try
            {
                // On efface le contenu de la feuille
                //Sheets(sheetName).Cells.ClearContents
                if (eraseContent)
                {
                    newSheet.Cells.ClearContents();
                }

                //On ajoute les titres s'il y en a
                if (titles != null)
                {
                    Range endTitle = sheets[sheetName].Cells(1, titles.Length + 1);
                    Range titleRange = sheets[sheetName].Range("A1", endTitle);
                    titleRange.Value = titles;
                    string tableLiens = "Tableau" + sheetName;

                    newSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, titleRange, XlYesNoGuess.xlYes).Name = tableLiens;
                    newSheet.ListObjects[tableLiens].TableStyle = "tableau de test";
                }
                //enlève l'affichage grille
                newSheet.Activate();
                RibbonHandler.ExcelApplication.ActiveWindow.DisplayGridlines = false;
            }
            catch { }
           
            if(! visible){
                newSheet.Visible = XlSheetVisibility.xlSheetHidden;
            }else{
                newSheet.Visible = XlSheetVisibility.xlSheetVisible;
            }
    
            //feuille renvoyée
            return newSheet;
    
        }

        /// <summary>
        /// Dit si une feuille existe dans le fichier
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        internal static Boolean WsExist(string name)
        {            
            //Nous dit si la feuille mis en paramètre existe
            try{
                int exist = RibbonHandler.ExcelApplication.ActiveWorkbook.Worksheets[name].index;
                if (exist > 0){
                    return true;
                }
            }
            catch { }
            return false; 
        }



        ////Fonction à appeler depuis toute macro appelée par un bouton de barre de macro externe
        ////return vrai si il y a un fichier d//ouvert
        internal static Boolean HasActiveBook(Boolean displayMsg = true)
        {
            Boolean hasActiveBook = false;
            try{
                //Si on a un nouveau classeur vide
                //If ActiveWorkbook.Name Like "Classeur*" Or ActiveWorkbook.Name Like "Book*" Then
                    //GoTo NoActiveWorkBook
                //End If
                if (RibbonHandler.ExcelApplication.Workbooks.Count > 0)
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

        /// <summary>
        /// Détecter si c'est bien un onglet de test au bon format
        /// </summary>
        /// <param name="displayMsg">Sortir avec message sinon. Vrai par défaut</param>
        /// <returns></returns>
        internal static Boolean isActivesheet_a_PR_Test( Boolean displayMsg  = true)
        {
            Boolean isActivesheet_a_PR_Test;

            if (Regex.IsMatch(RibbonHandler.ExcelApplication.ActiveSheet.name, TEST.TABLE.PREFIX.TEST + "*"))
            {
                isActivesheet_a_PR_Test = true;
            }else
            {
                isActivesheet_a_PR_Test = false;
                if (displayMsg){
                    MessageBox.Show("This sheet is not a PR test. You cannot use this function on this sheet.");
                }
            }
            return isActivesheet_a_PR_Test;
        }

        internal static string getTestNumber()
        {
            string getTestNumber="";
            string shName = "";
            try
            {
                shName = RibbonHandler.ExcelApplication.ActiveSheet.Name;
                getTestNumber = Regex.Split(shName, "_")[1];
            }
            catch { }
            return getTestNumber;
        }

        internal static string getTestNumber(string sheetName)
        {
            string getTestNumber = "";
            try
            {
                getTestNumber = Regex.Split(sheetName, "_")[1];
            }
            catch { }
            return getTestNumber;
        }

        internal static string GetFromResources(string resourceName)
        {
            Assembly assem = RibbonHandler.ExcelApplication.GetType().Assembly;
            using (Stream stream = assem.GetManifestResourceStream(resourceName))
            {
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
