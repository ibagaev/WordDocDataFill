using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Office.Core;

namespace WordDocDataFill
{
    public class WordDocDataFillInstance : IWordDocDataFill
    {
        private Application app;
        private Document _templateDoc;
        private readonly String _extPath;
        private object miss = Missing.Value;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateDocPath">Path to document template WITH file name</param>
        /// <param name="extPath">Path to save result directory. WITHOUT file name</param>
        public WordDocDataFillInstance(string templateDocPath, string extPath)
        {
            _extPath = extPath;
            app = new Application();

            object templateDocPathObj = templateDocPath;
            object confirmConversions = false;
            object readOnly = true;
            object addToRecentFiles = false;
            object passwordDocument = miss;
            object passwordTemplate = miss;
            object revert = false;
            object writePasswordDocument = miss;
            object writePasswordTemplate = miss;
            object format = WdOpenFormat.wdOpenFormatAuto;
            object encoding = miss;
            object visible = true;
            object openAndRepair = true;
            object documentDirection = miss;
            object noEncodingDialog = true;
            object xMLTransform = miss;

            try
            {
                _templateDoc = app.Documents.Open(ref templateDocPathObj, ref confirmConversions, ref readOnly, ref addToRecentFiles, ref passwordDocument, ref passwordTemplate, ref revert, ref writePasswordDocument, ref writePasswordTemplate, ref format, ref encoding, ref visible, ref openAndRepair, ref documentDirection, ref noEncodingDialog, ref xMLTransform);
                _templateDoc.Activate();
            }
            catch( Exception ex)
            {
                _templateDoc.Close(false, ref miss, ref miss);
                _templateDoc = null;
                app = null;
                throw ex;
            }
        }

        public void FillDocument(Dictionary<string, string> values, string exitDocName)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 2;
            object wrap = 1;

            try
            {
                foreach (var kvp in values)
                {
                    object replKey = kvp.Key;
                    object replValue = kvp.Value;

                    app.Selection.Find.Execute(ref replKey, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replValue, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }

                object fileName = _extPath + $"\\{exitDocName}";
                object fileFormat = WdSaveFormat.wdFormatDocumentDefault;

                _templateDoc.SaveAs2(fileName, fileFormat);
            }
            catch(Exception ex)
            {
                throw ex;
            }
            finally
            {
                _templateDoc.Close(false, ref miss, ref miss);
                app.Quit(false);
            }
        }
    }
}
