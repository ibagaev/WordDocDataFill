using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using System.Reflection;


namespace WordDocDataFill
{
    public class WordDocDataFillInstance : IWordDocDataFill
    {
        private Application app;
        private Document templateDoc;
        private String extPath;
        private object miss = Missing.Value;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateDocPath">Path to document template WITH file name</param>
        /// <param name="extPath">Path to save result directory. WITHOUT file name</param>
        public WordDocDataFillInstance(string templateDocPath, string extPath)
        {
            this.extPath = extPath;
            app = new Application();

            initializeTemplate(templateDocPath);
        }

        public void FillDocument(Dictionary<string, string> values, string exitDocName)
        {
            findAndReplace(values);
            
            saveResult(exitDocName);
        }

        private void initializeTemplate(string templateDocPath)
        {
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
                templateDoc = app.Documents.Open(
                    ref templateDocPathObj,
                    ref confirmConversions,
                    ref readOnly,
                    ref addToRecentFiles,
                    ref passwordDocument,
                    ref passwordTemplate,
                    ref revert,
                    ref writePasswordDocument,
                    ref writePasswordTemplate,
                    ref format,
                    ref encoding,
                    ref visible,
                    ref openAndRepair,
                    ref documentDirection,
                    ref noEncodingDialog,
                    ref xMLTransform);

                templateDoc.Activate();
            }
            catch (Exception ex)
            {
                templateDoc.Close(false, ref miss, ref miss);
                templateDoc = null;
                app = null;
                throw ex;
            }
        }

        private void findAndReplace(Dictionary<string, string> values)
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

                    app.Selection.Find.Execute(
                        ref replKey,
                        ref matchCase,
                        ref matchWholeWord,
                        ref matchWildCards,
                        ref matchSoundsLike,
                        ref matchAllWordForms,
                        ref forward,
                        ref wrap,
                        ref format,
                        ref replValue,
                        ref replace,
                        ref matchKashida,
                        ref matchDiacritics,
                        ref matchAlefHamza,
                        ref matchControl);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                templateDoc.Close(false, ref miss, ref miss);
                app.Quit(false);
            }
        }

        private void saveResult(string exitDocName)
        {
            try
            {
                object fileName = extPath + $"\\{exitDocName}";
                object fileFormat = WdSaveFormat.wdFormatDocumentDefault;

                templateDoc.SaveAs2(fileName, fileFormat);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                templateDoc.Close(false, ref miss, ref miss);
                app.Quit(false);
            }

        }
    }
}
