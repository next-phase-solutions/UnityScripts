using System;
using Hyland.Unity;
using System.Collections.Generic;

namespace UnityScripts
{
    /// <summary>
    ///  Update keywords from PropertyBag values
    /// </summary>
    public class DBPR_UnityScript_192 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "192 - LSCMH - SR - Mass Index KW Apply - Array Based Pbag";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gPropName1 = "LSCMH_WF_MassIndex_KeyName";
        private const string KWName1 = "Key Name";
        private const string gPropName2 = "LSCMH_WF_MassIndex_LicType";
        private const string KWName2 = "License Type";
        private const string gPropName3 = "LSCMH_WF_MassIndex_LicNum";
        private const string KWName3 = "License Number";
        private const string gPropName4 = "LSCMH_WF_MassIndex_AppNum";
        private const string KWName4 = "Application Number";
        private const string gPropName5 = "LSCMH_WF_MassIndex_FileNum";
        private const string KWName5 = "File Number";
        private const string gPropName6 = "LSCMH_WF_MassIndex_RemitNumber";
        private const string KWName6 = "Remittance Number";
        private const string gPropName7 = "LSCMH_WF_MassIndex_RemitAmount";
        private const string KWName7 = "Remittance Amount";
        private const string gPropName8 = "LSCMH_WF_MassIndex_FilingType";
        private const string KWName8 = "Filing Type";
        private const string gPropName9 = "LSCMH_WF_MassIndex_APNum";
        private const string KWName9 = "Key Name";
        private const string gPropName10 = "LSCMH_WF_MassIndex_TranCode";
        private const string KWName10 = "Tran Code";

        private List<string> lstProp1 = new List<string>();
        private List<string> lstProp2 = new List<string>();
        private List<string> lstProp3 = new List<string>();
        private List<string> lstProp4 = new List<string>();
        private List<string> lstProp5 = new List<string>();
        private List<string> lstProp6 = new List<string>();
        private List<string> lstProp7 = new List<string>();
        private List<string> lstProp8 = new List<string>();
        private List<string> lstProp9 = new List<string>();
        private List<string> lstProp10 = new List<string>();

        #endregion

        /***********************************************
         * USER/SE: PLEASE DO NOT EDIT BELOW THIS LINE *
         ***********************************************/

        #region Private Globals
        // Active workflow document
        private Document _currentDocument;
        #endregion

        #region IWorkflowScript
        /// <summary>
        /// Implementation of <see cref="IWorkflowScript.OnWorkflowScriptExecute" />.
        /// <seealso cref="IWorkflowScript" />
        /// </summary>
        /// <param name="app">Unity Application object</param>
        /// <param name="args">Workflow event arguments</param>
        //  public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args)
        //public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args = null)
        public void OnWorkflowScriptExecute(Hyland.Unity.Application app, Hyland.Unity.WorkflowEventArgs args)
        {
            try
            {
                // Initialize global settings
                IntializeScript(ref app, ref args);

                // Create keyword modifier object to hold keyword changes
                KeywordModifier keyModifier = _currentDocument.CreateKeywordModifier();

                Keyword kwdKeyName1 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp1))
                {
                }

                KeywordType kwtKeyName1 = app.Core.KeywordTypes.Find(KWName1);
                foreach (string props in lstProp1)
                {
                    kwdKeyName1 = CreateKeywordHelper(kwtKeyName1, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName1, props));
                    if (kwdKeyName1 != null) keyModifier.AddKeyword(kwdKeyName1);
                }

                Keyword kwdKeyName2 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp2))
                {
                }

                KeywordType kwtKeyName2 = app.Core.KeywordTypes.Find(KWName2);
                foreach (string props in lstProp2)
                {
                    kwdKeyName2 = CreateKeywordHelper(kwtKeyName2, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName2, props));
                    if (kwdKeyName2 != null) keyModifier.AddKeyword(kwdKeyName2);
                }

                Keyword kwdKeyName3 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp3))
                {
                }

                KeywordType kwtKeyName3 = app.Core.KeywordTypes.Find(KWName3);
                foreach (string props in lstProp3)
                {
                    kwdKeyName3 = CreateKeywordHelper(kwtKeyName3, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName3, props));
                    if (kwdKeyName3 != null) keyModifier.AddKeyword(kwdKeyName3);
                }

                Keyword kwdKeyName4 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp4))
                {
                }

                KeywordType kwtKeyName4 = app.Core.KeywordTypes.Find(KWName4);
                foreach (string props in lstProp4)
                {
                    kwdKeyName4 = CreateKeywordHelper(kwtKeyName4, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName4, props));
                    if (kwdKeyName4 != null) keyModifier.AddKeyword(kwdKeyName4);
                }

                Keyword kwdKeyName5 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp5))
                {
                }

                KeywordType kwtKeyName5 = app.Core.KeywordTypes.Find(KWName5);
                foreach (string props in lstProp5)
                {
                    kwdKeyName5 = CreateKeywordHelper(kwtKeyName5, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName5, props));
                    if (kwdKeyName5 != null) keyModifier.AddKeyword(kwdKeyName5);
                }

                Keyword kwdKeyName6 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp6))
                {
                }

                KeywordType kwtKeyName6 = app.Core.KeywordTypes.Find(KWName6);
                foreach (string props in lstProp6)
                {
                    kwdKeyName6 = CreateKeywordHelper(kwtKeyName6, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName6, props));
                    if (kwdKeyName6 != null) keyModifier.AddKeyword(kwdKeyName6);
                }

                Keyword kwdKeyName7 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp7))
                {
                }

                KeywordType kwtKeyName7 = app.Core.KeywordTypes.Find(KWName7);
                foreach (string props in lstProp7)
                {
                    kwdKeyName7 = CreateKeywordHelper(kwtKeyName7, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName7, props));
                    if (kwdKeyName7 != null) keyModifier.AddKeyword(kwdKeyName7);
                }

                Keyword kwdKeyName8 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp8))
                {
                }

                KeywordType kwtKeyName8 = app.Core.KeywordTypes.Find(KWName8);
                foreach (string props in lstProp8)
                {
                    kwdKeyName8 = CreateKeywordHelper(kwtKeyName8, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName8, props));
                    if (kwdKeyName8 != null) keyModifier.AddKeyword(kwdKeyName8);
                }

                Keyword kwdKeyName9 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp9))
                {
                }

                KeywordType kwtKeyName9 = app.Core.KeywordTypes.Find(KWName9);
                foreach (string props in lstProp9)
                {
                    kwdKeyName9 = CreateKeywordHelper(kwtKeyName9, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName9, props));
                    if (kwdKeyName9 != null) keyModifier.AddKeyword(kwdKeyName9);
                }

                Keyword kwdKeyName10 = null;
                if (!args.SessionPropertyBag.TryGetValue(gPropName1, out string[] lstProp10))
                {
                }

                KeywordType kwtKeyName10 = app.Core.KeywordTypes.Find(KWName10);
                foreach (string props in lstProp10)
                {
                    kwdKeyName10 = CreateKeywordHelper(kwtKeyName10, props);
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Property Bag [{0}] value is {1}.", gPropName10, props));
                    if (kwdKeyName10 != null) keyModifier.AddKeyword(kwdKeyName10);
                }

                using (DocumentLock documentLock = _currentDocument.LockDocument())
                {
                    // Ensure lock was obtained
                    if (documentLock.Status != DocumentLockStatus.LockObtained)
                    {
                        throw new Exception("Document lock not obtained");
                    }

                    // Apply keyword change to the document
                    keyModifier.ApplyChanges();

                    documentLock.Release();

                }

            }
            catch (Exception ex)
            {
                // Handle exceptions and log to Diagnostics Console and document history
                HandleException(ex, ref app, ref args);
            }
            finally
            {
                // Log script execution end
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info,
                    string.Format("End Script - [{0}]", ScriptName));
            }
        }
        #endregion

        #region Helper Functions
        /// <summary>
        /// Initialize global settings
        /// </summary>
        /// <param name="app">Unity Application object</param>
        /// <param name="args">Workflow event arguments</param>
        private void IntializeScript(ref Application app, ref WorkflowEventArgs args)
        {
            // Set the specified diagnostics level
            app.Diagnostics.Level = DiagLevel;

            // Log script execution start
            app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info,
                string.Format("{0} - Start Script - [{1}]", DateTime.Now.ToString(DateTimeFormat), ScriptName));

            // Capture active document as global
            _currentDocument = args.Document;
            //_currentDocument = app.Core.GetDocumentByID(TestDocId);
            // If an error was stored in the property bag from a previous execution, clear it
            if (args.SessionPropertyBag.ContainsKey(ErrorMessageProperty)) args.SessionPropertyBag.Remove(ErrorMessageProperty);

            // Set ScriptResult = true for workflow rules (will become false if an error is caught)
            args.ScriptResult = true;
        }

        /// <summary>
        /// Handle exceptions and log to Diagnostics Console and document history
        /// </summary>
        /// <param name="ex">Exception</param>
        /// <param name="app">Unity Application object</param>
        /// <param name="args">Workflow event arguments</param>
        /// <param name="otherDocument">Document on which to update history if not active workflow document</param>
        private void HandleException(Exception ex, ref Application app, ref WorkflowEventArgs args, Document otherDocument = null)
        {
            var history = app.Core.LogManagement;
            bool isInner = false;

            // Cycle through all inner exceptions
            while (ex != null)
            {
                // Construct error text to store to workflow property
                string propertyError = string.Format("{0}{1}: {2}",
                    isInner ? "Inner " : "",
                    ex.GetType().Name,
                    ex.Message);

                // Construct error text to store to document history
                string historyError = propertyError.Replace(ex.GetType().Name, "Unity Script Error");

                // Construct error text to log to diagnostics console
                string diagnosticsError = string.Format("{0} - ***ERROR***{1}{2}{1}{1}Stack Trace:{1}{3}",
                    DateTime.Now.ToString(DateTimeFormat),
                    Environment.NewLine,
                    propertyError,
                    ex.StackTrace);

                // Add error message to document history (on current or specified document)
                var document = otherDocument ?? _currentDocument;
                if (document != null && WriteErrorsToHistory) history.CreateDocumentHistoryItem(document, historyError);

                // Write error message to Diagnostcs Consonle
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, diagnosticsError);

                // Store the original (inner) exception message to error workflow property
                if (ex.InnerException == null) args.SessionPropertyBag.Set(ErrorMessageProperty, ex.Message);

                // Move on to next inner exception
                ex = ex.InnerException;
                isInner = true;
            }

            // Set ScriptResult = false for workflow rules
            args.ScriptResult = false;
        }
        private Keyword CreateKeywordHelper(KeywordType Keytype, string Value)
        {
            Keyword key = null;
            switch (Keytype.DataType)
            {
                case KeywordDataType.Currency:
                case KeywordDataType.Numeric20:
                    decimal decVal = decimal.Parse(Value);
                    key = Keytype.CreateKeyword(decVal);
                    break;
                case KeywordDataType.Date:
                case KeywordDataType.DateTime:
                    DateTime dateVal = DateTime.Parse(Value);
                    key = Keytype.CreateKeyword(dateVal);
                    break;
                case KeywordDataType.FloatingPoint:
                    double dblVal = double.Parse(Value);
                    key = Keytype.CreateKeyword(dblVal);
                    break;
                case KeywordDataType.Numeric9:
                    long lngVal = long.Parse(Value);
                    key = Keytype.CreateKeyword(lngVal);
                    break;
                default:
                    key = Keytype.CreateKeyword(Value);
                    break;
            }
            return key;
        }

        #endregion
    }
}