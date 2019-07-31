using System;
using Hyland.Unity;
using System.Collections.Generic;

namespace UnityScripts
{
    /// <summary>
    ///  Update keywords from PropertyBag values
    /// </summary>
    public class ArmyACC : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "US Army ACC - ISN Parse";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gParamISN = "ISN Number";
        private const string gSaveToCountry = "Capturing Country";
        private const string gSaveToTheater = "Theater";
        private const string gSaveToPower = "Power Served";
        private const string gSaveToSequence = "Sequence Number";
        private const string gSaveToDetainee = "Detainee Category";

        private const string gDocType = "ACC - Detainee File";

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

                KeywordType kwtISN = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamISN);
                string strISN = "";
                if (kwtISN != null)
                {
                    KeywordRecord keyISN = _currentDocument.KeywordRecords.Find(kwtISN);
                    if (keyISN != null)
                    {
                        Keyword kwdISN = keyISN.Keywords.Find(kwtISN);
                        if (kwdISN != null)
                            strISN = kwdISN.ToString();
                    }
                }

                if (strISN == "")
                {
                    throw new Exception(string.Format("Search keyword {0} is blank.", gSaveToISN));
                }

                string strCountry = "";
                string strTheater = "";
                string strPower = "";
                string strSequence = "";
                string strDetainee = "";

                strCountry = strISN.Substring(0, 2);
                strTheater = strISN.Substring(2, 1);
                strPower = strISN.Substring(3, 2);
                strSequence = strISN.Substring(5, 6);
                strDetainee = strISN.Substring(11, 2);

                string strTheaterResult = "";

                switch (strTheater)
                {
                    case "A":
                        strTheaterResult = "TRANSCOM";
                        break;
                    case "1":
                        strTheaterResult = "FORSCOM";
                        break;
                    case "2":
                        strTheaterResult = "AFRICOM";
                        break;
                    case "3":
                        strTheaterResult = "PACOM";
                        break;
                    case "4":
                        strTheaterResult = "EUCOM";
                        break;
                    case "5":
                        strTheaterResult = "PACOM";
                        break;
                    case "6":
                        strTheaterResult = "SOUTHCOM";
                        break;
                    case "7":
                        strTheaterResult = "SPACECOM";
                        break;
                    case "8":
                        strTheaterResult = "STRATCOM";
                        break;
                    case "9":
                        strTheaterResult = "CENTCOM";
                        break;
                    default:
                        strTheaterResult = "UNDETERMINED";
                        break;
                }

                Keyword kwdCountry = null;
                if (!String.IsNullOrEmpty(strCountry))
                {
                    KeywordType kwtCountry = app.Core.KeywordTypes.Find(gSaveToCountry);
                    if (kwtCountry != null)
                        kwdCountry = CreateKeywordHelper(kwtCountry, strCountry);
                }

                Keyword kwdTheater = null;
                if (!String.IsNullOrEmpty(strTheaterResult))
                {
                    KeywordType kwtTheater = app.Core.KeywordTypes.Find(gSaveToTheater);
                    if (kwtTheater != null)
                        kwdTheater = CreateKeywordHelper(kwtTheater, strTheaterResult);
                }

                Keyword kwdPower = null;
                if (!String.IsNullOrEmpty(strPower))
                {
                    KeywordType kwtPower = app.Core.KeywordTypes.Find(gSaveToPower);
                    if (kwtPower != null)
                        kwdPower = CreateKeywordHelper(kwtPower, strPower);
                }

                Keyword kwdSequence = null;
                if (!String.IsNullOrEmpty(strSequence))
                {
                    KeywordType kwtSequence = app.Core.KeywordTypes.Find(gSaveToSequence);
                    if (kwtSequence != null)
                        kwdSequence = CreateKeywordHelper(kwtSequence, strSequence);
                }

                Keyword kwdDetainee = null;
                if (!String.IsNullOrEmpty(strDetainee))
                {
                    KeywordType kwtDetainee = app.Core.KeywordTypes.Find(gSaveToDetainee);
                    if (kwtDetainee != null)
                        kwdDetainee = CreateKeywordHelper(kwtDetainee, strDetainee);
                }

                using (DocumentLock documentLock = _currentDocument.LockDocument())
                {
                    // Ensure lock was obtained
                    if (documentLock.Status != DocumentLockStatus.LockObtained)
                    {
                        throw new Exception("Document lock not obtained");
                    }
                    // Create keyword modifier object to hold keyword changes
                    KeywordModifier keyModifier = _currentDocument.CreateKeywordModifier();

                    keyModifier.AddKeyword(kwdCountry);
                    keyModifier.AddKeyword(kwdTheater);
                    keyModifier.AddKeyword(kwdPower);
                    keyModifier.AddKeyword(kwdSequence);
                    keyModifier.AddKeyword(kwdDetainee);

                    // Apply keyword change to the document
                    keyModifier.ApplyChanges();

                    string output = String.Format("Keyword: '{0}' Value: '{1}', {11}Keyword: '{2}' Value: '{3}', {11}Keyword: '{4}' Value: '{5}', {11}Keyword: '{6}' Value: '{7}'," +
                        "{11}Keyword: '{8}' Value: '{9}', {11}added to Document {10}.",
                        gSaveToCountry, strCountry, gSaveToTheater, strTheater, gSaveToPower, strPower, gSaveToSequence, strSequence, gSaveToDetainee, strDetainee, _currentDocument.ID, Environment.NewLine);
                    //Output the results to the OnBase Diagnostics Console
                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output);

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