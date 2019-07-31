using System;
using System.Text;
using System.Data;
using System.Data.Odbc;
using Hyland.Unity;
using Application = Hyland.Unity.Application;


namespace UnityScripts
{
    /// <summary>
    /// Description
    /// </summary>
    public class DBPR_UnityScript_408 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "408 - REG - Get Complainant";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gSaveToBoard= "Board";
        private const string gSaveToComplainant = "Key Name";
        private string propLicType = "";
        private string propRegion = "";
        private const string exREASON = "Exception Reason";

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
        public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args)
        //public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args = null)
        {
            try
            {
                // Initialize global settings
                IntializeScript(ref app, ref args);

                args.PropertyBag.TryGetValue("pbLicType", out propLicType);
                propLicType = propLicType.Substring(0, 2);
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("LicType trimmed: {0}", propLicType));

                string strBoard = "";

                switch (propLicType)
                {
                    case "01":
                        strBoard = "CPA";
                        break;
                    case "02":
                        strBoard = "Architecture";
                        break;
                    case "03":
                        strBoard = "Barbers";
                        break;
                    case "04":
                        strBoard = "Home Inspectors";
                        break;
                    case "05":
                        strBoard = "Cosmetology";
                        break;
                    case "06":
                        strBoard = "CILB";
                        break;
                    case "07":
                        strBoard = "Mold";
                        break;
                    case "08":
                        strBoard = "Electrical";
                        break;
                    case "09":
                        strBoard = "CPE";
                        break;
                    case "10":
                        strBoard = "PMW";
                        break;
                    case "11":
                        strBoard = "Funeral Directors and Embalmers";
                        break;
                    case "12":
                        strBoard = "Surveyors & Mappers";
                        break;
                    case "13":
                        strBoard = "Landscape Architecture";
                        break;
                    case "20":
                        strBoard = "Hotels and Restaurants";
                        break;
                    case "21":
                        strBoard = "Elevator Safety";
                        break;
                    case "23":
                        strBoard = "Harbor Pilot";
                        break;
                    case "25":
                        strBoard = "FREC";
                        break;
                    case "26":
                        strBoard = "Veterinary Medicine";
                        break;
                    case "33":
                        strBoard = "DDC";
                        break;
                    case "38":
                        strBoard = "CAM";
                        break;
                    case "40":
                        strBoard = "ABT";
                        break;
                    case "48":
                        strBoard = "Auctioneers";
                        break;
                    case "49":
                        strBoard = "Talent Agents";
                        break;
                    case "50":
                        strBoard = "Building Code Administrators";
                        break;
                    case "53":
                        strBoard = "Geologists";
                        break;
                    case "59":
                        strBoard = "Asbestos";
                        break;
                    case "60":
                        strBoard = "Athletic Agents";
                        break;
                    case "63":
                        strBoard = "Employee Leasing";
                        break;
                    case "64":
                        strBoard = "FREAB";
                        break;
                    case "74":
                        strBoard = "Labor Organization";
                        break;
                    case "75":
                        strBoard = "Farm Labor";
                        break;
                    case "76":
                        strBoard = "Child Labor";
                        break;
                    case "80":
                        strBoard = "Condos, Coops, Timeshares";
                        break;
                    case "81":
                        strBoard = "Mobile Homes";
                        break;
                    case "82":
                        strBoard = "Land Sales";
                        break;
                    case "83":
                        strBoard = "Other Entities";
                        break;
                    case "85":
                        strBoard = "Yacht and Ship Brokers";
                        break;
                    case "90":
                        strBoard = "Office of the Secretary";
                        break;
                    case "98":
                        strBoard = "OGC Miscellaneous";
                        break;
                    case "99":
                        strBoard = "Regulation";
                        break;
                }

                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Board KW to = {0}", strBoard));

                Keyword kwdBoard = null;
                if (!String.IsNullOrEmpty(strBoard))
                {
                    KeywordType kwtBoard = app.Core.KeywordTypes.Find(gSaveToBoard);
                    if (kwtBoard != null)
                        kwdBoard = CreateKeywordHelper(kwtBoard, strBoard);
                }

                args.PropertyBag.TryGetValue("pbRegion", out propRegion);
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Region = {0}", propRegion));

                string strRegionFullName = "";

                switch (propRegion)
                {
                    case "01":
                        strRegionFullName = "Fort Walton";
                        break;
                    case "02":
                        strRegionFullName = "Tallahassee";
                        break;
                    case "03":
                        strRegionFullName = "Jacksonville";
                        break;
                    case "04":
                        strRegionFullName = "Gainesville";
                        break;
                    case "05":
                        strRegionFullName = "Orlando";
                        break;
                    case "06":
                        strRegionFullName = "Tampa";
                        break;
                    case "07":
                        strRegionFullName = "Fort Myers";
                        break;
                    case "08":
                        strRegionFullName = "West Palm Beach";
                        break;
                    case "09":
                        strRegionFullName = "Margate";
                        break;
                    case "10":
                        strRegionFullName = "Miami";
                        break;
                    default:
                        strRegionFullName = "";
                        break;
                }

                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Region Value = {0}", strRegionFullName));

                string strFullComplainant = "DBPR - " + strRegionFullName;

                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Complainant KW to = {0}", strFullComplainant));


                Keyword kwdRegion = null;
                if (!String.IsNullOrEmpty(strRegionFullName))
                {
                    KeywordType kwtRegion = app.Core.KeywordTypes.Find(gSaveToComplainant);
                    if (kwtRegion != null)
                        kwdRegion = CreateKeywordHelper(kwtRegion, strRegionFullName);
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

                    // Add update keyword call to keyword modifier object
                    //Note Overloads available for use
                    //(I.E.): keyModifier.AddKeyword(keywordTypeName,keywordValue)
                    if (kwdBoard != null) keyModifier.AddKeyword(kwdBoard);
                    if (kwdRegion!= null) keyModifier.AddKeyword(kwdRegion);

                    // Apply keyword change to the document
                    keyModifier.ApplyChanges();

                    string output = String.Format("Keyword: '{0}' Value: '{1}', {5}Keyword: '{2}' Value: '{3}', {5}added to Document {4}.",
                        gSaveToBoard, strBoard, gSaveToComplainant, strFullComplainant, _currentDocument.ID, Environment.NewLine);
                    //Output the results to the OnBase Diagnostics Console
                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output);
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
            //  _currentDocument = app.Core.GetDocumentByID(TestDocId);
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

        /// <summary>
        /// Clean Seed Keyword
        /// </summary>
        /// <param name="pValue">Keyword Value</param>
        private string CleanSeedKW(string pValue)
        {
            string temp = pValue;
            StringBuilder newString = new StringBuilder();

            int firstSpace = pValue.IndexOf(" ");
            if (firstSpace > 0)
            {
                temp = pValue.Substring(0, firstSpace);
                foreach (char c in temp)
                {
                    if (((c >= '0') && (c <= '9')) || ((c >= 'A') && (c <= 'Z')))
                    {
                        newString.Append(c);
                    }
                    else
                    {
                        newString.Append(" ");
                    }
                }
                newString = newString.Replace(" ", "");
                newString = newString.Replace("'", "''");
                return newString.ToString();
            }
            else
            {
                return pValue.ToString();
            }
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