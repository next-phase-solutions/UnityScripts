﻿// Skeleton generated by Hyland Unity Editor on 3/30/2017 12:35:32 PM
namespace CRDEMailAddressLookUp
{
    using System;
    using System.Text;
    using System.Data;
    using System.Data.Odbc;
    using Hyland.Unity;


    /// <summary>
    /// 
    /// </summary>
    public class CRDEMailAddressLookUp : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "419 - ABT - SYS - Complaint Lookup (Secondary)";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        //Search parameters and Save to Keywords
        private const string gSaveToCaseNum = "Case #";
        private const string gSaveToInspectionNum = "Inspection #";
        private const string gSaveToInspectionVisitID = "Inspection Visit ID";
        //private const string gSaveToInspectionVisitID = "License Type";

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
        /// <param name="app"></param>
        /// <param name="args"></param>
        public void OnWorkflowScriptExecute(Hyland.Unity.Application app, Hyland.Unity.WorkflowEventArgs args)
        {
            try
            {
                // Initialize global settings
                IntializeScript(ref app, ref args);

                //get and clean Inspection Visit ID keyword for passing to LicEase database - says license type, but is passing ID

                KeywordType kwtLicenseType = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gSaveToInspectionVisitID);
                string strLicenseType = "";
                if (kwtLicenseType != null)
                {
                    KeywordRecord keyRecLicenseType = _currentDocument.KeywordRecords.Find(kwtLicenseType);
                    if (keyRecLicenseType != null)
                    {
                        Keyword kwdLicenseType = keyRecLicenseType.Keywords.Find(kwtLicenseType);
                        if (kwdLicenseType != null)
                            strLicenseType = CleanSeedKW(kwdLicenseType.ToString());
                    }
                }

                if (strLicenseType == "")
                {
                    throw new Exception(string.Format("Inspection Visit ID is blank!"));
                }

                //access Config Item for LicEase User
                string gUSER = "";
                if (app.Configuration.TryGetValue("LicEaseUser", out gUSER))
                {
                }

                //access Config Item for LicEase Password
                string gPASS = "";
                if (app.Configuration.TryGetValue("LicEasePassword", out gPASS))
                {
                }

                /* COMMENT THIS SECTION OUT WHEN MOVING TO PROD */
                //access Config Item for LicEase UAT ODBC
                string gODBC = "";
                if (app.Configuration.TryGetValue("LicEaseUAT", out gODBC))
                {
                }

                /* UNCOMMENT THIS SECTION WHEN MOVING TO PROD
				//access Config Item for LicEase PROD ODBC
				string gODBC = "";
				if (app.Configuration.TryGetValue("LicEasePROD", out gODBC))
				{
				}
				*/

                string connectionString = string.Format("DSN={0};Uid={1};Pwd={2};", gODBC, gUSER, gPASS);
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Connection string: {0}", connectionString));

                StringBuilder strSql = new StringBuilder();
                strSql.Append(@"select a.cmpln_nbr As COMPLAINTNUM from insp_vst iv, insp_hist ih, cmpln a ");
                strSql.Append(@"  where iv.insp_vst_id = '");
                strSql.Append(strLicenseType);
                strSql.Append(@"' and iv.insp_hist_id = ih.insp_hist_id and ih.cmpln_id = a.cmpln_id");

                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Sql Query: {0}", strSql.ToString()));

                using (OdbcConnection con = new OdbcConnection(connectionString))
                {
                    try
                    {
                        con.Open();
                        using (OdbcCommand command = new OdbcCommand(strSql.ToString(), con))
                        using (OdbcDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                string strCaseNum = "";

                                reader.Read();

                                strCaseNum = reader["COMPLAINTNUM"].ToString();
                                
                                Keyword kwdLicNum = null;
                                if (!String.IsNullOrEmpty(strCaseNum))
                                {
                                    KeywordType kwtLicNum = app.Core.KeywordTypes.Find(gSaveToCaseNum);
                                    if (kwtLicNum != null)
                                        kwdLicNum = CreateKeywordHelper(kwtLicNum, strCaseNum);
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
                                    if (kwdLicNum != null) keyModifier.AddKeyword(kwdLicNum);

                                    // Apply keyword change to the document
                                    keyModifier.ApplyChanges();

                                    string output = String.Format("Keyword: '{0}' Value: '{1}' added to Document {2}.", gSaveToCaseNum, strCaseNum, _currentDocument.ID);
                                    //Output the results to the OnBase Diagnostics Console
                                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output);

                                    documentLock.Release();
                                }
                            }
                            else
                            {
                                throw new Exception(string.Format("No records found in database for  {0}='{1}' and {4}{2}='{3}' ", gSaveToInspectionVisitID, strLicenseType, gParamAppNum, strCaseNum, Environment.NewLine));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException("Error during database operations!", ex);
                    }
                    finally
                    {
                        if (con.State == ConnectionState.Open) con.Close();
                    }
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