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
    public class DBPR_UnityScript_424 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "424 - HR - FL - Get File Number";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gParamLicType = "License Type";
        private const string gParamLicNum = "License Number";
        private const string gSaveToFileNum = "File Number";
        private const string gSaveToKeyName = "Key Name";
        private const string gSaveToIndNum = "Indv/Org Number";
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

                KeywordType kwtLicenseType = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamLicType);
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

                KeywordType kwtLicenseNum = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamLicNum);
                string strLicenseNum = "";
                if (kwtLicenseNum != null)
                {
                    KeywordRecord keyRecLicenseNum = _currentDocument.KeywordRecords.Find(kwtLicenseNum);
                    if (keyRecLicenseNum != null)
                    {
                        Keyword kwdLicenseNum = keyRecLicenseNum.Keywords.Find(kwtLicenseNum);
                        if (kwdLicenseNum != null)
                            strLicenseNum = CleanSeedKW(kwdLicenseNum.ToString());
                    }
                }

                if ((strLicenseNum == "") || (strLicenseType == ""))
                {
                    throw new Exception(string.Format("Either {0} or {1} is blank.", gParamLicNum, gParamLicType));
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
                strSql.Append(@"SELECT l.file_nbr AS FILENUM, n.key_nme AS KEYNAME, l.xent_id AS INDNUM ");
                strSql.Append(@"  FROM Lic l ");
                strSql.Append(@"  left join name n on l.xent_id = n.xent_id  ");
                strSql.Append(@"  WHERE l.clnt_cde = '");
                strSql.Append(strLicenseType);
                strSql.Append(@"' CAST(l.lic_nbr AS number) = '");
                strSql.Append(strLicenseNum);
                strSql.Append(@"' AND n.ent_nme_typ = 'D' AND n.cur_nme_ind = 'Y' AND rownum = '1'");

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
                                string strFileNum = "";
                                string strKeyName = "";
                                string strIndNum = "";

                                reader.Read();

                                strFileNum = reader["FILENUM"].ToString();
                                strKeyName = reader["KEYNAME"].ToString();
                                strIndNum = reader["INDNUM"].ToString();

                                Keyword kwdFileNum = null;
                                if (!String.IsNullOrEmpty(strFileNum))
                                {
                                    KeywordType kwtFileNum = app.Core.KeywordTypes.Find(gSaveToFileNum);
                                    if (kwtFileNum != null)
                                        kwdFileNum = CreateKeywordHelper(kwtFileNum, strFileNum);
                                }
                                Keyword kwdKeyName = null;
                                if (!String.IsNullOrEmpty(strKeyName))
                                {
                                    KeywordType kwtKeyName = app.Core.KeywordTypes.Find(gSaveToKeyName);
                                    if (kwtKeyName != null)
                                        kwdKeyName = CreateKeywordHelper(kwtKeyName, strKeyName);
                                }
                                Keyword kwdIndNum = null;
                                if (!String.IsNullOrEmpty(strIndNum))
                                {
                                    KeywordType kwtIndNum = app.Core.KeywordTypes.Find(gSaveToIndNum);
                                    if (kwtIndNum != null)
                                        kwdIndNum = CreateKeywordHelper(kwtIndNum, strIndNum);
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
                                    if (kwdFileNum != null) keyModifier.AddKeyword(kwdFileNum);
                                    if (kwdKeyName != null) keyModifier.AddKeyword(kwdKeyName);
                                    if (kwdIndNum != null) keyModifier.AddKeyword(kwdIndNum);

                                    // Apply keyword change to the document
                                    keyModifier.ApplyChanges();

                                    string output = String.Format("Keyword: '{0}' Value: '{1}', {7}Keyword: '{2}' Value: '{3}', {7}Keyword: '{4}' Value: '{5}', {7}added to Document {6}.",
                                        gSaveToFileNum, strFileNum, gSaveToKeyName, strKeyName, gSaveToIndNum, strIndNum, _currentDocument.ID, Environment.NewLine);
                                    //Output the results to the OnBase Diagnostics Console
                                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output);
                                }
                            }
                            else
                            {
                                throw new Exception(string.Format("No records found in database for  {0}='{1}' and {4}{2}='{3}' ", gParamLicType, strLicenseType, gParamLicNum, strLicenseNum, Environment.NewLine));
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