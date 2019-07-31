using System;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using Hyland.Unity;


namespace UnityScripts
{
    /// <summary>
    /// Description
    /// </summary>
    public class DBPR_UnityScript_146 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "146 - BET - Assign App Proc Load Balance User";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gODBC = "OBTEST13";
        private const string gUSER = "hsi";
        private const string gPASS = "wstinol";

        private const string exREASON = "Exception Reason";

        // Search parameter Keywords
        private const string gParamLicType = "License Type";
        private const string gParamTranCode = "Tran Code";
        private const string gParamBatchType = "Batch Type";

        // Output KWs
        private const string gSaveToAssignedProcUnit = "Assigned Processing Unit";
        private const string gSaveToAssignedProc = "Assigned Processor";

        // Lifecycle
        private const string gLCNUM = "114";
        private string strWFUser = "";
        private StringBuilder strSql = new StringBuilder();

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
        // public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args=null)
        {
            try
            {
                // Initialize global settings
                IntializeScript(ref app, ref args);

                KeywordType ktwBatchType = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamBatchType);
                string strBatchType = "";
                if (ktwBatchType != null)
                {
                    KeywordRecord keyRecBatchType = _currentDocument.KeywordRecords.Find(ktwBatchType);
                    if (keyRecBatchType != null)
                    {
                        Keyword kwdBatchType = keyRecBatchType.Keywords.Find(ktwBatchType);
                        if (kwdBatchType != null)
                            strBatchType = kwdBatchType.ToString();
                    }
                }

                if (strBatchType == "")
                {
                    throw new Exception(string.Format("{0} is blank.", gParamBatchType));
                }

                KeywordType ktwTranCode = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamTranCode);
                string strTranCode = "";
                if (ktwTranCode != null)
                {
                    KeywordRecord keyRecTranCode = _currentDocument.KeywordRecords.Find(ktwTranCode);
                    if (keyRecTranCode != null)
                    {
                        Keyword kwdTranCode = keyRecTranCode.Keywords.Find(ktwTranCode);
                        if (kwdTranCode != null)
                            strTranCode = kwdTranCode.ToString();
                    }
                }

                KeywordType ktwLicType = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamLicType);
                string strLicType = "";
                if (ktwLicType != null)
                {
                    KeywordRecord keyRecLicType = _currentDocument.KeywordRecords.Find(ktwLicType);
                    if (keyRecLicType != null)
                    {
                        Keyword kwdLicType = keyRecLicType.Keywords.Find(ktwLicType);
                        if (kwdLicType != null)
                            strLicType = kwdLicType.ToString();
                    }
                }

                if (strLicType == "")
                {
                    throw new Exception(string.Format("{0} is blank.", gParamLicType));
                }

                //access Config Item for OnBase User
                string gUSER = "";
                if (app.Configuration.TryGetValue("OnBaseUser", out gUSER))
                {
                }

                //access Config Item for OnBase Password
                string gPASS = "";
                if (app.Configuration.TryGetValue("OnBasePassword", out gPASS))
                {
                }

                /* COMMENT THIS SECTION OUT WHEN MOVING TO PROD */
                //access Config Item for OnBase UAT ODBC
                string gODBCBET = "";
                if (app.Configuration.TryGetValue("BETOnBaseMISC", out gODBCBET))
                {
                }

                string gODBCOnBase = "";
                if (app.Configuration.TryGetValue("OnBaseUAT", out gODBCOnBase))
                {
                }

                /* UNCOMMENT THIS SECTION WHEN MOVING TO PROD
				//access Config Item for OnBase PROD ODBC
				string gODBCOnBase = "";
				if (app.Configuration.TryGetValue("OnBasePROD", out gODBCOnBase))
				{
				}
				*/

                string obDocTypeFull = _currentDocument.DocumentType.Name.ToString();
                string obDocType = obDocTypeFull.Substring(7);

                if (strTranCode != "" && strLicType != "")
                {
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Trans Code and License Type Present"));

                    //StringBuilder strSql = new StringBuilder();
                    strSql.Append(@"SELECT users AS USER, proc_unit AS UNIT, last_user AS LASTUSER ");
                    strSql.Append(@"  FROM dbo.tbl_ccb ");
                    strSql.Append(@"  WHERE license_type = '");
                    strSql.Append(strLicType);
                    strSql.Append(@"' AND tran_code = '");
                    strSql.Append(strTranCode);
                    strSql.Append(@"'");

                                    }
                else if (strLicType != "" && strTranCode == "" && (strBatchType == "BET - CAND. (FEE)" || strBatchType == "BET - CAND. (NON FEE)"))
                {
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("BET - CAND FEE or NON FEE"));

                    string strMidLicType = strLicType.Substring(1, 2);

                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("LicType {0}, Doc Type - Description {1}", strMidLicType, obDocType));

                    //StringBuilder strSql = new StringBuilder();
                    strSql.Append(@"SELECT users AS USER, proc_unit AS UNIT, last_user AS LASTUSER ");
                    strSql.Append(@"  FROM dbo.tbl_ccb ");
                    strSql.Append(@"  WHERE license_type = '");
                    strSql.Append(strMidLicType);
                    strSql.Append(@"' AND description = '");
                    strSql.Append(obDocType);
                    strSql.Append(@"'");

                }
                else if (strLicType != "" && strTranCode == "")
                {
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("License Type but NO Trans Code"));

                    strSql.Append(@"SELECT users AS USER, proc_unit AS UNIT, last_user AS LASTUSER ");
                    strSql.Append(@"  FROM dbo.tbl_ccb ");
                    strSql.Append(@"  WHERE license_type = '");
                    strSql.Append(strLicType);
                    strSql.Append(@"' AND description = '");
                    strSql.Append(obDocType);
                    strSql.Append(@"'");
                }

                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Sql Query: {0}", strSql.ToString()));

                string connectionString = string.Format("DSN={0};Uid={1};Pwd={2};", gODBCBET, gUSER, gPASS);
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Connection string: {0}", connectionString));

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
                                //string strUser = "";
                                string strUnit = "";
                                List<string> lstUser = new List<string>();
                                //string strLastUser = "";

                                reader.Read();

                                lstUser.Add(reader["USER"].ToString());
                                //strUser = reader["USER"].ToString();
                                //strLastUser = reader["LASTUSER"].ToString();
                                strUnit = reader["UNIT"].ToString();

                                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Checking Users"));

                                string strSQLUsers = "";

                                foreach (string result in lstUser)
                                {
                                    if (result != "" && strSQLUsers == "")
                                    {
                                        strSQLUsers = "'" + strSQLUsers + result + "'";
                                    }
                                    else if (result != "" && strSQLUsers !="")
                                    {
                                        strSQLUsers = strSQLUsers + ",'" + result + "'";
                                    }
                                }

                                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("users to pass to OnBase " + strSQLUsers));

                                //StringBuilder strSql = new StringBuilder();
                                strSql.Append(@"SELECT TOP 1 ua.username AS WFUSER ");
                                strSql.Append(@"  FROM hsi.useraccount ua ");
                                strSql.Append(@"  LEFT OUTER JOIN (SELECT usernum FROM hsi.itemlcxuser ");
                                strSql.Append(@"  WHERE hsi.itemlcxuser.lcnum = ");
                                strSql.Append(gLCNUM);
                                strSql.Append(@" ) ilcu ON us.usernum = ilcu.usernum ");
                                strSql.Append(@"  WHERE ua.username IN(");
                                strSql.Append(strSQLUsers);
                                strSql.Append(@"'");

                                string connectionStringOnBase = string.Format("DSN={0};Uid={1};Pwd={2};", gODBCOnBase, gUSER, gPASS);
                                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Connection string: {0}", connectionStringOnBase));

                                using (OdbcConnection conOnBase = new OdbcConnection(connectionStringOnBase))
                                {
                                    try
                                    {
                                        conOnBase.Open();
                                        using (OdbcCommand commandOnBase = new OdbcCommand(strSql.ToString(), conOnBase))
                                        using (OdbcDataReader readerOnBase = commandOnBase.ExecuteReader())
                                        {
                                            if (readerOnBase.HasRows)
                                            {
                                                strWFUser = readerOnBase["WFUSER"].ToString();

                                                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("WFUser set to " + strWFUser));
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new ApplicationException("Error during database operations!", ex);
                                    }
                                    finally
                                    {
                                        if (conOnBase.State == ConnectionState.Open) conOnBase.Close();
                                    }
                                }

                                Keyword kwdUnit = null;
                                if (!String.IsNullOrEmpty(strUnit))
                                {
                                    KeywordType kwtUnit = app.Core.KeywordTypes.Find(gSaveToAssignedProcUnit);
                                    if (kwtUnit != null)
                                        kwdUnit = CreateKeywordHelper(kwtUnit, strUnit);
                                }

                                Keyword kwdUser = null;
                                if (!String.IsNullOrEmpty(strWFUser))
                                {
                                    KeywordType kwtUser = app.Core.KeywordTypes.Find(gSaveToAssignedProc);
                                    if (kwtUser != null)
                                        kwdUser = CreateKeywordHelper(kwtUser, strWFUser);
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
                                    if (kwdUnit != null) keyModifier.AddKeyword(kwdUnit);
                                    if (kwdUser != null) keyModifier.AddKeyword(kwdUser);

                                    // Apply keyword change to the document
                                    keyModifier.ApplyChanges();

                                    string output = String.Format("Keyword: '{0}' Value: '{1}', {3}added to Document {2}.",
                                        gSaveToAssignedProcUnit, strUnit, _currentDocument.ID, Environment.NewLine);

                                    //Output the results to the OnBase Diagnostics Console
                                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output);

                                    string output2 = String.Format("Keyword: '{0}' Value: '{1}', {3}added to Document {2}.",
                                        gSaveToAssignedProc, strWFUser, _currentDocument.ID, Environment.NewLine);

                                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output2);
                                }
                            }
                            else
                            {
                                throw new Exception(string.Format("No records found in database for  {0}='{1}'", gParamBatchType, strBatchType));
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