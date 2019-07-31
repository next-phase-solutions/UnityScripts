using System;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using Hyland.Unity;


namespace UnityScripts
{
    /// <summary>
    /// </summary>
    public class DBPR_UnityScript_110 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "PMW - RULE - Assign App Proc Load Balance User from DB - 374";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gParamLicType = "License Type";     // Key Number 104
        private const string gParamTranCode = "Tran Code";       // Key Number 106
        private const string gSaveToAssignedProcUnit = "Assigned Processing Unit";
        private const string gSaveToAssignedProc = "Assigned Processor";

        private const string gLCNUM = "146";

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
        //    public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args=null)
        {
            try
            {
                // Initialize global settings
                IntializeScript(ref app, ref args);

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

                if (strLicType == "" || strTranCode == "")
                {
                    throw new Exception(string.Format("{0} or {1} is blank.", gParamLicType, gParamTranCode));
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
                string gODBCMISC = "";
                if (app.Configuration.TryGetValue("OnbaseMISC", out gODBCMISC))
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

                // **************
                // Get Routing Values - This query gets the Proc Unit, Users and LastUser from SQL

                strSql.Append(@"SELECT users AS USERS, proc_unit AS UNIT ");
                strSql.Append(@"  FROM dbo.pmwccb ");
                strSql.Append(@"  WHERE license_type = '");
                strSql.Append(strLicType);
                strSql.Append(@"' AND tran_code = '");
                strSql.Append(strTranCode);
                strSql.Append(@"'");

                string connectionString = string.Format("DSN={0};Uid={1};Pwd={2};", gODBCMISC, gUSER, gPASS);
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Sql Query: {0}", strSql.ToString()));


                using (OdbcConnection con = new OdbcConnection(connectionString))
                {
                    try
                    {
                        con.Open();
                        using (OdbcCommand command = new OdbcCommand(strSql, con))
                        using (OdbcDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                reader.Read();

                                string strUnit = "";
                                string lstUser = "";

                                lstUser = reader["USERS"].ToString();
                                strUnit = reader["UNIT"].ToString();

                                string[] result = lstUser.Split(',');

                                string strSQLUsers = "";

                                foreach (string item in result)
                                {
                                    if (item != "" && strSQLUsers == "")
                                    {
                                        strSQLUsers = "'" + item + "'";
                                    }
                                    else if (item != "" && strSQLUsers != "")
                                    {
                                        strSQLUsers = strSQLUsers + ",'" + item + "'";
                                    }
                                }

                                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Users from CIU Database = {0}", strSQLUsers));

                                strSql.Append(@"select CT1.username, sum(CT1.DocCount) as DocCount from ((select ua.username, count(ilcu.usernum) as DocCount ");
                                strSql.Append(@" from hsi.useraccount ua left outer join (select usernum,itemnum from hsi.itemlcxuser where hsi.itemlcxuser.lcnum = ");
                                strSql.Append(gLCNUM);
                                strSql.Append(@") ilcu on ua.usernum = ilcu.usernum inner join hsi.itemlc il on il.itemnum = ilcu.itemnum inner join hsi.itemdata id on ilcu.itemnum = id.itemnum ");
                                strSql.Append(@" inner join hsi.lcstate lcs on il.statenum = lcs.statenum where ua.username in (");
                                strSql.Append(strSQLUsers);
                                strSql.Append(@") and id.status=0 and lcs.statename like 'PMWAP - Initial Review%' and lcs.statename not like 'PMWAP - Research' ");
                                strSql.Append(@" and lcs.statename not like 'PMWAP - Pending Review' and lcs.statename not like 'PMWAP - Application Processing Routing Exceptions' group by ua.username) ");
                                strSql.Append(@" UNION (select ua.username, 0 as DocCount from hsi.useraccount ua where ua.username in ( ");
                                strSql.Append(strSQLUsers);
                                strSql.Append(@") ) ) CT1 group by CT1.username order by DocCount asc ");

                                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("2nd Sql Query: {0}", strSql.ToString()));

                                string connectionStringOnBase = string.Format("DSN={0};Uid={1};Pwd={2};", gODBCOnBase, gUSER, gPASS);
                                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Connection string: {0}", connectionStringOnBase));

                                string strNextUser = "";
                                using (OdbcConnection conOnBase = new OdbcConnection(connectionStringOnBase))
                                {
                                    try
                                    {
                                        conOnBase.Open();
                                        using (OdbcCommand command = new OdbcCommand(strSql, conOnBase))
                                        using (OdbcDataReader readerOnBase = command.ExecuteReader())
                                        {
                                            if (readerOnBase.HasRows)
                                            {
                                                while (readerOnBase.Read())
                                                {
                                                    strNextUser = readerOnBase[0].ToString();
                                                }
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

                                Keyword kwdUser = null;
                                if (!String.IsNullOrEmpty(strNextUser))
                                {
                                    KeywordType kwtUser = app.Core.KeywordTypes.Find(gSaveToAssignedProc);
                                    if (kwtUser != null)
                                        kwdUser = CreateKeywordHelper(kwtUser, strNextUser);
                                }


                                Keyword kwdUnit = null;
                                if (!String.IsNullOrEmpty(strUnit))
                                {
                                    KeywordType kwtUnit = app.Core.KeywordTypes.Find(gSaveToAssignedProcUnit);
                                    if (kwtUnit != null)
                                        kwdUnit = CreateKeywordHelper(kwtUnit, strUnit);
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

                                    documentLock.Release();

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
                                throw new Exception(string.Format("No records found in database"));
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