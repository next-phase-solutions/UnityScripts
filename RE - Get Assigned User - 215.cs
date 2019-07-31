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
    public class DBPR_UnityScript_215 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "215 - RE - Get Assigned User by Related Casae # in Analyst Queue";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";


        private const string exREASON = "Exception Reason";
        private const string gSaveToReassignedAnalyst = "REassignedAnalyst";
        private const string gPropCase = "relatedCase";

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

                string strRelatedCase = "";
                args.SessionPropertyBag.TryGetValue(gPropCase, out strRelatedCase);

                app.Diagnostics.Write(string.Format("property for related case # = {0}", strRelatedCase));
                
                if (string.IsNullOrEmpty(strRelatedCase))
                {
                    throw new Exception(String.Format("Property '{0}' not found", strRelatedCase));
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
                string gODBC = "";
                if (app.Configuration.TryGetValue("OBUAT", out gODBC))
                {
                }

                /* UNCOMMENT THIS SECTION WHEN MOVING TO PROD
				//access Config Item for OnBase PROD ODBC
				string gODBC = "";
				if (app.Configuration.TryGetValue("OnBasePROD", out gODBC))
				{
				}
				*/

                string connectionString = string.Format("DSN={0};Uid={1};Pwd={2};", gODBC, gUSER, gPASS);
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Verbose, string.Format("Connection string: {0}", connectionString));

                StringBuilder strSql = new StringBuilder();
                strSql.Append("select distinct ua.username AS USER FROM hsi.itemlc ilc ");
                strSql.Append("join hsi.itemlcxuser ilcxu on ilc.itemnum = ilcxu.itemnum ");
                strSql.Append("join hsi.useraccount ua on ilcxu.usernum = ua.usernum ");
                strSql.Append("join hsi.keyxitem138 ki138 on ilc.itemnum = ki138.itemnum ");
                strSql.Append("join hsi.keytable138 kt138 on ki138.keywordnum = kt138.keywordnum ");
                strSql.Append("where ilc.lcnum = 177 and ilc.statenum = 568 and kt138.keyvaluechar = '");
                strSql.Append(strRelatedCase);
                strSql.Append("'");

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
                                string strProf = "";

                                reader.Read();

                                strProf = reader["USER"].ToString();

                                Keyword kwdProf = null;
                                if (!String.IsNullOrEmpty(strProf))
                                {
                                    KeywordType kwtProf = app.Core.KeywordTypes.Find(gSaveToReassignedAnalyst);
                                    if (kwtProf != null)
                                        kwdProf = CreateKeywordHelper(kwtProf, strProf);
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
                                    if (kwdProf != null) keyModifier.AddKeyword(kwdProf);

                                    // Apply keyword change to the document
                                    keyModifier.ApplyChanges();

                                    string output = String.Format("Keyword: '{0}' Value: '{1}', {3}added to Document {2}.",
                                        gSaveToReassignedAnalyst, strProf, _currentDocument.ID, Environment.NewLine);
                                    //Output the results to the OnBase Diagnostics Console
                                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output);

                                    documentLock.Release();
                                }
                            }
                            else
                            {
                                throw new Exception(string.Format("No records found in database for  {0}='{1}'", gPropCase, strRelatedCase));
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