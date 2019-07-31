using System;
using System.Text;
using System.Data;
using System.Collections.Generic;
using System.Data.Odbc;
using Hyland.Unity;
using Application = Hyland.Unity.Application;

namespace UnityScripts
{
    /// <summary>
    ///  Update keywords from PropertyBag values
    /// </summary>
    public class DBPR_UnityScript_449 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "449 - HR - Compliance Lookup to VR for CSS - 03";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gParamFileNumber = "File Number";
        private const string gParamLicType = "License Type";

        private List<string> lstProp1 = new List<string>();
        private List<string> lstProp2 = new List<string>();
        private List<string> lstProp3 = new List<string>();
        private List<string> lstProp4 = new List<string>();
        private List<string> lstProp5 = new List<string>();
        private List<string> lstProp6 = new List<string>();
        private List<string> lstProp7 = new List<string>();


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

                KeywordType kwtLicenseNum = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamFileNumber);
                string strFileNum = "";
                if (kwtLicenseNum != null)
                {
                    KeywordRecord keyRecLicenseNum = _currentDocument.KeywordRecords.Find(kwtLicenseNum);
                    if (keyRecLicenseNum != null)
                    {
                        Keyword kwdLicenseNum = keyRecLicenseNum.Keywords.Find(kwtLicenseNum);
                        if (kwdLicenseNum != null)
                            strFileNum = CleanSeedKW(kwdLicenseNum.ToString());
                    }
                }

                if ((strFileNum == "") || (strLicenseType == ""))
                {
                    throw new Exception(string.Format("Either {0} or {1} is blank.", gParamFileNumber, gParamLicType));
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
                strSql.Append(@"select * from (select distinct(c.cmpln_nbr) as CASENUMBER, becec.enf_cde as ACTYPE, c.rcv_dte as ISSUEDATE, to_char(cacat.actv_strt_dte, 'MM/DD/YYYY')  as CLERKEDDATE, ccscs.cmpln_sta_cde as COMPSTATUS, c.clnt_cde as LICTYPE ");
                strSql.Append(@", c1_mc.chrg_amt as FINEAMT  from cmpln c inner join (select ec.enf_cde,  bec.clnt_enf_cde_id from brd_enf_cde bec inner join enf_cde ec on bec.enf_cde_id = ec.enf_cde_id where (ec.enf_cde like 'CLOS' or ec.clnt_cde like 'AD%' or ec.enf_cde like 'AC%') ) becec on  c.clnt_cmpln_cls_id = becec.clnt_enf_cde_id inner join ");
                strSql.Append(@"(select cs.cmpln_sta_cde, ccs.clnt_cmpln_sta_id from clnt_cmpln_sta ccs inner join  cmpln_sta cs on ccs.cmpln_sta_id = cs.cmpln_sta_id ) ccscs on c.clnt_cmpln_sta_id = ccscs.clnt_cmpln_sta_id inner join (select  l.file_nbr,  l.clnt_cde,  l.lic_id,  r.cmpln_id from rspn r inner join lic l on  r.lic_id = l.lic_id )  rl  ");
                strSql.Append(@" on c.cmpln_id = rl.cmpln_id left outer join  (select ca.actv_strt_dte, ca.cmpln_id from cmpln_actv ca inner join cmpln_actv_typ cat on ca.cmpln_actv_typ_id = cat.cmpln_actv_typ_id where cat.cmpln_actv_typ_cde = 'A400' ) cacat on c.cmpln_id = cacat.cmpln_id left outer join (select mc.chrg_amt, c1.cmpln_id from cmply_ordr c1 inner join misc_chrg mc on c1.misc_chrg_id = mc.misc_chrg_id ) c1_mc on c.cmpln_id = c1_mc.cmpln_id where rl.file_nbr = '");
                strSql.Append(strFileNum);
                strSql.Append(@"' and rl.clnt_cde = '");
                strSql.Append(strLicenseType);
                strSql.Append(@"' and c.rcv_dte > (SYSDATE - 1826) and c.clnt_cde like '20%' and c.cmpln_nbr > '2010%') order by 1 desc ");


                


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
                                while (reader.Read())
                                {
                                    lstProp1.Add(reader["CASENUMBER"].ToString());
                                    lstProp2.Add(reader["ACTYPE"].ToString());
                                    lstProp3.Add(reader["ISSUEDATE"].ToString());
                                    lstProp4.Add(reader["CLERKEDDATE"].ToString());
                                    lstProp5.Add(reader["COMPSTATUS"].ToString());
                                    lstProp6.Add(reader["LICTYPE"].ToString());
                                    lstProp7.Add(reader["FINEAMT"].ToString());
                                }

                                // Create keyword modifier object to hold keyword changes
                                EForm currentForm = _currentDocument.EForm;

                                FieldModifier fieldModifier = currentForm.CreateFieldModifier();

                                foreach (string props in lstProp1)
                                {
                                    if (props != null) fieldModifier.UpdateField("CASENUMBER", props);
                                }

                                foreach (string props in lstProp2)
                                {
                                    if (props != null) fieldModifier.UpdateField("ACTYPE", props);
                                }

                                foreach (string props in lstProp3)
                                {
                                    if (props != null) fieldModifier.UpdateField("ISSUEDATE", props);
                                }

                                foreach (string props in lstProp4)
                                {
                                    if (props != null) fieldModifier.UpdateField("CLERKEDDATE", props);
                                }

                                foreach (string props in lstProp5)
                                {
                                    if (props != null) fieldModifier.UpdateField("COMPSTATUS", props);
                                }

                                foreach (string props in lstProp6)
                                {
                                    if (props != null) fieldModifier.UpdateField("LICTYPE", props);
                                }

                                foreach (string props in lstProp7)
                                {
                                    if (props != null) fieldModifier.UpdateField("FINEAMT", props);
                                }
                                
                                using (DocumentLock documentLock = _currentDocument.LockDocument())
                                {
                                    // Ensure lock was obtained
                                    if (documentLock.Status != DocumentLockStatus.LockObtained)
                                    {
                                        throw new Exception("Document lock not obtained");
                                    }

                                    // Apply keyword change to the document
                                    fieldModifier.ApplyChanges();

                                    documentLock.Release();

                                }
                            }
                            else
                            {
                                throw new Exception(string.Format("No records found in database for  {0}='{1}' and {4}{2}='{3}' ", gParamLicType, strLicenseType, gParamFileNumber, strFileNum, Environment.NewLine));
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