﻿// Skeleton generated by Hyland Unity Editor on 5/12/2017 1:14:45 PM
namespace HRInspectionLookuptoVR
{
    using System;
    using System.Text;
    using System.Data;
    using System.Data.Odbc;
    using Hyland.Unity;
    using Application = Hyland.Unity.Application;


    /// <summary>
    /// 388 - HR lookup to VR
    /// </summary>
    public class HRInspectionLookuptoVR : Hyland.Unity.IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "388 - HR - Inspection Lookup to VR";

        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string gSaveToLicNum = "License Number";
        private const string gSaveToLicType = "License Type";
        private const string gSaveToInspNum = "Inspection #";
        private const string gSaveToFileNum = "File Number";
        private const string gSaveToKeyName = "Key Name";
        private const string gSaveToSubject = "Subject";
        private const string gSaveToDBA = "DBA";
        private const string gSaveToIndOrgNum = "Indv/Org Number";
        private const string gSaveToVisitNum = "Visit #";
        private const string gSaveToInspectVisitDate = "Inspection Visit Date";
        private const string gSaveToDisposition = "Disposition";
        private const string gSaveToCity = "City";
        private const string gSaveToCounty = "County";
        private const string gSaveToRegion = "Region";
        private const string gSaveToInspectorName = "Inspector Name";
        private const string gSaveToInspTypeDesc = "Inspection Type (HR-FL)";
        private const string gParamInspectionID = "Inspection Visit ID";

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

                //get and clean InspectionID keyword for passing to LicEase database
                KeywordType kwtInspectionID = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(gParamInspectionID);
                string strInspectionID = "";
                if (kwtInspectionID != null)
                {
                    KeywordRecord keyRecInspectionID = _currentDocument.KeywordRecords.Find(kwtInspectionID);
                    if (keyRecInspectionID != null)
                    {
                        Keyword kwdInspectionID = keyRecInspectionID.Keywords.Find(kwtInspectionID);
                        if (kwdInspectionID != null)
                            strInspectionID = CleanSeedKW(kwdInspectionID.ToString());
                    }
                }

                if (strInspectionID == "")
                {
                    throw new Exception(string.Format("Search keyword {0} is blank. {3}, {4}, {5}", gParamInspectionID));
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
                strSql.Append(@"SELECT a.insp_nbr, key_name, dba_name as dba_kw, a.insp_vst_id, a.lic_type AS lic_type, a.indorg_num, a.file_num, ");
                strSql.Append(@" a.lic_num AS lic_num, a.visit_num, a.visit_date, a.insp_typ_desc, a.disposition_kw, city_kw, county_kw, region_kw, a.inspector_name  ");
                strSql.Append(@" FROM(SELECT DISTINCT TO_CHAR (insp_hist.insp_nbr) AS insp_nbr, TO_CHAR (insp_vst.insp_vst_id) AS insp_vst_id, lic.clnt_cde AS lic_type, lic.xent_id AS indorg_num, ");
                strSql.Append(@" lic.lic_nbr AS lic_num, lic.file_nbr as file_num, TO_CHAR(insp_vst.insp_vst_nbr) AS visit_num, TO_CHAR (insp_vst.insp_vst_strt_dte) AS visit_date,");
                strSql.Append(@" insp_typ_defn.insp_typ_desc, insp_disp_typ.insp_disp_typ_desc AS disposition_kw, stff.frst_nme || '.' || stff.surnme AS inspector_name, insp_hist.lic_id as insp_lic_id, ");
                strSql.Append(@" (select max(key_nme) from name n, link k where k.link_id = insp_hist.link_id and k.nme_id = n.nme_id) as key_name, ");
                strSql.Append(@" (SELECT MAX(key_nme) FROM NAME n, LINK k WHERE k.prnt_id = insp_hist.lic_id ");
                strSql.Append(@" AND k.link_prnt_cde = 'L' AND k.curr_ind = 'Y' AND k.clnt_link_typ_id IN (SELECT clnt_link_typ_id FROM clnt_link_typ c, link_typ t WHERE t.link_typ_id = c.link_typ_id AND t.link_typ_cde = 'DBA') ");
                strSql.Append(@" AND k.nme_id = n.nme_id) AS dba_name,(select addr_cty from addr n, link k where k.link_id = insp_hist.link_id and k.addr_id = n.addr_id) as city_kw, ");
                strSql.Append(@" (select cnty_desc from cnty c, addr n, link k where k.link_id = insp_hist.link_id and k.addr_id = n.addr_id and c.cnty = n.cnty) as county_kw, ");
                strSql.Append(@" (select  insp_regn_cde from insp_regn r, link k where k.link_id = insp_hist.link_id and k.insp_regn_id = r.insp_regn_id) as region_kw FROM insp_vst, insp_hist, ");
                strSql.Append(@" insp_typ_defn, insp_disp_typ, inspr, stff, lic WHERE insp_hist.insp_hist_id = insp_vst.insp_hist_id ");
                strSql.Append(@" AND insp_vst.insp_vst_id = (SELECT NVL(max(s.alt_insp_vst_id), max(s.insp_vst_id)) FROM insp_vst_synch s WHERE s.insp_vst_id = '");
                strSql.Append(strInspectionID);
                strSql.Append(@"') AND insp_typ_defn.insp_typ_defn_id = insp_hist.insp_typ_defn_id AND insp_vst.inspr_id = inspr.inspr_id AND stff.stff_oper_id = inspr.stff_oper_id AND lic.lic_id = insp_hist.lic_id AND insp_disp_typ.insp_disp_typ_id = insp_hist.insp_disp_typ_id) a ");


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
                                string strLicType = "";
                                string strLicNum = "";
                                string strFileNum = "";
                                string strKeyName = "";
                                string strDBAName = "";
                                string strInDorgNum = "";
                                string strVisitNum = "";
                                string strDisposition = "";
                                string strCity = "";
                                string strCounty = "";
                                string strRegion = "";
                                string strInspector = "";
                                string strInspNum = "";
                                string strSubject = "";
                                string strInspType = "";

                                reader.Read();

                                strLicType = reader["lic_Type"].ToString();
                                strLicNum = reader["lic_Num"].ToString();
                                strFileNum = reader["file_Num"].ToString();
                                strKeyName = reader["key_Name"].ToString();
                                strDBAName = reader["dba_kw"].ToString();
                                strInDorgNum = reader["indorg_num"].ToString();
                                strVisitNum = reader["visit_Num"].ToString();
                                strDisposition = reader["disposition_kw"].ToString();
                                strCity = reader["city_kw"].ToString();
                                strInspector = reader["inspector_name"].ToString();
                                strInspNum = reader["insp_nbr"].ToString();
                                strSubject = reader["key_Name"].ToString();
                                strInspType = reader["insp_typ_desc"].ToString();

                                if (reader["county_kw"] != DBNull.Value)
                                {
                                    strCounty = reader["county_kw"].ToString();
                                }
                                else
                                {
                                    strCounty = "Not Available";
                                }

                                if (reader["region_kw"] != DBNull.Value)
                                {
                                    strRegion = reader["region_kw"].ToString();
                                }
                                else
                                {
                                    strRegion = "Not Available";
                                }

                                Keyword kwdLicType = null;

                                if (!String.IsNullOrEmpty(strLicType))
                                {
                                    KeywordType kwtLicType = app.Core.KeywordTypes.Find(gSaveToLicType);
                                    if (kwtLicType != null)
                                        kwdLicType = CreateKeywordHelper(kwtLicType, strLicType);
                                }

                                Keyword kwdLicenseNum = null;
                                if (!String.IsNullOrEmpty(strLicNum))
                                {
                                    KeywordType kwtLicenseNum = app.Core.KeywordTypes.Find(gSaveToLicNum);
                                    if (kwtLicenseNum != null)
                                        kwdLicenseNum = CreateKeywordHelper(kwtLicenseNum, strLicNum);
                                }

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

                                Keyword kwdSubject = null;
                                if (!String.IsNullOrEmpty(strSubject))
                                {
                                    KeywordType kwtSubject = app.Core.KeywordTypes.Find(gSaveToSubject);
                                    if (kwtSubject != null)
                                        kwdSubject = CreateKeywordHelper(kwtSubject, strSubject);
                                }

                                Keyword kwdDBAName = null;
                                if (!String.IsNullOrEmpty(strDBAName))
                                {
                                    KeywordType kwtDBAName = app.Core.KeywordTypes.Find(gSaveToDBA);
                                    if (kwtDBAName != null)
                                        kwdDBAName = CreateKeywordHelper(kwtDBAName, strDBAName);
                                }

                                Keyword kwdInDorgNum = null;
                                if (!String.IsNullOrEmpty(strInDorgNum))
                                {
                                    KeywordType kwtInDorgNum = app.Core.KeywordTypes.Find(gSaveToIndOrgNum);
                                    if (kwtInDorgNum != null)
                                        kwdInDorgNum = CreateKeywordHelper(kwtInDorgNum, strInDorgNum);
                                }

                                Keyword kwdVisitNum = null;
                                if (!String.IsNullOrEmpty(strVisitNum))
                                {
                                    KeywordType kwtVisitNum = app.Core.KeywordTypes.Find(gSaveToVisitNum);
                                    if (kwtVisitNum != null)
                                        kwdVisitNum = CreateKeywordHelper(kwtVisitNum, strVisitNum);
                                }

                                Keyword kwdDisposition = null;
                                if (!String.IsNullOrEmpty(strDisposition))
                                {
                                    KeywordType kwtDisposition = app.Core.KeywordTypes.Find(gSaveToLicType);
                                    if (kwtDisposition != null)
                                        kwdDisposition = CreateKeywordHelper(kwtDisposition, strDisposition);
                                }

                                Keyword kwdCity = null;
                                if (!String.IsNullOrEmpty(strCity))
                                {
                                    KeywordType kwtCity = app.Core.KeywordTypes.Find(gSaveToCity);
                                    if (kwtCity != null)
                                        kwdCity = CreateKeywordHelper(kwtCity, strCity);
                                }

                                Keyword kwdCounty = null;
                                if (!String.IsNullOrEmpty(strCounty))
                                {
                                    KeywordType kwtCounty = app.Core.KeywordTypes.Find(gSaveToCounty);
                                    if (kwtCounty != null)
                                        kwdCounty = CreateKeywordHelper(kwtCounty, strCounty);
                                }

                                Keyword kwdRegion = null;
                                if (!String.IsNullOrEmpty(strRegion))
                                {
                                    KeywordType kwtRegion = app.Core.KeywordTypes.Find(gSaveToRegion);
                                    if (kwtRegion != null)
                                        kwdRegion = CreateKeywordHelper(kwtRegion, strRegion);
                                }

                                Keyword kwdInspector = null;
                                if (!String.IsNullOrEmpty(strInspector))
                                {
                                    KeywordType kwtInspector = app.Core.KeywordTypes.Find(gSaveToInspectorName);
                                    if (kwtInspector != null)
                                        kwdInspector = CreateKeywordHelper(kwtInspector, strInspector);
                                }

                                Keyword kwdInspectorNum = null;
                                if (!String.IsNullOrEmpty(strInspNum))
                                {
                                    KeywordType kwtInspNum = app.Core.KeywordTypes.Find(gSaveToInspNum);
                                    if (kwtInspNum != null)
                                        kwdInspectorNum = CreateKeywordHelper(kwtInspNum, strInspNum);
                                }

                                Keyword kwdInspTypeDesc = null;
                                if (!String.IsNullOrEmpty(strInspType))
                                {
                                    KeywordType kwtInspTypeDesc = app.Core.KeywordTypes.Find(gSaveToInspTypeDesc);
                                    if (kwtInspTypeDesc != null)
                                        kwdInspTypeDesc = CreateKeywordHelper(kwtInspTypeDesc, strInspType);
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
                                    if (kwdLicType != null) keyModifier.UpdateKeyword(kwdLicType, kwdLicType);
                                    if (kwdLicenseNum != null) keyModifier.AddKeyword(kwdLicenseNum);
                                    if (kwdFileNum != null) keyModifier.AddKeyword(kwdFileNum);
                                    if (kwdKeyName != null) keyModifier.AddKeyword(kwdKeyName);
                                    if (kwdDBAName != null) keyModifier.AddKeyword(kwdDBAName);
                                    if (kwdInDorgNum != null) keyModifier.AddKeyword(kwdInDorgNum);
                                    if (kwdVisitNum != null) keyModifier.AddKeyword(kwdVisitNum);
                                    if (kwdDisposition != null) keyModifier.AddKeyword(kwdDisposition);
                                    if (kwdCity != null) keyModifier.AddKeyword(kwdCity);
                                    if (kwdCounty != null) keyModifier.AddKeyword(kwdCounty);
                                    if (kwdRegion != null) keyModifier.AddKeyword(kwdRegion);
                                    if (kwdInspector != null) keyModifier.AddKeyword(kwdInspector);
                                    if (kwdInspTypeDesc != null) keyModifier.AddKeyword(kwdInspTypeDesc);
                                    if (kwdInspectorNum != null) keyModifier.AddKeyword(kwdInspectorNum);
                                    if (kwdSubject != null) keyModifier.AddKeyword(kwdSubject);

                                    // Apply keyword change to the document
                                    keyModifier.ApplyChanges();

                                    documentLock.Release();

                                    string output = String.Format("Keyword: '{0}' Value: '{1}', {33}Keyword: '{2}' Value: '{3}', {33}Keyword: '{4}' Value: '{5}', {33}Keyword: '{6}' Value: '{7}'," +
                                        "{33}Keyword: '{8}' Value: '{9}', {33}Keyword: '{10}' Value: '{11}', {33}Keyword: '{12}' Value: '{13}', {33}Keyword: '{14}' Value: '{15}', {33}Keyword: '{16}' Value: '{17}'," +
                                        "{33}Keyword: '{18}' Value: '{19}', {33}Keyword: '{20}' Value: '{21}', {33}Keyword: '{22}' Value: '{23}', {33}Keyword: '{24}' Value: '{25}', {33}Keyword: '{26}' Value: '{27}'," +
                                        "{33}Keyword: '{28}' Value: '{29}', {33}Keyword: '{30}' Value: '{31}', {33}added to Document {32}.",
                                        gSaveToLicNum, strLicNum, gSaveToLicType, strLicType, gSaveToInspNum, strInspNum, gSaveToFileNum, strFileNum, gSaveToKeyName, strKeyName, gSaveToSubject, strSubject,
                                        gSaveToDBA, strDBAName, gSaveToIndOrgNum, strInDorgNum, gSaveToVisitNum, strVisitNum, gSaveToDisposition, strDisposition, gSaveToCity, strCity, gSaveToCounty, strCounty,
                                        gSaveToRegion, strRegion, gSaveToInspectorName, strInspector, gSaveToInspTypeDesc, strInspType, gSaveToInspNum, strInspectionID, _currentDocument.ID, Environment.NewLine);
                                    //Output the results to the OnBase Diagnostics Console
                                    app.Diagnostics.WriteIf(Hyland.Unity.Diagnostics.DiagnosticsLevel.Verbose, output);
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
