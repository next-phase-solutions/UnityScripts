using System;
using Hyland.Unity;
using Hyland.Unity.Workflow;



namespace UnityScripts
{
    /// <summary>
    /// Description
    /// </summary>
    public class DBPR_UnityScript_277 : IWorkflowScript
    {
        #region User-Configurable Script Settings

        // Script name for diagnostics logging
        private const string ScriptName = "277 - CR - Create and update Consolidated Camera Ready doc";


        // Diagnostics logging level - set to Verbose for testing, Error for production
        private const Diagnostics.DiagnosticsLevel DiagLevel = Diagnostics.DiagnosticsLevel.Verbose;

        // Workflow property where error message will be stored
        private const string ErrorMessageProperty = "UnityError";

        // If true, errors will be added to document history
        private const bool WriteErrorsToHistory = true;

        // Date/Time format for diagnostics logging
        private const string DateTimeFormat = "MM-dd-yyyy HH:mm:ss.fff";

        private const string propBoardName = "CILB";

        // Constants for Keyword Types

        // Actual values
        //private const int PAPPNUMBERKWID = 105;
        //private const int PLICTYPEKWID = 104;
        //private const int PDOCNUMBERCONSTKWID = 588;

        private const int PAPPNUMBERKWID = 105;
        private const int PLICTYPEKWID = 104;
        private const int PDOCNUMBERCONSTKWID = 588;

        private const int PBUSINESS_UNIT_KWID = 213;
        private const int PAGENDA_MEETING_MONTH_KWID = 211;
        private const int PAGENDA_MEETING_YEAR_KWID = 212;
        private const int PAGENDA_ADDENDUM_KWID = 210;
        private const int PCAMERAREADYPROCESSEDKWID = 652;
        private const string DOCUMENTCONVERSIONSTATUS = "Conversion Status";

        private const string KNAPPNUM = "Application Number";
        private const string KNLICTYPE = "License Type";
        private const string KNBUSUNIT = "Business Unit";
        private const string KNAGDMON = "Agenda Meeting Month";
        private const string KNAGDYR = "Agenda Meeting Year";
        private const string KNAGDADD = "Agenda Addendum";
        private const string KNDOCNUMBERCONSTKWID = "Document Number - OnBase (Checksheet)";

        // Constants for Document Types

        private const int PCIU_APPLICATIONS_CR_DTID = 1308;
        private const int PCIU_CORRESPONDENCE_CR_DTID = 1313;
        private const int PCIU_CRIMINAL_RESULTS_CR_DTID = 1317;
        private const int PCIU_SUPPORTING_DOCS_CR_DTID = 1320;
        private const int PCIU_CR_CAMERA_READY_DOCUMENT_DTID = 1316;

        private const string PALL_MISSING_DOCS_TO_PROCESS_WITHOUT_REPORTING_ERROR = "propAllCameraReadyDocsProcessed";
        private const string PALL_PROCESSING_ERRORS = "propAllProcessingErrors";

        // Property Bag constants
        private const string PDATAFILE = "propDatafile";
        private const string PWRITEMETHOD = "propWriteMethod";
        private const string PWRITEMETHODHOLD = "propWriteMethodHold";
        private const string PADDENDUM = "propAgendaAddendum";
        private const string PHOLDADD = "propHoldAdd";

        private const string PDOCHANDLE = "propDocHandle";
        private const string PMSTRDOCHANDLE = "propMstrDocHandle";
        private const string PDOCTYPE = "propDocType";

        private const string strLogFile = "WF_SAVE_AS_PROCESS.LOG";

        // OnBase Vars
        private const string gUser = "IAS-SCRIPTACCOUNT";
        private const string gPass = "IASMAGIC";
        private const string gODBC = "OBTEST13";
        private Int64 sourceDocHandle;
        private Int64 destinationDocHandle;

        //        private const string gRootDir = @"D:\Program Files (x86)\Hyland\";
        private const string gRootDir = @"c:\test\DBPR";

        #endregion

        /***********************************************
         * USER/SE: PLEASE DO NOT EDIT BELOW THIS LINE *
         ***********************************************/

        #region Private Globals

        // Active workflow document
        private Document _currentDocument;
        private const long TestDocId = 15186;
        private string strProcessingErrors = "";
        private string ErrorMessage = "";

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
                //int intNumDocsProcessed = 0;
                string strProcessingErrors = "";
                //int gDocCount = 0;

                KeywordType kwtAppNum = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(KNAPPNUM);
                if (kwtAppNum == null)
                    throw new Exception(
                        string.Format(
                            "The Keyword Type '{0}' could not be found or is not on the document type with name of: {1} ",
                            KNAPPNUM, _currentDocument.Name));
                Keyword kwdAppNum = null;
                string strApplicationNumber = "";
                KeywordRecord keyRecAppNum = _currentDocument.KeywordRecords.Find(kwtAppNum);
                if (keyRecAppNum != null)
                {
                    kwdAppNum = keyRecAppNum.Keywords.Find(kwtAppNum);
                    if (kwdAppNum != null) strApplicationNumber = kwdAppNum.ToString();
                }
                KeywordType kwtLicenseType = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(KNLICTYPE);
                if (kwtLicenseType == null)
                    throw new Exception(
                        string.Format(
                            "The Keyword Type '{0}' could not be found or is not on the document type with name of: {1} ",
                            KNLICTYPE, _currentDocument.Name));
                Keyword kwdLicenseType = null;
                string strLicenseType = "";
                KeywordRecord keyRecLicenseType = _currentDocument.KeywordRecords.Find(kwtLicenseType);
                if (keyRecLicenseType != null)
                {
                    kwdLicenseType = keyRecLicenseType.Keywords.Find(kwtLicenseType);
                    if (kwdLicenseType != null) strLicenseType = kwdLicenseType.ToString().Trim();
                }

                KeywordType kwtBusinessUnit = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(KNBUSUNIT);
                if (kwtBusinessUnit == null)
                    throw new Exception(
                        string.Format(
                            "The Keyword Type '{0}' could not be found or is not on the document type with name of: {1} ",
                            KNBUSUNIT, _currentDocument.Name));
                Keyword kwdBusinessUnit = null;
                string strBusinessUnit = "";
                KeywordRecord keyRecBusinessUnit = _currentDocument.KeywordRecords.Find(kwtBusinessUnit);
                if (keyRecBusinessUnit != null)
                {
                    kwdBusinessUnit = keyRecBusinessUnit.Keywords.Find(kwtBusinessUnit);
                    if (kwdBusinessUnit != null) strBusinessUnit = kwdBusinessUnit.ToString().Trim();
                }
                if (strBusinessUnit == "") strBusinessUnit = "TEST";
                KeywordType kwtAgendaMtgMonth =
                    _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(KNAGDMON);
                if (kwtAgendaMtgMonth == null)
                    throw new Exception(
                        string.Format(
                            "The Keyword Type '{0}' could not be found or is not on the document type with name of: {1} ",
                            KNAGDMON, _currentDocument.Name));
                Keyword kwdAgendaMtgMonth = null;
                string strAGMeetingMonth = "";
                KeywordRecord keyRecAgendaMtgMonth = _currentDocument.KeywordRecords.Find(kwtAgendaMtgMonth);
                if (keyRecAgendaMtgMonth != null)
                {
                    kwdAgendaMtgMonth = keyRecAgendaMtgMonth.Keywords.Find(kwtAgendaMtgMonth);
                    if (kwdAgendaMtgMonth != null) strAGMeetingMonth = kwdAgendaMtgMonth.ToString().Trim();
                }
                KeywordType kwtAgendaMtgYear = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(KNAGDYR);
                if (kwtAgendaMtgYear == null)
                    throw new Exception(
                        string.Format(
                            "The Keyword Type '{0}' could not be found or is not on the document type with name of: {1} ",
                            KNAGDYR, _currentDocument.Name));
                Keyword kwdAgendaMtgYear = null;
                string strAGMeetingYear = "";
                KeywordRecord keyRecAgendaMtgYear = _currentDocument.KeywordRecords.Find(kwtAgendaMtgYear);
                if (keyRecAgendaMtgYear != null)
                {
                    kwdAgendaMtgYear = keyRecAgendaMtgYear.Keywords.Find(kwtAgendaMtgYear);
                    if (kwdAgendaMtgYear != null) strAGMeetingYear = kwdAgendaMtgYear.ToString().Trim();

                }
                KeywordType kwtAgendaAddendum =
                    _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(KNAGDADD);
                if (kwtAgendaAddendum == null)
                    throw new Exception(
                        string.Format(
                            "The Keyword Type '{0}' could not be found or is not on the document type with name of: {1} ",
                            KNAGDADD, _currentDocument.Name));
                Keyword kwdAgendaAddendum = null;
                string strAGAddendum = "";
                KeywordRecord keyRecAgendaAddendum = _currentDocument.KeywordRecords.Find(kwtAgendaAddendum);
                if (keyRecAgendaAddendum != null)
                {
                    kwdAgendaAddendum = keyRecAgendaAddendum.Keywords.Find(kwtAgendaAddendum);
                    if (kwdAgendaAddendum != null) strAGAddendum = kwdAgendaAddendum.ToString().Trim();
                }

                string strOkayToProcessWithoutError = "";

                // Get property values
                long lngDocHandle = _currentDocument.ID;

                KeywordType kwtMasterDocHandle = _currentDocument.DocumentType.KeywordRecordTypes.FindKeywordType(KNDOCNUMBERCONSTKWID);
                if (kwtMasterDocHandle == null)
                    throw new Exception(
                        string.Format(
                            "The Keyword Type '{0}' could not be found or is not on the document type with name of: {1} ",
                            KNAPPNUM, _currentDocument.Name));
                Keyword kwdMasterDocHandle = null;
                string lngMstrDocHandle = "";
                KeywordRecord keyRecMasterDocHandle = _currentDocument.KeywordRecords.Find(kwtMasterDocHandle);
                if (keyRecMasterDocHandle != null)
                {
                    kwdMasterDocHandle = keyRecMasterDocHandle.Keywords.Find(kwtMasterDocHandle);
                    if (kwdMasterDocHandle != null) lngMstrDocHandle = kwdMasterDocHandle.ToString();
                }

                app.Diagnostics.Write("Prop Master Doc Handle = " + lngMstrDocHandle);

                bool bOkayToProcessWithoutError = false;
                args.PropertyBag.TryGetValue(PALL_MISSING_DOCS_TO_PROCESS_WITHOUT_REPORTING_ERROR, out strOkayToProcessWithoutError);

                if (!string.IsNullOrEmpty(strOkayToProcessWithoutError))
                {
                    if (strOkayToProcessWithoutError.Trim() == "MISSING DOCS") bOkayToProcessWithoutError = true;
                }

                // Beginning of log file
                WriteLog(app, @"@@@@@@@@@@@@@@@@@@@@@@@@@@ SCRIPT START @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
                WriteLog(app, @"@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
                string strDocType = "CAMERA READY DOCUMENT";
                string strCameraReady = "YES";
                WriteLog(app, "Starting to Consolidating CIU - CR - Camera Ready Document");

                bool bReturn = RunCQ(app, args, strBusinessUnit, strAGMeetingMonth, strAGMeetingYear, strAGAddendum, strDocType, strLicenseType,
                    strApplicationNumber, lngDocHandle, strCameraReady, lngMstrDocHandle, bOkayToProcessWithoutError);
                if (!bReturn) args.ScriptResult = false;

                //WriteLog(app, string.Format("Completed Consolidating CIU - CR - Camera Ready Document.  Number of docs consolidated: {0}", intNumDocsProcessed.ToString().Trim()));

                if (strProcessingErrors.Trim().Length > 0)
                {
                    args.PropertyBag.Set(PALL_PROCESSING_ERRORS, strProcessingErrors);
                }

                //  End Log File
                WriteLog(app, "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
                WriteLog(app, "@@@@@@@@@@@@@@@@@@@@@@@@@@ SCRIPT COMPLETE @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
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
        private void HandleException(Exception ex, ref Application app, ref WorkflowEventArgs args,
            Document otherDocument = null)
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

        private bool WriteLog(Application app, string incoming)
        {
            app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, incoming);
            return true;
        }

        /*
        AppendToDocument() is the function that performs the operations designated for this script.
        This sub will retrieve all of the pages from the source document and
        append them to the destination document.
        NOTE: This script uses IOBXDocumentArchiver and the IOBXFileManager.
        Both objects are required to store a document in this fashion. 
        DocumentArchiver is used to retireve and append pages
        FileManager is used to commit changes to document already stored in OnBase
        */

        private bool AppendToDocument(Application app, WorkflowEventArgs args, Int64 sourceDocHandle, Int64 destinationDocHandle)
        {
            try
            {

                Document destDocument = app.Core.GetDocumentByID(destinationDocHandle);
                if (destDocument == null)
                {
                    WriteLog(app, string.Format("No Destination document found with handle ( {0} ).", destinationDocHandle.ToString()));
                    strProcessingErrors = string.Format("{0}{1}{2}   No Destination document found with handle ( {4} ).", strProcessingErrors, Environment.NewLine, DateTime.Now, destinationDocHandle.ToString());
                    return false;
                }
                Rendition destRendition = destDocument.DefaultRenditionOfLatestRevision;
                ImageDataProvider destImageDataProvider = app.Core.Retrieval.Image;
                PageData destPageData = destImageDataProvider.GetDocument(destRendition);

                Document sourceDocument = app.Core.GetDocumentByID(sourceDocHandle);
                if (sourceDocument == null)
                {
                    WriteLog(app, string.Format("No Source document found with handle ( {0}.", sourceDocHandle.ToString()));
                    strProcessingErrors = String.Format("{0}{1}{2}    No Source document found with handle ({3}).", strProcessingErrors, Environment.NewLine, DateTime.Now, sourceDocHandle.ToString());
                    return false;
                }

                Rendition sourceRendition = sourceDocument.DefaultRenditionOfLatestRevision;
                ImageDataProvider sourceImageDataProvider = app.Core.Retrieval.Image;
                PageData sourcePageData = sourceImageDataProvider.GetDocument(sourceRendition);
                PageRangeSet sourcePageRangeSet = sourceImageDataProvider.CreatePageRangeSet();
                sourcePageRangeSet.AddRange(1, sourceRendition.NumberOfPages);
                PageDataList sourcePageDataList = sourceImageDataProvider.GetPages(sourceRendition, sourcePageRangeSet);

                if (sourcePageDataList == null)
                {
                    WriteLog(app, string.Format("No pages found in Source document  (ID: {0}.", sourceDocHandle.ToString()));
                    return false;
                }
                else
                {
                    WriteLog(app, string.Format("Source Document Handle {0} has {1} pages.", sourceDocHandle.ToString(), sourcePageDataList.Count));
                    // Create a page range object

                    using (DocumentLock documentLock = destDocument.LockDocument())
                    {
                        // Ensure lock was obtained
                        if (documentLock.Status != DocumentLockStatus.LockObtained)
                        {
                            throw new Exception("Document lock not obtained");
                        }

                        long lngPageCount = destDocument.DefaultRenditionOfLatestRevision.NumberOfPages;
                        int intPageLoc = Convert.ToInt32(lngPageCount) + 1;

                        destDocument.DefaultRenditionOfLatestRevision.Imaging.AddPages(sourcePageDataList, intPageLoc);

                    }

                    app.Diagnostics.Write(string.Format("Pages added to document {0}", destinationDocHandle.ToString()));

                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Error in AppendDocument: {0}", ex.Message));
            }
            return true;
        }

        private bool RunCQ(Application app, WorkflowEventArgs args, string strBusinessUnit, string strAGMeetingMonth, string strAGMeetingYear, string strAGAddendum, string strDocType, string strLicenseType, string strApplicationNumber, long lngDocHandle, string strCameraReady, string lngMstrDocHandle, bool bOkayToProcessWithoutError)

        {
            const long CIU_APPLICATION_MSTR_DTID = 773;
            const long CIU_CRIMINAL_RESULTS_MSTR_DTID = 771;
            const long CIU_SUPPORTING_DOCS_MSTR_DTID = 770;
            const long CIU_CORRESPONDENCE_MSTR_DTID = 772;
            const long CIU_CR_CAMERA_READY_DOCUMENT_MSTR_DTID = 2343;

            //String to hold output information
            string output = string.Empty;
            WriteLog(app, string.Format("[{0}][{1}][{2}][{3}][{4}][{5}][{6}][{7}]", strBusinessUnit, strAGMeetingYear, strAGMeetingMonth, strAGAddendum, strDocType, strLicenseType, strApplicationNumber, lngDocHandle.ToString(), strCameraReady));

            long gDocTypeID;
            long lngCurrentMSTRDTID;

            switch (strDocType)
            {
                case "CAMERA READY DOCUMENT":
                    {
                        gDocTypeID = PCIU_CR_CAMERA_READY_DOCUMENT_DTID;
                        lngCurrentMSTRDTID = CIU_CR_CAMERA_READY_DOCUMENT_MSTR_DTID;
                        break;
                    }
                case "APPLICATION":
                    {
                        gDocTypeID = PCIU_APPLICATIONS_CR_DTID;
                        lngCurrentMSTRDTID = CIU_APPLICATION_MSTR_DTID;
                        break;
                    }
                case "CRIMINAL RESULTS":
                    {
                        gDocTypeID = PCIU_CRIMINAL_RESULTS_CR_DTID;
                        lngCurrentMSTRDTID = CIU_CRIMINAL_RESULTS_MSTR_DTID;
                        break;
                    }
                case "SUPPORTING DOCS":
                    {
                        gDocTypeID = PCIU_SUPPORTING_DOCS_CR_DTID;
                        lngCurrentMSTRDTID = CIU_SUPPORTING_DOCS_MSTR_DTID;
                        break;
                    }
                case "CORRESPONDENCE":
                    {
                        gDocTypeID = PCIU_CORRESPONDENCE_CR_DTID;
                        lngCurrentMSTRDTID = CIU_CORRESPONDENCE_MSTR_DTID;
                        break;
                    }
                default:
                    {
                        gDocTypeID = 0;
                        lngCurrentMSTRDTID = 0;
                        break;
                    }
            }


            // Create document query
            WriteLog(app, string.Format("Setting up Document Query [{0} - {1}]", strDocType, gDocTypeID.ToString()));
            DocumentQuery documentQuery = app.Core.CreateDocumentQuery();

            // Ensure custom query was found
            if (documentQuery == null)
            {
                throw new Exception("Unable to create document query");
            }

            // Add custom query to document query
            DocumentType docType = app.Core.DocumentTypes.Find(gDocTypeID);

            if (docType == null)
            {
                throw new Exception(string.Format("Could not find document type with ID: {0}", gDocTypeID.ToString()));
            }
            documentQuery.AddDocumentType(docType);

            KeywordType kwdType = app.Core.KeywordTypes.Find(PAPPNUMBERKWID);
            if (kwdType == null)
            {
                throw new Exception(string.Format("Could not find keyword type with ID: {0}", PAPPNUMBERKWID.ToString()));
            }
            documentQuery.AddKeyword(kwdType.Name, Convert.ToInt64(strApplicationNumber));

            kwdType = app.Core.KeywordTypes.Find(PLICTYPEKWID);
            if (kwdType == null)
            {
                throw new Exception(string.Format("Could not find keyword type with ID: {0}", PLICTYPEKWID.ToString()));
            }
            documentQuery.AddKeyword(kwdType.Name, strLicenseType);

            kwdType = app.Core.KeywordTypes.Find(PDOCNUMBERCONSTKWID);
            if (kwdType == null)
            {
                throw new Exception(string.Format("Could not find keyword type with ID: {0}", PDOCNUMBERCONSTKWID.ToString()));
            }
            documentQuery.AddKeyword(kwdType.Name, Convert.ToInt64(lngMstrDocHandle));

            documentQuery.AddSort(DocumentQuery.SortAttribute.DocumentID, true);

            WriteLog(app, string.Format("Running Document Query - Search KW [{0} = {1}]", PAPPNUMBERKWID.ToString(), strApplicationNumber));
            WriteLog(app, string.Format("Running Document Query - Search KW [{0} = {1}]", KNLICTYPE, strLicenseType));
            WriteLog(app, string.Format("Running Document Query - Search KW [{0} = {1}]", KNDOCNUMBERCONSTKWID, lngMstrDocHandle.ToString()));

            // Execute query
            const int MAX_DOCUMENTS = 1000;
            DocumentList documentList = documentQuery.Execute(MAX_DOCUMENTS);

            if (documentList.Count == 0)
            {
                if (!bOkayToProcessWithoutError)
                {
                    WriteLog(app, "  ");
                    WriteLog(app, "*************** ERROR: NO DOCUMENTS FOUND IN SEARCH **************");
                    WriteLog(app, "  ");
                    strProcessingErrors = string.Format("{0}{1}  Failed to locate any documents of type - {2}.  AN: {3} - LT: {4}", strProcessingErrors, DateTime.Now, strDocType, strApplicationNumber, strLicenseType);

                    return false;
                }
                else
                {
                    WriteLog(app, "  ");
                    WriteLog(app, "*************** ERROR BYPASSED: USER STATED TO PROCESS WITH MISSING DOCUMENTS **************");
                    WriteLog(app, "  ");
                    WriteLog(app, "TEST CAMERA READY");
                    WriteLog(app, "  ");
                    return true;
                }
            }
            // Line 298 from VB script.
            // Iterate through documents returned from query
            bool bIsFirstDocument = true;
            string ProcessedDocIDs = "";
            //string ProcessedDoc = "";
            //string strDocID = "";
            string laststrPrimaryDocID = "0";
            //int intNumDocsProcessed = 0;
            //int gDocCount = 0;

            foreach (Document document in documentList)
            {
                if (bIsFirstDocument)
                {
                    DocumentType objDocType = document.DocumentType;
                    WriteLog(app, "  ");
                    WriteLog(app, "*************** SETTING DOCUMENT TYPE TO IT MASTER EQUIVALENT **************");
                    WriteLog(app, "  ");
                    WriteLog(app, String.Format("*************** CURRENT DOCUMENT TYPE : {0} **************", objDocType.Name.ToString().Trim()));
                    WriteLog(app, String.Format("*************** CURRENT DOCUMENT TYPE : {0} **************", objDocType.ID.ToString().Trim()));

                    DocumentType newdoctype = app.Core.DocumentTypes.Find(lngCurrentMSTRDTID);

                    if (newdoctype == null)
                    {
                        throw new ApplicationException(string.Format("Document Type {0} does not exist", lngCurrentMSTRDTID));
                    }
                    Storage storage = app.Core.Storage;
                    ReindexProperties reindexProperties = storage.CreateReindexProperties(document, newdoctype);
                    Document newDocument = storage.ReindexDocument(reindexProperties);

                    WriteLog(app, String.Format("*************** UPDATED DOCUMENT TYPE : {0} **************", newdoctype.Name));
                    WriteLog(app, String.Format("*************** UPDATED DOCUMENT TYPE ID (Should be: {0} : {1} **************", lngCurrentMSTRDTID.ToString(), document.ID.ToString()));

                    // Update the document AutoName
                    //WriteLog(app,"*************** UPDATED DOCUMENT AUTO NAME: " & Trim(CStr(objDoc.Name)) & " **************")
                    //objDoc.AutoName  

                }
                app.Diagnostics.Write(string.Format("Document ID: {0} Document Name: {1}{2}", document.ID, document.Name, Environment.NewLine));
                //Line 349
                //int x = 1;
                // If there are multiple instances of a KW value on a document, we do not want to output the same doc again
                // as the result set will contain an document entry for each unique KW value of the same type.
                if (ProcessedDocIDs.Contains(document.ID.ToString().Trim()) || (document.ID.ToString() == laststrPrimaryDocID.Trim()))
                    app.Diagnostics.Write("Complete");
                else
                {
                    //intNumDocsProcessed++;
                    ProcessedDocIDs = string.Format("{0}{1}", ProcessedDocIDs, document.ID.ToString().Trim());
                    //gDocCount++;
                    string fileTypeExt = "";

                    long fileTypeID = 2;
                    Rendition objFormRendition = document.DefaultRenditionOfLatestRevision;
                    fileTypeID = objFormRendition.FileType.ID;

                    switch (fileTypeID)
                    {
                        case 1:
                            fileTypeExt = "txt";
                            break;

                        case 2:
                            fileTypeExt = "tif";
                            break;

                        case 16:
                            fileTypeExt = "pdf";
                            break;

                        case 17:
                            fileTypeExt = "htm";
                            break;

                        case 13:
                            fileTypeExt = "xls";
                            break;

                        case 12:
                            fileTypeExt = "doc";
                            break;

                        case 14:
                            fileTypeExt = "ppt";
                            break;

                        case 15:
                            fileTypeExt = "rtf";
                            break;

                        case 24:
                            fileTypeExt = "htm";
                            break;

                        case 32:
                            fileTypeExt = "xml";
                            break;

                        default:
                            fileTypeExt = "unknown";
                            break;

                    }

                    KeywordType kwtDocConversionStatus = null;
                    Keyword kwdDocConversionStatus = null;
                    if (fileTypeExt == "tif")
                    {
                        kwtDocConversionStatus = app.Core.KeywordTypes.Find(DOCUMENTCONVERSIONSTATUS);
                        if (kwtDocConversionStatus == null)
                            throw new Exception(String.Format("Keyword Type '{0}' not found", DOCUMENTCONVERSIONSTATUS));
                        kwdDocConversionStatus = CreateKeywordHelper(kwtDocConversionStatus, "CONV OK");
                        WriteLog(app, "Document file format is TIFF");
                    }
                    else
                    {
                        kwtDocConversionStatus = app.Core.KeywordTypes.Find(DOCUMENTCONVERSIONSTATUS);
                        if (kwtDocConversionStatus == null)
                            throw new Exception(String.Format("Keyword Type '{0}' not found", DOCUMENTCONVERSIONSTATUS));
                        kwdDocConversionStatus = CreateKeywordHelper(kwtDocConversionStatus, "NOT IMAGE FORMAT");
                        WriteLog(app, "Document file format is NOT TIFF - need to add logic to convert to TIFF");
                        strProcessingErrors = string.Format("{0}{1}{2}   " +
                                                            "Document is not Image File Format - {3}.  AN: {4} - LT: {5}", strProcessingErrors, Environment.NewLine, DateTime.Now.ToString(DateTimeFormat), strDocType, strApplicationNumber, strLicenseType);
                    }

                    using (DocumentLock documentLock = document.LockDocument())
                    {
                        // Ensure lock was obtained
                        if (documentLock.Status != DocumentLockStatus.LockObtained)
                        {
                            throw new Exception("Document lock not obtained");
                        }
                        // Create keyword modifier object to hold keyword changes
                        KeywordModifier keyModifier = document.CreateKeywordModifier();

                        //Add update keyword call to keyword modifier object
                        //Note Overloads available for use
                        //(I.E.): keyModifier.AddKeyword(keywordTypeName,keywordValue)
                        keyModifier.AddKeyword(kwdDocConversionStatus);

                        // Apply keyword change to the document
                        keyModifier.ApplyChanges();

                        WriteLog(app, string.Format("Keyword: '{0}' added to Document .", DOCUMENTCONVERSIONSTATUS));
                    }
                    if (fileTypeExt != "tif")
                        return false;

                    WriteLog(app, string.Format("  Processing DocID [{0}]", document.ID.ToString()));
                    // Set property bag for use by OnBase DocDataProvider object
                    args.PropertyBag.Set("docID", document.ID.ToString());

                    sourceDocHandle = document.ID;

                    // If the source file format is not image, we need to save it out and append to the current document.  
                    // Otherwise, we need to append the current image file in the diskgroup
                    // If currentDocumentFileFormatIsImage Then
                    if (bIsFirstDocument)
                    {
                        //app.Diagnostics.Write("in first doc logic");
                        destinationDocHandle = sourceDocHandle;
                        bIsFirstDocument = false;
                    }
                    else
                    {
                        //app.Diagnostics.Write("Made it in");
                        AppendToDocument(app, args, sourceDocHandle, destinationDocHandle);
                    }
                }
                laststrPrimaryDocID = document.ID.ToString();
            }
            return true;
        }

        #endregion
    }
}