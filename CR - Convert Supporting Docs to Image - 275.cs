using System;
using System.Collections.Generic;
using Hyland.Unity;
using UnityFileConversions;


namespace UnityScripts
{
    /// <summary>
    /// Description
    /// </summary>
    public class DBPR_UnityScript_275 : IWorkflowScript
    {
        #region User-Configurable Script Settings
        // Script name for diagnostics logging
        private const string ScriptName = "275 - CR - Convert Supporting Docs to Image";

        private const string CRD_DOCUMENT_TYPE = "CIU - CR - CAMERA READY DOCUMENT (PRECONV)";
        private const long IMAGE_FILE_TYPE = 2;
        private const long WORD_FILE_TYPE = 12;
        private const long PDF_FILE_TYPE = 16;
        private const string TEMP_DIRECTORY = @"C:\Temp\275\";
        private const string CONVERTER_DIRECTORY = @"C:\Converter\FileConverter.exe";
        private const float PAGE_RATIO = 1.00f;
        private string strFilePath = "";

        #endregion

        /***********************************************
         * USER/SE: PLEASE DO NOT EDIT BELOW THIS LINE *
         ***********************************************/

        #region IWorkflowScript

        /// <summary>
        /// Implementation of <see cref="IWorkflowScript.OnWorkflowScriptExecute" />.
        /// <seealso cref="IWorkflowScript" />
        /// </summary>
        /// <param name="app">Unity Application object</param>
        /// <param name="args">Workflow event arguments</param>
        public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args)
        // public void OnWorkflowScriptExecute(Application app, WorkflowEventArgs args = null)
        {
            app.Diagnostics.Level = Diagnostics.DiagnosticsLevel.Verbose;

            try
            {
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, "Begin Script - Convert Supporting Docs to Image - 275");

                // Get the active document
                Document objCurrentDocument = args.Document;
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, string.Format("Processing document: {0}", objCurrentDocument.ID.ToString()));

                // Get the default rendition
                Rendition objFormRendition = objCurrentDocument.DefaultRenditionOfLatestRevision;

                DocumentType objCrdDocType = app.Core.DocumentTypes.Find(CRD_DOCUMENT_TYPE);

                // Validate the document type
                //DocumentType objCrdDocType = app.Core.DocumentTypes.Find(CRD_DOCUMENT_TYPE);
                if (objCrdDocType == null)
                {
                    throw new InvalidProgramException(string.Format("Document type \"{0}\" does not exist!", CRD_DOCUMENT_TYPE));
                }

                //If the doc is already an image, pull it in its default format and save the new copy
                if (objFormRendition.FileType.ID == IMAGE_FILE_TYPE)
                {
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, "Document is an image, save copy to new Doc Type");

                    DefaultDataProvider objDefaultProvider = app.Core.Retrieval.Default;
                    PageData objDefaultImageData = objDefaultProvider.GetDocument(objFormRendition);

                    StoreNewDocumentProperties objDocProps = app.Core.Storage.CreateStoreNewDocumentProperties(objCrdDocType, app.Core.FileTypes.Find(IMAGE_FILE_TYPE));

                    foreach (KeywordRecord objKeyRecord in objCurrentDocument.KeywordRecords)
                    {
                        if (objKeyRecord.KeywordRecordType.RecordType == RecordType.StandAlone || objKeyRecord.KeywordRecordType.RecordType == RecordType.SingleInstance)
                        {
                            foreach (Keyword objKeyword in objKeyRecord.Keywords)
                            {
                                if (objCrdDocType.KeywordRecordTypes.FindKeywordType(objKeyword.KeywordType.ID) != null)
                                {
                                    objDocProps.AddKeyword(objKeyword);
                                }
                            }
                        }
                        else
                        {
                            EditableKeywordRecord objEditRecord = objKeyRecord.CreateEditableKeywordRecord();
                            objDocProps.AddKeywordRecord(objEditRecord);
                        }
                    }

                    // Store the new document
                    Document objNewDoc = null;
                    objNewDoc = app.Core.Storage.StoreNewDocument(objDefaultImageData, objDocProps);
                    if (objNewDoc == null)
                    {
                        throw new InvalidProgramException("Failed to store new document");
                    }
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, string.Format("Stored new document: {0}", objNewDoc.ID.ToString()));

                    // Clean up
                    objDefaultImageData.Dispose();

                    // If we got here, the script was successful
                    args.ScriptResult = true;
                }
                //If the Doc is Word or Excel, convert it with the conversion utility and save it
                else if (objFormRendition.FileType.ID == WORD_FILE_TYPE || objFormRendition.FileType.ID == PDF_FILE_TYPE)
                {
                    ConversionUtilities unityConverter = new ConversionUtilities(app, TEMP_DIRECTORY, CONVERTER_DIRECTORY);
                    app.Diagnostics.Write("Word/PDF file conversion");

                    List<FileDefinition> files = unityConverter.Convert(args.Document, UnityFileConversions.ImxFileType.Image, UnityFileConversions.ImportType.Document, app.Core.DocumentTypes.Find(CRD_DOCUMENT_TYPE));

                    unityConverter.CleanupFiles();

                    args.ScriptResult = true;
                }
                else
                {
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, "Document is not image/Word/PDF, convert to new Doc Type");

                    ImageDataProvider objImageProvider = app.Core.Retrieval.Image;
                    PageData objImageData = objImageProvider.GetDocument(objFormRendition);

                    StoreNewDocumentProperties objDocProps = app.Core.Storage.CreateStoreNewDocumentProperties(objCrdDocType, app.Core.FileTypes.Find(IMAGE_FILE_TYPE));

                    foreach (KeywordRecord objKeyRecord in objCurrentDocument.KeywordRecords)
                    {
                        if (objKeyRecord.KeywordRecordType.RecordType == RecordType.StandAlone || objKeyRecord.KeywordRecordType.RecordType == RecordType.SingleInstance)
                        {
                            foreach (Keyword objKeyword in objKeyRecord.Keywords)
                            {
                                if (objCrdDocType.KeywordRecordTypes.FindKeywordType(objKeyword.KeywordType.ID) != null)
                                {
                                    objDocProps.AddKeyword(objKeyword);
                                }
                            }
                        }
                        else
                        {
                            EditableKeywordRecord objEditRecord = objKeyRecord.CreateEditableKeywordRecord();
                            objDocProps.AddKeywordRecord(objEditRecord);
                        }
                    }

                    // Store the new document
                    Document objNewDoc = null;
                    objNewDoc = app.Core.Storage.StoreNewDocument(objImageData, objDocProps);
                    if (objNewDoc == null)
                    {
                        throw new InvalidProgramException("Failed to store new document");
                    }
                    app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, string.Format("Stored new document: {0}", objNewDoc.ID.ToString()));

                    // Clean up
                    objImageData.Dispose();

                    // If we got here, the script was successful
                    args.ScriptResult = true;

                }
            }
            catch (InvalidProgramException ex)
            {
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, string.Format("Invalid Program Exception: {0}", ex.Message));
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, string.Format("Stack Trace: {0}", ex.StackTrace));
                args.ScriptResult = false;
            }
            catch (UnityAPIException ex)
            {
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, string.Format("Unity API Exception: {0}", ex.Message));
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, string.Format("Stack Trace: {0}", ex.StackTrace));
                args.ScriptResult = false;
            }
            catch (Exception ex)
            {
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, string.Format("General Exception: {0}", ex.Message));
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Error, string.Format("Stack Trace: {0}", ex.StackTrace));
                args.ScriptResult = false;
            }
            finally
            {
                app.Diagnostics.WriteIf(Diagnostics.DiagnosticsLevel.Info, "End Script - Convert Supporting Doc to Image - 275");
            }
        }
        #endregion

    }
}