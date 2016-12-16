// <copyright file="IssuanceManager.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Classes
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using Aspose.Words;
    using Aspose.Words.Drawing;
    using Aspose.Words.Tables;
    using CorrespondenceServices.Interfaces;
    using DecisionModel.Models.Policy;
    using DecisionModel.Representations;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.DocumentGenerator.Functions;
    using Mkl.WebTeam.DocumentGenerator.Helpers;
    using Mkl.WebTeam.RestHelper.Enumerations;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.StorageProvider.Models;
    using Mkl.WebTeam.SubmissionShared.Enumerations;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Issuance Manager
    /// </summary>
    public class IssuanceManager : IIssuanceManager
    {
        /// <summary>
        /// The format methods
        /// </summary>
        private static Dictionary<string, MethodInfo> formatMethods = new Dictionary<string, MethodInfo>();

        /// <summary>
        /// The rest helper
        /// </summary>
        private readonly IRestHelper restHelper;

        /// <summary>
        /// Template processor
        /// </summary>
        private readonly ITemplateProcessor templateProcessor;

        /// <summary>
        /// Storage manager
        /// </summary>
        private readonly IStorageManager storageManager;

        /// <summary>
        /// Policy carrier
        /// </summary>
        private readonly Carrier policyCarrier;

        /// <summary>
        /// Initializes a new instance of the <see cref="IssuanceManager" /> class.
        /// </summary>
        /// <param name="restHelper">The rest helper.</param>
        /// <param name="policy">Policy</param>
        /// <param name="policyCarrier">Policy carrier</param>
        /// <param name="templateProcessor">ITemplateProcessor</param>
        /// <param name="storageManager">IStorageManager</param>
        public IssuanceManager(IRestHelper restHelper, Policy policy, Carrier policyCarrier, ITemplateProcessor templateProcessor, IStorageManager storageManager)
        {
            this.restHelper = restHelper;
            this.templateProcessor = templateProcessor;
            this.storageManager = storageManager;
            this.Policy = policy;
            this.policyCarrier = policyCarrier;

            if (policy != null)
            {
                this.PolicyJson = JObject.Parse(JsonConvert.SerializeObject(this.Policy));
            }
        }

        /// <summary>
        /// Gets or sets the policy.
        /// </summary>
        /// <value>The policy.</value>
        private Policy Policy { get; set; }

        /// <summary>
        /// Gets or sets the policy json.
        /// </summary>
        /// <value>The policy json.</value>
        private JObject PolicyJson { get; set; }

        /// <summary>
        /// Get the generated documents out of merged forms and template mappings
        /// </summary>
        /// <param name="authenticatedUser">User</param>
        /// <param name="mergedForms">list of merged forms</param>
        /// <param name="effectiveDate">DateTime</param>
        /// <param name="fileFormat">fileFormat</param>
        /// <param name="removeDocs">bool</param>
        /// <returns>MergeDocumentsResponse</returns>
        public MergeDocumentsResponse ProcessMergedFormTemplates(IUser authenticatedUser, List<MergedForm> mergedForms, DateTime? effectiveDate, DocumentManager.FileFormat fileFormat, bool removeDocs)
        {
            MergeDocumentsResponse mergeDocumentsResponse = this.RemovePolicyIssuanceDocuments(!removeDocs);

            this.CreateDefaultTemplateMapping(authenticatedUser, effectiveDate);

            foreach (var mergedForm in mergedForms)
            {
                try
                {
                    string templateMappingId;

                    var docTemplateBinaryData = this.GetDocTemplateBinaryData(authenticatedUser, mergedForm.DocumentId.ToString(), effectiveDate, out templateMappingId);

                    //if (mergedForm.FormNormalizedNumber.EqualsIgnoreCase("MECP1281"))
                    //{
                    //    docTemplateBinaryData = File.ReadAllBytes(@"E:\\Document\\MECP 1281 01 15.docx");
                    //}
                    //else if (mergedForm.FormNormalizedNumber.EqualsIgnoreCase("MECP1321"))
                    //{
                    //    docTemplateBinaryData = File.ReadAllBytes(@"E:\\Document\\MECP 1321 02 16.docx");
                    //}

                    if (mergedForm.FormNormalizedNumber.EqualsIgnoreCase("MECP1200"))
                    {
                        //docTemplateBinaryData = File.ReadAllBytes(@"E:\\Document\\CG 21 16 04 13.docx");
                    }

                    if (docTemplateBinaryData == null || docTemplateBinaryData.Length <= 0)
                    {
                        mergeDocumentsResponse.UploadedFilePathAndName.Add(new MergedForm()
                        {
                            FormNormalizedNumber = mergedForm.FormNormalizedNumber,
                            ErrorMessages = new List<string>() { "Document template binary data not found for : " + mergedForm.FormNormalizedNumber }
                        });

                        continue;
                    }

                    var instanceCount = this.GetNumberOfFormInstancesRequired(mergedForm);
                    Stream docTemplateStream = new MemoryStream(docTemplateBinaryData);

                    var customMappings = this.GetCustomMappingTemplate(authenticatedUser, templateMappingId, effectiveDate);
                    var stampWaterMark = this.Policy.WorkflowState == WorkflowState.PreIssue;
                    var docManager = this.GenerateMergedTemplateDocument(mergedForm, docTemplateStream, this.Policy.SubmissionNumber, customMappings, authenticatedUser, fileFormat, stampWaterMark, instanceCount);
                    var response = this.UploadDocumentToStore(ref docManager, mergedForm.FormNormalizedNumber, mergedForm.QuantityOrder, fileFormat);

                    if (!string.IsNullOrWhiteSpace(response?.DirectoryPath) &&
                        !string.IsNullOrWhiteSpace(response.FileName))
                    {
                        mergeDocumentsResponse.UploadedFilePathAndName.Add(new MergedForm()
                        {
                            FormNormalizedNumber = mergedForm.FormNormalizedNumber,
                            QuantityOrder = mergedForm.QuantityOrder,
                            CompressionLevel = 0,
                            Name = response.FileName,
                            Path = response.DirectoryPath,
                            Templated = false
                        });
                    }
                    else
                    {
                        mergeDocumentsResponse.UploadedFilePathAndName.Add(new MergedForm()
                        {
                            FormNormalizedNumber = mergedForm.FormNormalizedNumber,
                            ErrorMessages = new List<string>() { "The document was not saved." }
                        });
                    }
                }
                catch (Exception ae)
                {
                    mergeDocumentsResponse.UploadedFilePathAndName.Add(new MergedForm()
                    {
                        FormNormalizedNumber = mergedForm.FormNormalizedNumber,
                        QuantityOrder = mergedForm.QuantityOrder,
                        ErrorMessages = new List<string>() { ae.Message }
                    });
                    continue;
                }
            }

            return mergeDocumentsResponse;
        }

        /// <summary>
        /// Create default template mapping
        /// </summary>
        /// <param name="authenticatedUser">User</param>
        /// <param name="effectiveDate">DateTime</param>
        public void CreateDefaultTemplateMapping(IUser authenticatedUser, DateTime? effectiveDate)
        {
            JToken tokenTemplate = this.GetDefaultMappingTemplate(authenticatedUser, effectiveDate);
            this.templateProcessor.CreateDefaultDictionary(tokenTemplate);
        }

        /// <summary>
        /// Get Signature Bytes
        /// </summary>
        /// <param name="authenticatedUser">user</param>
        /// <param name="signaturetype">CorporateSignatureType</param>
        /// <param name="datetime">Effective Date of the policy</param>
        /// <param name="stateCode">State code of the producer</param>
        /// <returns>byte[]</returns>
        public byte[] GetSignatureBytes(IUser authenticatedUser, CorporateSignatureType signaturetype, DateTime? datetime, string stateCode = "")
        {
            byte[] contents = null;
            var fileName = this.GetSignatureFileName(authenticatedUser, signaturetype, datetime, stateCode);

            // Call Storage Provider
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                contents = this.storageManager.GetCorporateSignatureBinary(fileName);
            }

            return contents;
        }

        /// <summary>
        /// Get Signature File Name
        /// </summary>
        /// <param name="authenticatedUser">user</param>
        /// <param name="signatureType">CorporateSignatureType</param>
        /// <param name="effectiveDate">Effective Date of the policy</param>
        /// <param name="stateCode">State code of the producer</param>
        /// <returns>string</returns>
        public string GetSignatureFileName(IUser authenticatedUser, CorporateSignatureType signatureType, DateTime? effectiveDate, string stateCode = "")
        {
            JToken responses;

            if (CorporateSignatureType.President.Equals(signatureType))
            {
                responses = this.restHelper.GetServiceResult("GetPresidentSignature?stateCode=" + stateCode, RestSharp.Method.GET, authenticatedUser, null, WebTeamServiceEndpoint.BaseServices);
            }
            else
            {
                responses = this.restHelper.GetServiceResult("GetSecretarySignature?effectiveDate=" + effectiveDate.Value.ToShortDateString(), RestSharp.Method.GET, authenticatedUser, null, WebTeamServiceEndpoint.BaseServices);
            }

            var reader = responses.CreateReader();
            string fileName = reader.ReadAsString();

            return fileName;
        }

        /// <summary>
        /// Consolidate all the issuance documents of the policy into one document and upload it to storage provider.
        /// </summary>
        /// <returns>Document Upload response</returns>
        public UploadDocumentResponse GenerateConsolidatedPolicyDocument()
        {
            ////Get ordered list of documents
            var documentPaths = new List<string>();
            var errors = new List<string>();

            ////Get policy forms lob vise
            var documents = this.GetOrderedPolicyForms(this.Policy);

            foreach (var key in documents.Keys.OrderBy(o => o.Order))
            {
                foreach (var document in ((IEnumerable<DecisionModel.Models.Policy.Document>)documents[key].Values).OrderBy(d => d.Order == 0 ? 99999 : d.Order).ThenBy(d => d.NormalizedNumber).ThenBy(d => d.EditionDate).ThenBy(d => d.QuantityOrder))
                {
                    var policyDocument = this.Policy.RollUp.Documents.FirstOrDefault(x => x.NormalizedNumber.EqualsIgnoreCase(document.NormalizedNumber) && x.EditionDate == document.EditionDate && x.QuantityOrder == document.QuantityOrder);

                    if (policyDocument != null)
                    {
                        if (!(string.IsNullOrWhiteSpace(policyDocument.FilePath) || string.IsNullOrWhiteSpace(policyDocument.FileName)))
                        {
                            var filePath = Path.Combine(policyDocument.FilePath, policyDocument.FileName);
                            if (File.Exists(filePath))
                            {
                                documentPaths.Add(filePath);
                            }
                            else
                            {
                                errors.Add(string.Format("{0} : File {1} does not exist.", policyDocument.DisplayNumber, filePath));
                            }
                        }
                        else
                        {
                            errors.Add(string.Format("{0} : FilePath/FileName is missing.", policyDocument.DisplayNumber));
                        }
                    }
                    else
                    {
                        errors.Add(string.Format("{0} : Line document missing in policy RollUp.", document.DisplayNumber));
                    }
                }
            }

            ////If there were any errors return errors
            if (errors.Any())
            {
                return new UploadDocumentResponse
                {
                    Errors = errors
                };
            }

            ////Generate policy document and upload to storage provider if there are no errors
            var fileName = this.GetPolicyDocumentName();
            var response = this.CombineAllDocumentsAndUploadToStore(documentPaths.ToArray(), fileName);

            return response;
        }

        /// <summary>
        /// RemovePolicyIssuanceDocuments - Method removes any already generated documents.
        /// </summary>
        /// <param name="removeDirty">bool</param>
        /// <returns>MergeDocumentsResponse</returns>
        public MergeDocumentsResponse RemovePolicyIssuanceDocuments(bool removeDirty = false)
        {
            MergeDocumentsResponse response = new MergeDocumentsResponse();
            response.UploadedFilePathAndName = new List<MergedForm>();

            List<DecisionModel.Models.Policy.Document> documents = new List<DecisionModel.Models.Policy.Document>();
            if (removeDirty)
            {
                documents = this.Policy.RollUp.Documents.Where(x => x.IsGeneratedDocumentDirty == true).ToList();
            }
            else
            {
                documents = this.Policy.RollUp.Documents.Select(x => x).ToList();
            }

            foreach (var document in documents)
            {
                if (!(string.IsNullOrWhiteSpace(document.FilePath) || string.IsNullOrWhiteSpace(document.FileName)))
                {
                    string filePath = Path.Combine(document.FilePath, document.FileName);
                    try
                    {
                        if (File.Exists(filePath))
                        {
                            File.Delete(filePath);
                        }
                    }
                    catch (Exception)
                    {
                        response.UploadedFilePathAndName.Add(new MergedForm()
                        {
                            FormNormalizedNumber = document.NormalizedNumber,
                            ErrorMessages = new List<string>() { $"Unable to delete policy issuance document - {filePath}" }
                        });
                    }
                }
            }

            return response;
        }

        /// <summary>
        /// Combine specified documents to the current document.
        /// </summary>
        /// <param name="sourceDocuments">Path of documents to be merged.</param>
        /// <param name="fileName">File name</param>
        /// <returns>Document Upload response</returns>
        public UploadDocumentResponse CombineAllDocumentsAndUploadToStore(string[] sourceDocuments, string fileName)
        {
            MemoryStream memoryStream = new MemoryStream();
            UploadDocumentResponse response;
            var errors = new List<string>();

            try
            {
                //// Open the output document
                PdfSharp.Pdf.PdfDocument outputDocument = new PdfSharp.Pdf.PdfDocument();

                //// Iterate files
                foreach (string file in sourceDocuments)
                {
                    // Open the document to import pages from it.
                    PdfSharp.Pdf.PdfDocument inputDocument = PdfSharp.Pdf.IO.PdfReader.Open(file, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);

                    //// Iterate pages
                    int count = inputDocument.PageCount;

                    for (int idx = 0; idx < count; idx++)
                    {
                        //// Get the page from the external document...
                        PdfSharp.Pdf.PdfPage page = inputDocument.Pages[idx];
                        //// ...and add it to the output document.
                        outputDocument.AddPage(page);
                    }
                }

                //// Save the document...
                outputDocument.Save(memoryStream, false);
                response = this.UploadDocumentToStore(memoryStream, fileName, DocumentManager.FileFormat.PDF);

                ////Check if the consolidated policy document was saved
                if (
                    string.IsNullOrWhiteSpace(response.DirectoryPath) ||
                    string.IsNullOrWhiteSpace(response.FileName) ||
                   !File.Exists(Path.Combine(response.DirectoryPath, response.FileName)))
                {
                    if (response.Errors == null)
                    {
                        response.Errors = new List<string>();
                    }

                    response.Errors.Add("Consolidated Policy document was not saved.");
                }
            }
            catch (Exception ex)
            {
                errors.Add(string.Format("Error occured while consolidating policy document.\r\n{0}", ex.Message));
                response = new UploadDocumentResponse
                {
                    Errors = errors
                };
            }
            finally
            {
                if (memoryStream != null)
                {
                    memoryStream.Flush();
                    memoryStream.Dispose();
                }
            }

            return response;
        }

        /// <summary>
        /// Gets the working question.
        /// </summary>
        /// <param name="question">The question.</param>
        /// <param name="formFromRollUp">The form from roll up.</param>
        /// <returns>Question.</returns>
        private static Question GetWorkingQuestion(Question question, DecisionModel.Models.Policy.Document formFromRollUp)
        {
            var replacementQuestion = formFromRollUp
                .Questions
                .FirstOrDefault(rq =>

                        // The current question's code does not match the replacement questions code (same question, circular reference)
                        question.Code != rq.Code &&

                        // The current question and the replacement question are contained in the same question grouping
                        question.MultipleRowGroupingNumber == rq.MultipleRowGroupingNumber &&

                        // AND the current questions code matches the replacement controlling question code
                        question.Code == rq.ControllingQuestionCode &&

                        // AND (the current question's answer matches the replacement questions answer code replacement value
                        (string.Compare(question.AnswerValue, rq.AnswerCodeReplacement, StringComparison.InvariantCultureIgnoreCase) == 0 ||

                        // OR the replacement question has the questions selected answer code as the replacement answer code)
                        (question.Answers != null &&
                        question.Answers.Any(a => a.IsSelected &&
                                                    a.Value == rq.AnswerCodeReplacement))));
            return replacementQuestion ?? question;
        }

        /// <summary>
        /// Check if the bookmark has already been processed, if so exit, otherwise add to cache and replace bookmark value
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="docManager">The document manager.</param>
        /// <param name="workingQuestion">The working question.</param>
        /// <param name="handledMergeFields">The handled merge fields.</param>
        /// <param name="policy">The policy.</param>
        /// <param name="formatMethodInfo">The format method information.</param>
        private static void HandleMergeField(string bookmarkName, IPolicyDocumentManager docManager, Question workingQuestion, HashSet<string> handledMergeFields, Policy policy, MethodInfo formatMethodInfo = null)
        {
            var hashKey = bookmarkName.ToLowerInvariant();
            if (handledMergeFields.Contains(hashKey))
            {
                // We have already handled this merge field name and do not need to do so again.
                return;
            }

            // Add the item to the hash list so it doesnt get processed again
            handledMergeFields.Add(hashKey);

            var bookmark = docManager.GetBookMarkByName(bookmarkName);
            var nodeValue = docManager.GetQuestionValue(bookmarkName, workingQuestion);
            if (formatMethodInfo != null)
            {
                nodeValue = ExecuteFormatMethod(formatMethodInfo, nodeValue, docManager, workingQuestion, bookmark, bookmarkName, policy);
            }

            // Handle bookmarks configured as TextBox type Q&A
            if (!string.IsNullOrWhiteSpace(nodeValue))
            {
                docManager.ReplaceNodeValue(bookmark, nodeValue);
            }

            ////docManager.ReplaceQuestionBookmarkValue(bookmarkName, workingQuestion);
        }

        /// <summary>
        /// Check if the bookmark has already been processed, if so exit, otherwise add to cache and replace bookmark value
        /// </summary>
        /// <param name="instance">The instance of the document to update.</param>
        /// <param name="itemInstance">The item instance being manipulated.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="docManager">The document manager.</param>
        /// <param name="workingQuestion">The working question.</param>
        /// <param name="handledMergeFields">The handled merge fields.</param>
        /// <param name="policy">The policy.</param>
        /// <param name="formatMethodInfo">Format function applied to the text value as a MethodInfo</param>
        private static void HandleMultipleMergeField(int instance, int itemInstance, string bookmarkName, IPolicyDocumentManager docManager, Question workingQuestion, HashSet<Tuple<int, string>> handledMergeFields, Policy policy, MethodInfo formatMethodInfo = null)
        {
            var bookmarkNames = bookmarkName.Split(',');
            string handledName = bookmarkNames.Length > itemInstance ? bookmarkNames[itemInstance] : bookmarkNames[0];

            var tupleKey = new Tuple<int, string>(instance, handledName.ToLowerInvariant());
            if (handledMergeFields.Contains(tupleKey))
            {
                // We have already handled this merge field name and do not need to do so again.
                return;
            }

            // Add the item to the hash list so it doesn't get processed again
            handledMergeFields.Add(tupleKey);

            var bookmark = docManager.GetBookMarkByName(handledName);
            var nodeValue = docManager.GetQuestionValue(handledName, workingQuestion);
            if (formatMethodInfo != null)
            {
                nodeValue = ExecuteFormatMethod(formatMethodInfo, nodeValue, docManager, workingQuestion, bookmark, handledName, policy);
            }

            // Handle bookmarks configured as TextBox type Q&A
            if (!string.IsNullOrWhiteSpace(nodeValue))
            {
                docManager.ReplaceNodeValue(bookmark, nodeValue);
            }

            ////docManager.ReplaceQuestionBookmarkValue(handledName, workingQuestion, null, formatMethodInfo);
        }

        /// <summary>
        /// Executes the format method.
        /// </summary>
        /// <param name="methodInfo">The method information.</param>
        /// <param name="value">The value.</param>
        /// <param name="manager">The manager.</param>
        /// <param name="question">The question.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="policy">The policy.</param>
        /// <returns>System.String.</returns>
        private static string ExecuteFormatMethod(MethodInfo methodInfo, string value, IPolicyDocumentManager manager, Question question, Bookmark bookmark, string bookmarkName, Policy policy)
        {
            if (methodInfo == null)
            {
                return value;
            }

            object classInstance = Activator.CreateInstance(typeof(CustomFunctions), null);
            var parameters = new object[] { value, manager, question, bookmark, bookmarkName, policy };
            var result = methodInfo.Invoke(classInstance, parameters);

            return result as string;
        }

        /// <summary>
        /// Gets the number of form instances required.
        /// </summary>
        /// <param name="mergedForm">The merged form.</param>
        /// <returns>System.Int32.</returns>
        private int GetNumberOfFormInstancesRequired(MergedForm mergedForm)
        {
            // We have a form, now check the policy questions and answers for that form and determine if we have the possibility or multiple
            // copies of the form.  If we do then we need to generate all copies as we go.
            var formQuestions = this.Policy
                .RollUp
                .Documents
                .FirstOrDefault(x => x.NormalizedNumber.EqualsIgnoreCase(mergedForm.FormNormalizedNumber) &&
                                        x.QuantityOrder.GetValueOrDefault(0) == mergedForm.QuantityOrder.GetValueOrDefault())?.Questions;

            if (formQuestions == null || !formQuestions.Any())
            {
                return 1;
            }

            var formQuestionsWithMaximumMultipleRowCount = formQuestions
                .Where(q => q.MaximumMultipleRowCount.GetValueOrDefault(0) > 0)
                .Select(q => q.Code)
                .Distinct()
                .ToList();

            if (!formQuestionsWithMaximumMultipleRowCount.Any())
            {
                return 1;
            }

            var maxRowCounter = formQuestionsWithMaximumMultipleRowCount
                .Select(questionCode => new
                {
                    Code = questionCode,
                    MaxRowCount = (formQuestions.FirstOrDefault(q => q.Code == questionCode)?.MaximumMultipleRowCount).GetValueOrDefault(0)
                })
                .Where(item => item.MaxRowCount != 0)
                .ToList();
            if (!maxRowCounter.Any())
            {
                return 1;
            }

            var questionCounter = maxRowCounter
                .Select(item => new
                {
                    item.Code,
                    item.MaxRowCount,
                    QuestionCount = formQuestions.Count(q => q.Code == item.Code)
                })
                .Where(questionCount => questionCount.QuestionCount > questionCount.MaxRowCount)
                .ToList();
            if (!questionCounter.Any())
            {
                return 1;
            }

            var instanceCounter = questionCounter
                .Select(item => new
                {
                    item.Code,
                    item.MaxRowCount,
                    item.QuestionCount,
                    ////Quotient = item.QuestionCount / item.MaxRowCount,
                    ////Remainder = item.QuestionCount % item.MaxRowCount,
                    InstanceCount = (item.QuestionCount / item.MaxRowCount) + ((item.QuestionCount % item.MaxRowCount) == 0 ? 0 : 1)
                }).ToList();
            if (!instanceCounter.Any())
            {
                return 1;
            }

            var maxInstanceCount = instanceCounter
                .Max(item => item.InstanceCount);
            return maxInstanceCount;
        }

        /// <summary>
        /// Returns the policy document name.
        /// </summary>
        /// <returns>File name</returns>
        private string GetPolicyDocumentName()
        {
            ////KLPolicy_<Policy Number>_<CreationDate>.pdf
            ////Example: MKLPolicy_EZXS1234567_20151103.pdf
            var created = this.Policy.Created.ToString("yyyyMMdd");
            return $"MKLPolicy_{this.Policy.PolicyNumber}_{created}";
        }

        /// <summary>
        /// Returns the name of document
        /// </summary>
        /// <param name="normalizedFormNumber">Normalized form number</param>
        /// <param name="quantityOrder">The Quantity Order</param>
        /// <param name="documentCreationDate">Document creation date</param>
        /// <returns>Form name</returns>
        private string GetFormName(string normalizedFormNumber, int? quantityOrder, DateTime documentCreationDate)
        {
            var created = documentCreationDate.ToString("yyyyMMdd");
            return $"{this.Policy.PolicyNumber}_{normalizedFormNumber}_{quantityOrder.GetValueOrDefault()}_{created}";
        }

        /// <summary>
        /// Process the merged form and apply template mappings including default and custom
        /// </summary>
        /// <param name="mergedForm">The merged form.</param>
        /// <param name="templateBinaryData">templateBinaryData</param>
        /// <param name="submissionNumber">submissionNumber</param>
        /// <param name="customMappings">customMappings</param>
        /// <param name="authenticatedUser">user</param>
        /// <param name="fileFormat">fileFormat</param>
        /// <param name="isWaterMark">Indicates if a watermarked document is being generated at Pre-Issue</param>
        /// <param name="instancesRequired">The number of instances required to be generated</param>
        /// <returns>UploadDocumentResponse</returns>
        private IPolicyDocumentManager GenerateMergedTemplateDocument(
            MergedForm mergedForm,
            Stream templateBinaryData,
            string submissionNumber,
            Dictionary<string, JToken> customMappings,
            IUser authenticatedUser,
            DocumentManager.FileFormat fileFormat,
            bool isWaterMark,
            int instancesRequired = 1)
        {
            string formNormalizedName = mergedForm.FormNormalizedNumber;
            IPolicyDocumentManager docManager = new PolicyDocumentManager(templateBinaryData, submissionNumber);
            docManager.FormNormalizedNumber = mergedForm.FormNormalizedNumber;
            docManager.QuantityOrder = mergedForm.QuantityOrder.GetValueOrDefault(0);

            this.ApplyDefaultMapping(ref docManager);
            this.ApplySignature(ref docManager, authenticatedUser);
            this.ApplyCustomMapping(ref docManager, customMappings);
            if (isWaterMark == true)
            {
                docManager.InsertWaterMark();
            }
            else
            {
                docManager.InsertWaterMark(string.Empty);
            }

            this.ReplaceSingleInstanceDocumentQuestionMergeFields(ref docManager, formNormalizedName);
            this.ReplaceDocumentQuestionMergeFields(ref docManager, formNormalizedName, customMappings, instancesRequired);

            return docManager;
        }

        /// <summary>
        /// Apply signatures
        /// </summary>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <param name="authenticatedUser">User</param>
        private void ApplySignature(ref IPolicyDocumentManager docManager, IUser authenticatedUser)
        {
            ////replace signature bookmark on document
            var images = docManager.Document.GetChildNodes(NodeType.Shape, true);
            foreach (Shape image in images)
            {
                switch (image.AlternativeText.ToLower())
                {
                    case "pressig":
                        var bytearray = this.GetSignatureBytes(authenticatedUser, CorporateSignatureType.President, this.Policy.EffectiveDate, this.Policy.Agency.StreetAddress.StateCode);
                        image.ImageData.SetImage(new MemoryStream(bytearray));
                        break;

                    case "secsig":
                        var secbytearray = this.GetSignatureBytes(authenticatedUser, CorporateSignatureType.Secretary, this.Policy.EffectiveDate);
                        image.ImageData.SetImage(new MemoryStream(secbytearray));
                        break;

                    case "authsig":
                        break;
                }
            }
        }

        /// <summary>
        /// Apply default mapping
        /// </summary>
        /// <param name="docManager">PolicyDocumentManager</param>
        private void ApplyDefaultMapping(ref IPolicyDocumentManager docManager)
        {
            var bookMarks = docManager.GetControlsBookMarks();
            foreach (Bookmark control in bookMarks)
            {
                var val = this.GetCarrierInformation(control.Name);

                if (string.IsNullOrWhiteSpace(val))
                {
                    val = this.templateProcessor.GetFieldValue(this.Policy, this.PolicyJson, control, ref docManager);
                }

                // If the value for bookmark was not resolved then set val = string.Empty
                // so that the default bookmark text does not appear on the generated document
                if (string.IsNullOrWhiteSpace(val))
                {
                    val = string.Empty;
                }

                docManager.ReplaceNodeValue(control, val);
            }
        }

        /// <summary>
        /// Apply custom mappings
        /// </summary>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <param name="customMappings">customMappings</param>
        /// <returns>bool</returns>
        private bool ApplyCustomMapping(ref IPolicyDocumentManager docManager, Dictionary<string, JToken> customMappings)
        {
            if (customMappings != null && customMappings.Count > 0)
            {
                foreach (KeyValuePair<string, JToken> entry in customMappings)
                {
                    // Loop through all the items in the custom mapping, break each item into key/value pairs and their associated
                    // parameters
                    this.templateProcessor.CreateCustomDictionary(entry.Value);

                    // Loop through key/value pairs and retrieve the actual document value from policy provided
                    foreach (KeyValuePair<string, MappingDetail> customFieldMapping in this.templateProcessor.CustomPolicyDictionary)
                    {
                        string val = this.templateProcessor.GetFieldValueFromCustom(customFieldMapping.Key, this.Policy, this.PolicyJson, ref docManager);

                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            Bookmark bookMark = docManager.GetBookMarkByName(customFieldMapping.Key);
                            if (bookMark != null)
                            {
                                docManager.ReplaceNodeValue(bookMark, val);
                            }
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Gets the carrier information
        /// </summary>
        /// <param name="fieldName">field name</param>
        /// <returns>field value</returns>4
        private string GetCarrierInformation(string fieldName)
        {
            if (this.policyCarrier == null)
            {
                return string.Empty;
            }

            switch (fieldName.ToLowerInvariant())
            {
                case "companyaddress1":
                    return this.policyCarrier.AddressLine1;

                case "companycity":
                    return this.policyCarrier.City;

                case "companystate":
                    return this.policyCarrier.State;

                case "companyzip":
                    return this.policyCarrier.ZipCode;

                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Replaces the single instance document question merge fields.
        /// </summary>
        /// <param name="docManager">The document manager.</param>
        /// <param name="formNormalizedNumber">The form normalized number.</param>
        private void ReplaceSingleInstanceDocumentQuestionMergeFields(ref IPolicyDocumentManager docManager, string formNormalizedNumber)
        {
            var quantityOrder = docManager.QuantityOrder;
            var formFromRollUp = this.Policy
                .RollUp
                .Documents
                .FirstOrDefault(x => x.NormalizedNumber.EqualsIgnoreCase(formNormalizedNumber) &&
                                    x.QuantityOrder.GetValueOrDefault(0) == quantityOrder);

            // Get all of the questions or answers that have a merge field name
            var singleInstanceQuestions = formFromRollUp?.Questions?.Where(x => (!string.IsNullOrWhiteSpace(x.MergeFieldName) ||
                                                                                (x.Answers != null &&
                                                                                x.Answers.Any(a => !string.IsNullOrWhiteSpace(a.MergeFieldName)))) &&
                                                                                string.IsNullOrWhiteSpace(x.ControllingQuestionCode) &&
                                                                                x.MultipleRowGroupingNumber == 0)
                .ToList();

            if (singleInstanceQuestions == null ||
                !singleInstanceQuestions.Any())
            {
                return;
            }

            // Set the list of merge fields that have already been processed.
            var handledMergeFields = new HashSet<string>();
            var editionDate = docManager.EditionDate.ToString("mmyy");
            var methodName = editionDate == "0001" ? $"Format_{formNormalizedNumber}" : $"Format_{formNormalizedNumber}_{editionDate}";

            MethodInfo formatMethod = null;
            formatMethods.TryGetValue(methodName, out formatMethod);
            if (formatMethods.ContainsKey(methodName))
            {
                formatMethods.TryGetValue(methodName, out formatMethod);
            }
            else
            {
                formatMethod = this.GetFormatMethod(methodName);
                formatMethods.Add(methodName, formatMethod);
            }

            // Loop through the questions that have merge field names
            foreach (var question in singleInstanceQuestions)
            {
                var workingQuestion = GetWorkingQuestion(question, formFromRollUp);
                if (workingQuestion == null)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(workingQuestion.MergeFieldName) &&
                    workingQuestion.Answers != null)
                {
                    // Loop through the answers and handle the merge fields
                    foreach (var answer in workingQuestion.Answers.Where(a => !string.IsNullOrWhiteSpace(a.MergeFieldName)).ToList())
                    {
                        HandleMergeField(answer.MergeFieldName, docManager, workingQuestion, handledMergeFields, this.Policy, formatMethod);
                    }

                    continue;
                }

                HandleMergeField(workingQuestion.MergeFieldName, docManager, workingQuestion, handledMergeFields, this.Policy, formatMethod);
            }
        }

        /// <summary>
        /// Replaces Simple blank merge field with values from Policy.RollUp.Documents.Questions.
        /// NOTE: This method will NOT work if there are multiple tables that have different MaximumMultipleRowCount or have different
        ///         number of answers such that a different number of forms would be required.
        /// </summary>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <param name="formNormalizedNumber">string</param>
        /// <param name="customMappings">customMappings</param>
        /// <param name="instancesRequired">The number of instances of this form that we need to replace merge fields for</param>
        private void ReplaceDocumentQuestionMergeFields(ref IPolicyDocumentManager docManager, string formNormalizedNumber, Dictionary<string, JToken> customMappings, int instancesRequired = 1)
        {
            var quantityOrder = docManager.QuantityOrder;
            var formFromRollUp = this.Policy
                                    .RollUp
                                    .Documents
                                    .FirstOrDefault(x => x.NormalizedNumber.EqualsIgnoreCase(formNormalizedNumber) &&
                                                            x.QuantityOrder.GetValueOrDefault(0) == quantityOrder);

            // Get all of the questions or answers that have a merge field name
            var bookmarkQuestions = formFromRollUp?.Questions?.Where(x => (!string.IsNullOrWhiteSpace(x.MergeFieldName) ||
                                                                            x.Answers.Any(a => !string.IsNullOrWhiteSpace(a.MergeFieldName))) &&
                                                                            string.IsNullOrWhiteSpace(x.ControllingQuestionCode) &&
                                                                            x.MultipleRowGroupingNumber > 0 &&
                                                                            x.MaximumMultipleRowCount.GetValueOrDefault(0) > 0)
                    .OrderBy(q => q.MultipleRowGroupingNumber)
                    .ToList();
            if (bookmarkQuestions == null ||
                !bookmarkQuestions.Any())
            {
                return;
            }

            // Set the list of merge fields that have already been processed.
            var handledRowGroupingNumber = new Dictionary<int, int>();
            var instanceItemNumber = new Dictionary<int, int>();
            var handledMergeFields = new HashSet<Tuple<int, string>>();
            var docInstances = new IPolicyDocumentManager[instancesRequired];
            var preFilledOutDocument = docManager.Document.Clone();
            docInstances[0] = docManager;

            var editionDate = docManager.EditionDate.ToString("mmyy");
            var methodName = editionDate == "0001" ? $"Format_{formNormalizedNumber}" : $"Format_{formNormalizedNumber}_{editionDate}";

            MethodInfo formatMethod = null;
            formatMethods.TryGetValue(methodName, out formatMethod);
            if (formatMethods.ContainsKey(methodName))
            {
                formatMethods.TryGetValue(methodName, out formatMethod);
            }
            else
            {
                formatMethod = this.GetFormatMethod(methodName);
                formatMethods.Add(methodName, formatMethod);
            }

            // Loop through the questions that have merge field names
            foreach (var question in bookmarkQuestions)
            {
                // Determine what instance of the form the row grouping number belongs on.
                var rowGroupingNumber = question.MultipleRowGroupingNumber;
                var documentInstance = 0;
                var itemInstance = 0;
                if (!handledRowGroupingNumber.ContainsKey(rowGroupingNumber))
                {
                    handledRowGroupingNumber.Add(rowGroupingNumber, documentInstance);
                    documentInstance = handledRowGroupingNumber.Count / question.MaximumMultipleRowCount.GetValueOrDefault(0);
                    if (handledRowGroupingNumber.Count % question.MaximumMultipleRowCount.GetValueOrDefault(0) != 0)
                    {
                        documentInstance++;
                    }

                    handledRowGroupingNumber[rowGroupingNumber] = documentInstance;

                    itemInstance = handledRowGroupingNumber.Count(i => i.Value == documentInstance);
                    instanceItemNumber.Add(rowGroupingNumber, itemInstance);
                }

                documentInstance = handledRowGroupingNumber[rowGroupingNumber] - 1;
                itemInstance = instanceItemNumber[rowGroupingNumber] - 1;

                if (docInstances[documentInstance] == null)
                {
                    var newDoc = preFilledOutDocument.Clone();
                    var manager = new PolicyDocumentManager(newDoc);
                    docInstances[documentInstance] = manager;
                }

                // Get the working question
                var workingQuestion = GetWorkingQuestion(question, formFromRollUp);
                if (workingQuestion == null)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(workingQuestion.MergeFieldName) &&
                    workingQuestion.Answers != null)
                {
                    foreach (var answer in workingQuestion.Answers.Where(a => !string.IsNullOrWhiteSpace(a.MergeFieldName)).ToList())
                    {
                        HandleMultipleMergeField(documentInstance, itemInstance, answer.MergeFieldName, docInstances[documentInstance], workingQuestion, handledMergeFields, this.Policy, formatMethod);
                    }

                    continue;
                }

                HandleMultipleMergeField(documentInstance, itemInstance, question.MergeFieldName, docInstances[documentInstance], workingQuestion, handledMergeFields, this.Policy, formatMethod);
            }

            if (instancesRequired > 1)
            {
                docInstances[0].UpdateTotalPageNumberFooterField();
                for (var instance = 1; instance < instancesRequired; instance++)
                {
                    if (docInstances[instance] == null)
                    {
                        break;
                    }

                    docInstances[instance].UpdateTotalPageNumberFooterField();
                    docManager.AppendDocument(docInstances[instance].Document, true, true);
                }

                docManager.Document.UpdatePageLayout();
            }
        }

        /// <summary>
        /// Gets the format method.
        /// </summary>
        /// <param name="methodName">Name of the method.</param>
        /// <returns>MethodInfo.</returns>
        private MethodInfo GetFormatMethod(string methodName)
        {
            object o = new CustomFunctions();
            var method = o.GetType()
                .GetMethod(
                    methodName,
                    BindingFlags.Instance | BindingFlags.Public,
                    null,
                    new Type[] { typeof(string), typeof(IPolicyDocumentManager), typeof(Question), typeof(Bookmark), typeof(string), typeof(Policy) },
                    null);
            if (method == null || method.ReturnType != typeof(string))
            {
                return null;
            }

            return method;
        }

        /// <summary>
        /// Upload the document to storage provider location
        /// </summary>
        /// <param name="docManager">Document manager</param>
        /// <param name="formName">Form normalized number</param>
        /// <param name="quantityOrder">The quantity order</param>
        /// <param name="fileFormat">File format</param>
        /// <returns>Document upload response</returns>
        private UploadDocumentResponse UploadDocumentToStore(ref IPolicyDocumentManager docManager, string formName, int? quantityOrder, DocumentManager.FileFormat fileFormat)
        {
            UploadDocumentResponse docResponse = new UploadDocumentResponse();
            if (docManager == null)
            {
                return docResponse;
            }

            using (var docStream = new MemoryStream())
            {
                docManager.Save(docStream, fileFormat);
                docStream.Position = 0;
                var contents = docStream.ToArray();
                docResponse = this.storageManager.UploadIssuanceDocument(new UploadDocumentRequest()
                {
                    DestinationFileName = this.GetFormName(formName, quantityOrder, DateTime.Now) + "." + fileFormat,
                    DocumentBinary = contents,
                    PolicyId = this.Policy.Id.ToString()
                });
                docStream.Flush();
            }

            return docResponse;
        }

        /// <summary>
        /// Upload the document to storage provider location
        /// </summary>
        /// <param name="docStream">Stream to be saved</param>
        /// <param name="fileName">File name</param>
        /// <param name="format">FileFormat</param>
        /// <returns>UploadDocumentResponse</returns>
        private UploadDocumentResponse UploadDocumentToStore(MemoryStream docStream, string fileName, DocumentManager.FileFormat format)
        {
            UploadDocumentResponse docResponse = new UploadDocumentResponse();
            if (docStream != null)
            {
                docStream.Position = 0;
                byte[] contents = docStream.ToArray();
                docResponse = this.storageManager.UploadIssuanceDocument(new UploadDocumentRequest()
                {
                    DestinationFileName = fileName + "." + format,
                    DocumentBinary = contents,
                    PolicyId = this.Policy.Id.ToString()
                });
            }

            return docResponse;
        }

        /// <summary>
        /// Get default template mapping from services
        /// </summary>
        /// <param name="authenticatedUser">authenticatedUser</param>
        /// <param name="effectiveDate">DateTime</param>
        /// <returns>JToken</returns>
        private JToken GetDefaultMappingTemplate(IUser authenticatedUser, DateTime? effectiveDate)
        {
            JToken templates = this.restHelper.GetServiceResult("DefaultMappingTemplates?effectiveDate=" + effectiveDate.GetValueOrDefault(DateTime.UtcNow).ToShortDateString(), RestSharp.Method.GET, authenticatedUser, null, WebTeamServiceEndpoint.BaseServices);
            return templates;
        }

        /// <summary>
        /// Get custom mapping templates associated to a form
        /// </summary>
        /// <param name="authenticatedUser">authenticatedUser</param>
        /// <param name="templateMappingId">string</param>
        /// <param name="effectiveDate">DateTime</param>
        /// <returns>Dictionary of JToken</returns>
        private Dictionary<string, JToken> GetCustomMappingTemplate(IUser authenticatedUser, string templateMappingId, DateTime? effectiveDate)
        {
            if (!string.IsNullOrWhiteSpace(templateMappingId))
            {
                Guid id;
                Guid.TryParse(templateMappingId, out id);
                var mappingTemplates = this.restHelper.GetServiceResult("CustomMappingTemplates?templateId=" + id + "&effectiveDate=" + effectiveDate.Value.ToShortDateString(), RestSharp.Method.GET, authenticatedUser, null, WebTeamServiceEndpoint.BaseServices);
                var customMappingTemplates = JsonConvert.DeserializeObject<Dictionary<string, JToken>>(mappingTemplates.ToString());
                return customMappingTemplates;
            }

            return null;
        }

        /// <summary>
        /// GetDocumentMetadata - Gets the document's metadata info based on form number and effective date
        /// </summary>
        /// <param name="authenticatedUser">User</param>
        /// <param name="formNumber">string</param>
        /// <param name="effectiveDate">DateTime</param>
        /// <returns>JToken</returns>
        private JToken GetDocumentMetadata(IUser authenticatedUser, string formNumber, DateTime? effectiveDate)
        {
            DocumentSearchRepresentation searchRepresentation = new DocumentSearchRepresentation() { EditionDate = null, EffectiveOnlyDate = effectiveDate, FormNumber = formNumber };

            var response = this.restHelper.GetServiceResult("DocumentMetaData", RestSharp.Method.POST, authenticatedUser, searchRepresentation, WebTeamServiceEndpoint.BaseServices);
            return response;
        }

        /// <summary>
        /// GetDocTemplateBinaryData - Gets the word template contents in binary format
        /// </summary>
        /// <param name="authenticatedUser">User</param>
        /// <param name="documentId">string</param>
        /// <param name="effectiveDate">DateTime</param>
        /// <param name="templateMappingId">out string</param>
        /// <returns> byte[]</returns>
        private byte[] GetDocTemplateBinaryData(IUser authenticatedUser, string documentId, DateTime? effectiveDate, out string templateMappingId)
        {
            ////Get document header
            var headerUrl = "Document?headerId=" + new Guid(documentId);
            if (effectiveDate.HasValue)
            {
                headerUrl += "&asOfDate=" + effectiveDate.Value;
            }

            var documentHeader = this.restHelper.GetServiceResult(headerUrl, RestSharp.Method.GET, authenticatedUser, null, WebTeamServiceEndpoint.BaseServices);

            ////Get document template mapping id required to map a template to a custom mapping
            var docTemplateMappingId = documentHeader.SelectToken("DocumentTemplateMappingId", false);
            templateMappingId = docTemplateMappingId?.ToString() ?? string.Empty;

            var detailId = documentHeader.SelectToken("DetailId", true);
            var url = "Document/" + detailId.ToString() + "?DocumentType=WordTemplate";

            ////Get document details
            var documentBinary = this.restHelper.GetServiceResult(url, RestSharp.Method.GET, authenticatedUser, null, WebTeamServiceEndpoint.BaseServices);
            var wordTemplateData = documentBinary.SelectToken("WordTemplateBinaryData", true).ToString();

            return Convert.FromBase64String(wordTemplateData);
        }

        /// <summary>
        /// Returns the list of documents (grouped by lob) to be consolidated into a policy document.
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <returns>Ordered list of documents grouped by LOB</returns>
        private Dictionary<dynamic, Dictionary<string, DecisionModel.Models.Policy.Document>> GetOrderedPolicyForms(Policy policy)
        {
            var formsByLob = new Dictionary<dynamic, Dictionary<string, DecisionModel.Models.Policy.Document>>();
            var commonDocuments = policy.RollUp.Documents.Where(d => d.IsCommonForm() && d.ReturnIsSelected() && d.Section != DocumentSection.Application).ToList();
            var commonFormsDictionary = new Dictionary<string, DecisionModel.Models.Policy.Document>();
            commonDocuments.ForEach(d => commonFormsDictionary.Add(this.GetDocumentKey(d), d));
            dynamic lobKey = new System.Dynamic.ExpandoObject();
            lobKey.Title = "COMMON";
            lobKey.Order = 1;
            formsByLob.Add(lobKey, commonFormsDictionary);

            // Group and dedupe the forms within each line of business.
            // This is required to pull GL, OCP, liquor, and special events into one list.
            foreach (var lob in policy.LobOrder)
            {
                lobKey = new System.Dynamic.ExpandoObject();
                switch (lob.Code)
                {
                    case LineOfBusiness.CF:
                        lobKey.Title = "PROPERTY";
                        lobKey.Order = 3;
                        formsByLob.Add(lobKey, new Dictionary<string, DecisionModel.Models.Policy.Document>());
                        break;
                    case LineOfBusiness.IM:
                        lobKey.Title = "INLAND MARINE";
                        lobKey.Order = 4;
                        formsByLob.Add(lobKey, new Dictionary<string, DecisionModel.Models.Policy.Document>());
                        break;
                    case LineOfBusiness.XS:
                        lobKey.Title = "EXCESS LIABILITY";
                        lobKey.Order = 5;
                        formsByLob.Add(lobKey, new Dictionary<string, DecisionModel.Models.Policy.Document>());
                        break;
                    default:
                        lobKey.Title = "GENERAL LIABILITY";
                        lobKey.Order = 2;
                        if (!formsByLob.ContainsKey(lobKey))
                        {
                            formsByLob.Add(lobKey, new Dictionary<string, DecisionModel.Models.Policy.Document>());
                        }

                        break;
                }

                foreach (var document in policy.GetLine(lob.Code).RollUp.Documents.Where(d => !d.IsCommonForm() && d.ReturnIsSelected() && d.Section != DocumentSection.Application))
                {
                    var key = this.GetDocumentKey(document);
                    if (!formsByLob[lobKey].ContainsKey(key))
                    {
                        formsByLob[lobKey].Add(key, document);
                    }
                }
            }

            return formsByLob;
        }

        /// <summary>
        /// Gets the document key.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns>The key</returns>
        private string GetDocumentKey(DecisionModel.Models.Policy.Document document)
        {
            return document.NormalizedNumber + document.EditionDate.ToString("MMyy") + document.QuantityOrder;
        }
    }
}