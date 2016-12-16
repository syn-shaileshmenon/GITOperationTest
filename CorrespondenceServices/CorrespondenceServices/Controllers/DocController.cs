// <copyright file="DocController.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Web;
    using System.Web.Http;
    using System.Web.Http.Results;
    using Aspose.Words;
    using Aspose.Words.Markup;
    using Aspose.Words.Replacing;
    using CorrespondenceServices.Attributes;
    using CorrespondenceServices.Classes.Parameters;
    using DecisionModel.Classes;
    using DecisionModel.Interfaces;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.DocumentGenerator.Helpers;
    using Mkl.WebTeam.RestHelper.Enumerations;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.SubmissionShared.Enumerations;
    using Mkl.WebTeam.WebCore2;
    using RestSharp;

    /// <summary>
    /// To manage generation and downloading of documents
    /// </summary>
    [Security]
    public class DocController : BaseApiController
    {
        /// <summary>
        /// The echo service URL
        /// </summary>
        private readonly string echoServices = ConfigurationManager.AppSettings["EchoServicesUrl"].ToLower();

        /// <summary>
        /// The echo service URL
        /// </summary>
        private readonly string echoServicesEnabled = ConfigurationManager.AppSettings["EchoService:Enabled"].ToLower();

        /// <summary>
        /// The document path for the template
        /// </summary>
        private string dataPath = AppDomain.CurrentDomain.GetData("DataDirectory").ToString();

        /// <summary>
        /// The docManager
        /// </summary>
        private PolicyDocumentManager docManager;

        /// <summary>
        /// The session identifier
        /// </summary>
        private string folderId = string.Empty;

        /// <summary>
        /// The session key
        /// </summary>
        private string sessionKey = "SaveParams";

        /// <summary>
        /// The temporary path
        /// </summary>
        private string tempPath = ConfigurationManager.AppSettings["tempPath"].ToString();

        /// <summary>
        /// The temporary path
        /// </summary>
        private string saveToFile = ConfigurationManager.AppSettings["saveToFile"].ToString();

        /// <summary>
        /// The temporary path
        /// </summary>
        private string threaded = ConfigurationManager.AppSettings["threaded"].ToString();

        /// <summary>
        /// The policy
        /// </summary>
        private Policy policy;

        /// <summary>
        /// Issuance controller
        /// </summary>
        ////TODO: VB - This is for testing and needs to be be removed once testing is done.
        private IssuanceController issuanceController;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocController" /> class.
        /// </summary>
        /// <param name="bootStrapper">IBootstrapper</param>
        /// <param name="jsonManager">IJSONManager</param>
        /// <param name="restHelper">The rest helper.</param>
        /// <param name="storageManager">The storage manager.</param>
        /// <param name="templateProcessor">The template processor.</param>
        public DocController(IBootstrapper bootStrapper, IJsonManager jsonManager, IRestHelper restHelper, IStorageManager storageManager, ITemplateProcessor templateProcessor)
            : base(jsonManager, bootStrapper, restHelper, storageManager, templateProcessor)
        {
            ////TODO: VB - This is the testing code and needs to be removed once the testing is complete
            this.issuanceController = new IssuanceController(bootStrapper, jsonManager, restHelper, storageManager, templateProcessor);
        }

        /// <summary>
        /// Pings this instance.
        /// </summary>
        /// <returns>IHttpActionResult.</returns>
        [HttpGet]
        [HttpPost]
        [HttpPut]
        [HttpDelete]
        [HttpOptions]
        [Route("api/KeepAlive")]
        public IHttpActionResult KeepAlive()
        {
            var message = string.Empty;

            try
            {
                if (this.echoServicesEnabled == "true")
                {
                    var response = this.ExecuteRequest("Echo", Method.GET, this.echoServices);

                    // call correspondence services (declination)
                    if (!response.Contains("OK"))
                    {
                        throw new HttpException("Service Unavailable", 500);
                    }
                }
            }
            catch (Exception ex)
            {
                message += "Echo service failed" + Environment.NewLine;
                message += ex.ToString();
            }

            if (!string.IsNullOrWhiteSpace(message))
            {
                return new InternalServerErrorResult(this);
            }

            return this.Ok("OK");
        }

        /// <summary>
        /// This GET method will support the request to download a document based on the user's session. This will only
        /// be active if the user has an active session with information for a saved document
        /// </summary>
        /// <param name="docFormat">The document format.</param>
        /// <returns>
        /// HttpResponseMessage
        /// </returns>
        [Route("api/GetDocument/{format}")]
        public HttpResponseMessage GetDocument(DocumentManager.FileFormat docFormat)
        {
            HttpResponseMessage response = this.Request.CreateResponse();
            var session = HttpContext.Current.Session;
            if (session != null && session[this.sessionKey] != null)
            {
                DownloadParameters dloadParams = session[this.sessionKey] as DownloadParameters;
                if (dloadParams != null)
                {
                    if (File.Exists(dloadParams.TempFileName))
                    {
                        // this calls the class that handles the Aspose library
                        PolicyDocumentManager dm = new PolicyDocumentManager(dloadParams.TempFileName, session[this.sessionKey].ToString());

                        System.IO.MemoryStream docStream = new System.IO.MemoryStream();
                        dm.Save(docStream, docFormat);
                        docStream.Position = 0;
                        response.Content = new StreamContent(docStream);
                        response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(MimeMapping.GetMimeMapping(dm.SaveNameWithExt));
                        response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
                        response.Content.Headers.ContentDisposition.FileName = dm.SaveNameWithExt;
                    }
                }
            }

            return response;
        }

        /// <summary>
        /// Generates and Saves the bind letter to the database.
        /// </summary>
        /// <param name="passedPolicy">The passed policy.</param>
        [HttpPost]
        [Route("api/SaveBind")]
        public void SaveBindLetter(Policy passedPolicy)
        {
            this.policy = passedPolicy;
            bool success = false;
            try
            {
                success = this.CreateDocument(this.GetQuoteBindTemplates(true), true);
            }
            catch (Exception e)
            {
                string error = " Bind Error. Transaction: " + this.policy.SubmissionNumber + ". Error message: " + e.Message;
                LogEvent.Log.Error(error);
                LogEvent.TraceLog.Error(error, e);
                throw new ArgumentException(error, e);
            }

            string policyNum = (this.policy.PolicyNumber == null) ? "NoPolicyNum" : this.policy.PolicyNumber.ToString();
            string fileName = "MklBinder_" + policyNum + DateTime.Now.ToString("MMddyyyyhhmmss");

            // uncomment these lines if you want to save these to the temp/[policyId] directory
            if (this.saveToFile == "true")
            {
                success = this.SaveToFile(fileName, DocumentManager.FileFormat.PDF);
                success = this.SaveToFile(fileName, DocumentManager.FileFormat.DOCX);
            }

            MemoryStream pngStream = new MemoryStream();
            this.docManager.Save(pngStream, DocumentManager.FileFormat.TIFF);

            Thread pngThread = new Thread(() => this.SavePNG("BindImage/" + this.policy.Id, fileName, pngStream));
            if (this.threaded == "true")
            {
                pngThread.Start();
            }

            MemoryStream pdfStream = new MemoryStream();
            this.docManager.Save(pdfStream, DocumentManager.FileFormat.PDF);

            MemoryStream docStream = new MemoryStream();
            this.docManager.Save(docStream, DocumentManager.FileFormat.DOCX);

            if (this.threaded == "true")
            {
                Thread pdfThread = new Thread(() => this.SavePDF("BindLetter/" + this.policy.Id.ToString(), pdfStream));
                pdfThread.Start();

                Thread docThread = new Thread(() => this.SaveWord("BindDoc/" + this.policy.Id.ToString(), docStream));
                docThread.Start();

                if (pngThread.IsAlive)
                {
                    pngThread.Join();
                }

                if (pdfThread.IsAlive)
                {
                    pdfThread.Join();
                }

                if (docThread.IsAlive)
                {
                    docThread.Join();
                }
            }
            else
            {
                this.SavePDF("BindLetter/" + this.policy.Id.ToString(), pdfStream);
                this.SaveWord("BindDoc/" + this.policy.Id.ToString(), docStream);
                this.SavePNG("BindImage/" + this.policy.Id, fileName, pngStream);
            }

            pngStream.Close();
            docStream.Close();
            pdfStream.Close();
        }

        /// <summary>
        /// Saves the quote letter.
        /// </summary>
        /// <param name="passedPolicy">The passed policy.</param>
        /// <returns>boolean</returns>
        [HttpPost]
        [Route("api/SaveDec")]
        public bool SaveDeclinationLetter(Policy passedPolicy)
        {
            this.policy = passedPolicy;
            bool success = false;
            List<string> docs = new List<string>();
            docs.Add(this.dataPath + "/decLetterCover.docx");
            try
            {
                success = this.CreateDocument(docs.ToArray(), false);
            }
            catch (Exception e)
            {
                string error = " Declination Letter Error. Transaction: " + this.policy.SubmissionNumber + ". Error message: " + e.Message;
                LogEvent.Log.Error(error);
                LogEvent.TraceLog.Error(error, e);
                throw new ArgumentException(error, e);
            }

            string submissionNum = (this.policy.SubmissionNumber == null) ? "NoSubmissionNum" : this.policy.SubmissionNumber.ToString();
            string fileName = "MklDeclination_" + submissionNum + DateTime.Now.ToString("MMddyyyyhhmmss");

            MemoryStream pdfStream = new MemoryStream();
            this.docManager.Save(pdfStream, DocumentManager.FileFormat.PDF);
            this.SavePDF("DeclinationLetter/" + this.policy.Id.ToString(), pdfStream);
            pdfStream.Close();

            return success;
        }

        /// <summary>
        /// Generates and Saves the quote letter to the database.
        /// </summary>
        /// <param name="passedPolicy">The passed policy.</param>
        [HttpPost]
        [Route("api/SaveDocument")]
        public void SaveQuoteLetter(Policy passedPolicy)
        {
            this.policy = passedPolicy;
            bool success = false;
            try
            {
                success = this.CreateDocument(this.GetQuoteBindTemplates(), false);
            }
            catch (Exception e)
            {
                string error = " Quote Error. Transaction: " + this.policy.SubmissionNumber + ". Error message: " + e.Message;
                LogEvent.Log.Error(error);
                LogEvent.TraceLog.Error(error, e);
                throw new ArgumentException(error, e);
            }

            string submissionNum = (this.policy.SubmissionNumber == null) ? "NoSubmissionNum" : this.policy.SubmissionNumber.ToString();
            string fileName = "MklQuote_" + submissionNum + DateTime.UtcNow.ToString("MMddyyyyhhmmss");

            // uncomment these lines if you want to save these to the temp/[folderId] directory
            if (this.saveToFile == "true")
            {
                success = this.SaveToFile(fileName, DocumentManager.FileFormat.PDF);
                success = this.SaveToFile(fileName, DocumentManager.FileFormat.DOCX);
            }

            MemoryStream pngStream = new MemoryStream();
            this.docManager.Save(pngStream, DocumentManager.FileFormat.TIFF);

            Thread pngThread = new Thread(() => this.SavePNG("QuoteImage/" + this.policy.Id, fileName, pngStream));
            if (this.threaded == "true")
            {
                pngThread.Start();
            }

            MemoryStream pdfStream = new MemoryStream();
            this.docManager.Save(pdfStream, DocumentManager.FileFormat.PDF);

            MemoryStream docStream = new MemoryStream();
            this.docManager.Save(docStream, DocumentManager.FileFormat.DOCX);

            if (this.threaded == "true")
            {
                Thread pdfThread = new Thread(() => this.SavePDF("QuoteLetter/" + this.policy.Id.ToString(), pdfStream));
                pdfThread.Start();

                Thread docThread = new Thread(() => this.SaveWord("QuoteDoc/" + this.policy.Id.ToString(), docStream));
                docThread.Start();

                if (pngThread.IsAlive)
                {
                    pngThread.Join();
                }

                if (pdfThread.IsAlive)
                {
                    pdfThread.Join();
                }

                if (docThread.IsAlive)
                {
                    docThread.Join();
                }
            }
            else
            {
                this.SavePDF("QuoteLetter/" + this.policy.Id.ToString(), pdfStream);
                this.SaveWord("QuoteDoc/" + this.policy.Id.ToString(), docStream);
                this.SavePNG("QuoteImage/" + this.policy.Id, fileName, pngStream);
            }

            pngStream.Close();
            pdfStream.Close();
            docStream.Close();
        }

        /// <summary>
        /// Saves the PDF.
        /// </summary>
        /// <param name="id">file identifier.</param>
        /// <param name="docStream">The document stream.</param>
        private void SavePDF(string id, MemoryStream docStream)
        {
            var docFile = docStream.ToArray();
            this.RestHelper.GetServiceResult("Policy/" + id, RestSharp.Method.PUT, this.AuthenticatedUser, docFile, WebTeamServiceEndpoint.DecisionServices);
        }

        /// <summary>
        /// Saves the PNG.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="docStream">The document stream.</param>
        private void SavePNG(string id, string fileName, MemoryStream docStream)
        {
            var docFile = this.docManager.SaveToPng(docStream, Path.Combine(this.folderId, fileName), this.saveToFile == "true");
            this.RestHelper.GetServiceResult("Policy/" + id, RestSharp.Method.PUT, this.AuthenticatedUser, docFile, WebTeamServiceEndpoint.DecisionServices);
        }

        /// <summary>
        /// Saves the word.
        /// </summary>
        /// <param name="id">The identifier.</param>
        /// <param name="docStream">The document stream.</param>
        private void SaveWord(string id, MemoryStream docStream)
        {
            var docFile = docStream.ToArray();
            this.RestHelper.GetServiceResult("Policy/" + id, RestSharp.Method.PUT, this.AuthenticatedUser, docFile, WebTeamServiceEndpoint.DecisionServices);
        }

        /// <summary>
        /// Standard POST method to save a WORD doc with values from JSON object
        /// </summary>
        /// <param name="docs">The docs.</param>
        /// <param name="isBindLetter">if set to <c>true</c> [is bind letter].</param>
        /// <returns>
        /// boolean
        /// </returns>
        private bool CreateDocument(string[] docs, bool isBindLetter = false)
        {
            bool success = false;
            Dictionary<string, string> policyDict = new Dictionary<string, string>();

            this.folderId = this.policy.Id.ToString();

            // this calls the class that handles the Aspose library
            this.docManager = new PolicyDocumentManager(docs, this.policy.SubmissionNumber);

            foreach (var control in this.docManager.GetControls())
            {
                string switchVal = string.Empty;
                string controlVal;
                string controlId = control.Tag.ToString();
                int pointIndex = controlId.IndexOf(".");
                switchVal = (pointIndex == -1) ? controlId : controlId.Substring(0, pointIndex);

                if (switchVal == "_IfQuote")
                {
                    if (!isBindLetter)
                    {
                        if (control.ParentNode == null)
                        {
                            continue;
                        }

                        controlId = controlId.Substring(pointIndex + 1);
                        pointIndex = controlId.IndexOf(".");
                        switchVal = (pointIndex == -1) ? controlId : controlId.Substring(0, pointIndex);
                    }
                    else
                    {
                        control.Remove();
                    }
                }

                if (switchVal == "_KeepFirst" || switchVal == "_KeepLast")
                {
                    if (switchVal == "_KeepFirst")
                    {
                        this.KeepFirst(control.Tag.ToString());
                    }
                    else if (switchVal == "_KeepLast")
                    {
                        this.KeepLast(control.Tag.ToString());
                    }

                    if (control.ParentNode == null)
                    {
                        continue;
                    }

                    controlId = controlId.Substring(pointIndex + 1);
                    pointIndex = controlId.IndexOf(".");
                    switchVal = (pointIndex == -1) ? controlId : controlId.Substring(0, pointIndex);
                }

                if (switchVal == "_EffBefore" || switchVal == "_EffAfter")
                {
                    string i = controlId.Substring(pointIndex + 1);
                    pointIndex = i.IndexOf(".");
                    string dateValue = (pointIndex == -1) ? i : i.Substring(0, pointIndex);
                    bool continueOn = false;
                    if (switchVal == "_EffBefore")
                    {
                        continueOn = this.EffectiveBefore(control.Tag.ToString(), dateValue);
                    }
                    else if (switchVal == "_EffAfter")
                    {
                        continueOn = this.EffectiveAfter(control.Tag.ToString(), dateValue);
                    }

                    controlId = (pointIndex == -1) ? i : i.Substring(pointIndex + 1);
                    pointIndex = controlId.IndexOf(".");
                    switchVal = (pointIndex == -1 || !continueOn) ? switchVal : controlId.Substring(0, pointIndex);
                }

                switch (switchVal)
                {
                    case "_KeepFirst":
                    case "_KeepLast":
                    case "_EffBefore":
                    case "_EffAfter":
                    case "_IfQuote":
                        break;

                    case "_Clauses":
                        controlVal = controlId.Substring(pointIndex + 1);
                        ImRiskUnit imClauseRisk = this.GetImRiskByType(controlVal);
                        IEnumerable<Clause> clauses = imClauseRisk.Clauses.Where(x => x.TypeCode != "OPTCOV" && imClauseRisk.RollUp.SelectedClauses.Contains(x.Code)).OrderBy(x => x.Name);
                        if (clauses != null && clauses.Count() > 0)
                        {
                            IEnumerable<DecisionModel.Models.Policy.Document> documents = this.policy.ImLine.RollUp.Documents.FindAll(x => x.ReturnIsSelected()).Distinct(new DocumentComparer());
                            this.docManager.InsertClauses(clauses, control, documents, false);
                        }

                        control.Remove();
                        break;

                    case "_CoverageOptions":
                        controlVal = controlId.Substring(pointIndex + 1);
                        ImRiskUnit imCovRisk = this.GetImRiskByType(controlVal);
                        IEnumerable<Clause> covOptions = imCovRisk.Clauses.Where(x => x.TypeCode == "OPTCOV" && imCovRisk.RollUp.SelectedClauses.Contains(x.Code)).OrderBy(x => x.Name);
                        if (covOptions != null && covOptions.Count() > 0)
                        {
                            IEnumerable<DecisionModel.Models.Policy.Document> documents = this.policy.ImLine.RollUp.Documents.FindAll(x => x.ReturnIsSelected()).Distinct(new DocumentComparer());
                            IEnumerable<OptionalCoverage> imOptCov = imCovRisk.OptionalCoverages.Where(x => imCovRisk.RollUp.SelectedClauses.Contains(x.Code));
                            this.docManager.InsertCoverageOptions(covOptions, control, documents, imOptCov, false);
                        }

                        control.Remove();
                        break;

                    case "_Coverages":
                        controlVal = (pointIndex == -1) ? "RollUp.OptionalCoverages" : controlId.Substring(pointIndex + 1) + ".RollUp.OptionalCoverages";
                        IEnumerable<DecisionModel.Models.Policy.OptionalCoverage> covObj = (List<DecisionModel.Models.Policy.OptionalCoverage>)StringExtension.GetPropObject(this.policy, controlVal);
                        covObj = covObj.Where(x => x.IsSelected && x.Code != "TRIA").OrderBy(x => x.Order);
                        if (covObj != null && covObj.Count() > 0)
                        {
                            int occurLimit = this.policy.GlLine == null ? 0 : this.policy.GlLine.Limits.PerOccurrence.GetValueOrDefault();
                            int deductible = this.policy.GlLine == null ? 0 : this.policy.GlLine.Limits.Deductible.GetValueOrDefault();
                            this.docManager.InsertCoveragesTable(covObj, control, occurLimit, deductible, controlId.Substring(pointIndex + 1) != "SpecialEventLine");
                        }

                        control.Remove();
                        break;

                    case "_Documents":
                        controlVal = (pointIndex == -1) ? "RollUp.Documents" : controlId.Substring(pointIndex + 1) + ".RollUp.Documents";
                        List<DecisionModel.Models.Policy.Document> docObj = (List<DecisionModel.Models.Policy.Document>)StringExtension.GetPropObject(this.policy, controlVal);
                        docObj = docObj?.FindAll(x => x.ReturnIsSelected());

                        this.docManager.InsertLinkListTable(docObj, control, true, isBindLetter);
                        control.Remove();
                        break;

                    case "_Exposures":
                        if (pointIndex != -1)
                        {
                            controlVal = controlId.Substring(pointIndex + 1);
                            IEnumerable<IBaseGlRiskUnit> riskObj = this.ReturnRiskUnits(controlVal);
                            if (riskObj != null)
                            {
                                this.docManager.InsertExposuresTable(riskObj, control);
                            }

                            control.Remove();
                        }

                        break;

                    case "_If":
                        if (pointIndex != -1)
                        {
                            controlVal = controlId.Substring(pointIndex + 1);
                            bool result = true;
                            if (policyDict.ContainsKey(controlId))
                            {
                                result = policyDict[controlId] == true.ToString();
                            }
                            else
                            {
                                string[] andEquations = controlVal.Split('&');
                                foreach (string andEquation in andEquations)
                                {
                                    string[] forEquations = andEquation.Split('|');
                                    int start = 0;
                                    if (!result)
                                    {
                                        start = 1;
                                    }

                                    for (int i = start; i < forEquations.Length; i++)
                                    {
                                        string[] operands = this.ParseIfStatement(forEquations[i], control);
                                        result = this.Equation(StringExtension.GetPropValue<string>(this.policy, operands[0]), operands[1], operands[2]);
                                        if (result)
                                        {
                                            break;
                                        }
                                    }
                                }

                                policyDict.Add(controlId, result.ToString());
                            }

                            this.EvaluateEquation(result, control);
                        }

                        break;

                    case "_IfCoverage":
                        if (pointIndex != -1)
                        {
                            controlVal = controlId.Substring(pointIndex + 1);
                            string[] controlOptions = controlVal.Split('|');
                            bool removeTheControl = true;
                            foreach (string controlOption in controlOptions)
                            {
                                string[] objArray = controlOption.Split('.');
                                List<DecisionModel.Models.Policy.OptionalCoverage> ifCovObj = this.ReturnCoverages(controlVal);
                                if (ifCovObj == null)
                                {
                                    ifCovObj = null;
                                }
                                else
                                {
                                    if (this.HandleCoverageControls(objArray[1], control, ifCovObj))
                                    {
                                        removeTheControl = false;
                                        break;
                                    }
                                }
                            }

                            if (removeTheControl)
                            {
                                control.Remove();
                            }
                        }

                        break;

                    case "_IfImRiskItem":
                        string[] compares = this.ParseIfStatement(controlId, control);
                        string compareLeft = string.Empty;
                        if (policyDict.ContainsKey(compares[0]))
                        {
                            compareLeft = policyDict[compares[0]];
                        }
                        else
                        {
                            compareLeft = this.ImItemValue(compares[0]) ?? "0";
                            policyDict.Add(compares[0], compareLeft);
                        }

                        this.EvaluateEquation(compareLeft, compares[1], compares[2], control);
                        break;

                    case "_IfWindHailExcluded":
                        controlVal = "false";
                        if (policyDict.ContainsKey(controlId))
                        {
                            controlVal = policyDict[controlId];
                        }
                        else
                        {
                            ImRiskUnit windHailRisk = this.GetImRiskByType(controlId.Substring(pointIndex + 1));
                            Question windHailQuestion = windHailRisk.Questions.Find(x => x.Code == "IM Wind Hail Surcharge");
                            if (windHailQuestion != null && windHailQuestion.AnswerValue == "yes")
                            {
                                controlVal = "true";
                            }
                            else
                            {
                                DecisionModel.Models.Policy.Document windHailDoc = this.policy.RollUp.Documents.Find(x => x.NormalizedNumber == "IMWINDHAILEX");
                                if (windHailDoc != null && windHailDoc.ReturnIsSelected())
                                {
                                    controlVal = "true";
                                }
                            }

                            policyDict.Add(controlId, controlVal);
                        }

                        if (controlVal == "false")
                        {
                            control.Remove();
                        }

                        break;

                    case "_IfLOB":
                        if (pointIndex != -1)
                        {
                            controlVal = controlId.Substring(pointIndex + 1);
                            string[] objArray = controlVal.Split('.');
                            if (!this.policy.LobOrder.Any(x => x.Code.ToString() == objArray[0]))
                            {
                                control.Remove();
                            }
                        }

                        break;

                    case "_IfNotCoverage":
                        if (pointIndex != -1)
                        {
                            // if coverage exist and is selected remove control
                            controlVal = controlId.Substring(pointIndex + 1);
                            string[] objArray = controlVal.Split('.');
                            List<DecisionModel.Models.Policy.OptionalCoverage> notCovObjs = this.ReturnCoverages(objArray[0]);
                            if (notCovObjs != null && notCovObjs.Find(x => x.Code == objArray[1]) != null)
                            {
                                OptionalCoverage cov = notCovObjs.Find(x => x.Code == objArray[1]);
                                if (cov.IsSelected)
                                {
                                    control.Remove();
                                }
                            }
                        }

                        break;

                    case "_ImRisk":
                        this.ImRisk(controlId, control, policyDict);
                        break;

                    case "_ImItems":
                        controlVal = controlId.Substring(pointIndex + 1);
                        this.docManager.InsertImItemsTable(this.GetImRiskByType(controlVal), control, false);
                        control.Remove();
                        break;

                    case "_Layers":
                        if (this.policy.XsLine.IsDisplayOptionsOnQuote && !this.policy.Progress.Any(x => x.Step == TransactionStep.Bind))
                        {
                            List<XsLayer> layers = this.policy.XsLine.XsLimits;
                            this.docManager.InsertLayerTable(layers, control);
                        }

                        control.Remove();
                        break;

                    case "_List":
                        if (this.policy.RollUp.Subjectivities != null)
                        {
                            List<Subjectivity> list = this.policy.RollUp.Subjectivities.FindAll(x => x.IsSelected && !x.IsHidden);
                            if (list.Count > 0)
                            {
                                this.docManager.BuildUnorderedList(list.OrderBy(x => x.Order).ThenBy(x => x.Name), control, this.policy.Progress.Any(x => x.Step == TransactionStep.Bind));
                            }
                        }

                        control.Remove();
                        break;

                    case "_MailingAddress":
                        this.docManager.InsertMailingAddress(this.policy.SecondaryInsured?.StreetAddress, control);
                        control.Remove();
                        break;

                    case "_NonSupplementalDocuments":
                        if (this.policy.RollUp != null)
                        {
                            List<DecisionModel.Models.Policy.Document> nonSupplimentDocs = (List<DecisionModel.Models.Policy.Document>)this.policy.RollUp.Documents;
                            if (nonSupplimentDocs != null)
                            {
                                nonSupplimentDocs = nonSupplimentDocs.FindAll(x => x.Section != DocumentSection.Application && x.ReturnIsSelected());
                            }

                            this.docManager.InsertLinkListTable(nonSupplimentDocs, control, true, isBindLetter);
                        }

                        control.Remove();
                        break;

                    case "_PolicyDuration":
                        if (policyDict.ContainsKey(controlId))
                        {
                            controlVal = policyDict[controlId];
                        }
                        else
                        {
                            controlVal = ((DateTime)this.policy.ExpirationDate - (DateTime)this.policy.EffectiveDate).TotalDays.ToString("N0");
                            policyDict.Add(controlId, controlVal);
                        }

                        this.docManager.InsertValue(controlVal, control);
                        break;

                    case "_PolicyRenewalText":
                        controlVal = controlId.Substring(pointIndex + 1);
                        if (this.policy.ParentPolicy != null && this.policy.ParentPolicy.Relationship == RelatedPolicyType.Renewal && this.policy.ParentPolicy.PolicyNumber != null)
                        {
                            this.docManager.InsertValue(string.Concat("Renewal of:		", this.policy.ParentPolicy.PolicyNumber), control, string.Empty);
                        }
                        else
                        {
                            control.Remove();
                        }

                        break;

                    case "_Premium":
                        string premium = string.Empty;
                        controlVal = controlId.Substring(pointIndex + 1);
                        if (policyDict.ContainsKey(controlVal))
                        {
                            premium = policyDict[controlVal];
                        }
                        else
                        {
                            premium = StringExtension.GetPropValue<string>(this.policy, controlVal + ".AgentAdjustedPremium");
                            if (premium == null || !(Convert.ToDecimal(premium) > 0))
                            {
                                premium = StringExtension.GetPropValue<string>(this.policy, controlVal + ".AgentAdjustedPremium");
                            }

                            if (premium == null || !(Convert.ToDecimal(premium) > 0))
                            {
                                premium = StringExtension.GetPropValue<string>(this.policy, controlVal + ".RollUp.Premium");
                            }

                            string minPremium = StringExtension.GetPropValue<string>(this.policy, controlVal + ".RollUp.IsMinimumPremium");
                            if (minPremium == "True")
                            {
                                minPremium = "\t MP";
                            }
                            else
                            {
                                minPremium = string.Empty;
                            }

                            policyDict.Add(controlVal + "MP", minPremium);
                            policyDict.Add(controlVal, premium);
                        }

                        if (premium != null && ((premium.IsNumeric() && Convert.ToDecimal(premium) > 0) || !premium.IsNumeric()))
                        {
                            this.docManager.InsertValue(premium, control, policyDict[controlVal + "MP"]);
                        }
                        else if (this.docManager.IsBlockControl(control))
                        {
                            control.Remove();
                        }

                        break;

                    case "_Property":
                        this.docManager.InsertBuildingsTable(this.policy.CfLine.RiskUnits.OrderBy(x => x.PremisesNumber).ThenBy(x => x.BuildingNumber), control);
                        control.Remove();
                        break;

                    case "_Questions":
                        // _Questions.OCPLine.RiskUnits[0]:Location of Operations
                        string questVal = controlId.Substring(controlId.IndexOf(':') + 1);
                        controlVal = (pointIndex == -1) ? string.Empty : controlId.Substring(pointIndex + 1, controlId.IndexOf(':') - pointIndex - 1);
                        if (controlVal == "OCPLine.RiskUnits[0]")
                        {
                            List<DecisionModel.Models.Policy.Question> quesObj = this.policy.OCPLine.RiskUnits[0].Questions;
                            string answerVal = string.Empty;
                            var question = quesObj.Find(x => x.Code == questVal);
                            if (question != null)
                            {
                                answerVal = question.AnswerValue ?? string.Empty;
                            }
                            else
                            {
                                string error = " Question is not found for control [" + control.Tag + "] in transaction [" + this.policy.SubmissionNumber + "].";
                                LogEvent.Log.Error(error);
                                LogEvent.TraceLog.Error(error);
                            }

                            this.docManager.InsertValue(answerVal, control, string.Empty);
                        }
                        else if (controlVal.StartsWith("ImLine"))
                        {
                            string riskTypeStr = controlVal.Substring(controlVal.IndexOf("."));
                            IMClassType riskType = (IMClassType)Enum.Parse(typeof(IMClassType), riskTypeStr);
                            var riskUnit = this.policy.ImLine.RiskUnits.Find(x => x.ImClassType == riskType);
                            List<DecisionModel.Models.Policy.Question> quesObj = riskUnit.Questions;
                            string answerVal = string.Empty;
                            var question = quesObj.Find(x => x.Code == questVal);
                            if (question != null)
                            {
                                answerVal = question.AnswerValue ?? string.Empty;
                            }
                            else
                            {
                                string error = " Question is not found for control [" + control.Tag + "] in transaction [" + this.policy.SubmissionNumber + "].";
                                LogEvent.Log.Error(error);
                                LogEvent.TraceLog.Error(error);
                            }

                            this.docManager.InsertValue(answerVal, control, string.Empty);
                        }

                        break;

                    case "_specEventRisk":
                        controlVal = controlId.Substring(pointIndex + 1);
                        string eventVal = string.Empty;
                        string eventSuffix = string.Empty;
                        if (policyDict.ContainsKey(controlId))
                        {
                            eventVal = policyDict[controlId];
                        }
                        else
                        {
                            eventVal = StringExtension.GetPropValue<string>(this.policy.SpecialEventLine.RiskUnits[0], controlVal);
                            policyDict.Add(controlId, eventVal);
                        }

                        this.InsertValue(eventVal, control, eventSuffix);
                        break;

                    case "_SupplementalDocuments":
                        if (this.policy.RollUp != null)
                        {
                            List<DecisionModel.Models.Policy.Document> supplimentDocs = (List<DecisionModel.Models.Policy.Document>)this.policy.RollUp.Documents;
                            if (supplimentDocs != null)
                            {
                                supplimentDocs = supplimentDocs.FindAll(x => x.Section == DocumentSection.Application && x.ReturnIsSelected());
                            }

                            this.docManager.InsertLinkListTable(supplimentDocs, control, true, isBindLetter);
                        }

                        control.Remove();
                        break;

                    case "_Taxes":
                        if (this.policy.TaxesAndFees != null && this.policy.TaxesAndFees.Count > 0)
                        {
                            this.docManager.InsertNameAmountTable(this.policy.TaxesAndFees.OrderBy(x => x.Order), control);
                        }

                        control.Remove();
                        break;

                    case "_Today":
                        string valDate = DateTime.Now.ToString("MMMM d, yyyy");
                        this.docManager.InsertValue(valDate, control, string.Empty);
                        break;

                    case "_UnderlyingCoverage":
                        controlVal = (pointIndex == -1) ? string.Empty : controlId.Substring(pointIndex + 1);
                        List<Limit> limits = new List<Limit>();
                        XsUnderlyingCoverage underlyingCoverage = new XsUnderlyingCoverage();
                        if (controlVal == "GL")
                        {
                            underlyingCoverage = this.policy.XsLine.XsGeneralLiabilityCoverage;
                            limits.Add(this.policy.XsLine.XsGeneralLiabilityCoverage.PerOccurrenceLimit);
                            limits.Add(this.policy.XsLine.XsGeneralLiabilityCoverage.GeneralAggregateLimit);
                            limits.Add(this.policy.XsLine.XsGeneralLiabilityCoverage.ProductOperationsLimit);
                            if (this.policy.XsLine.XsGeneralLiabilityCoverage.PersonalAndAdvertisingLimit != null)
                            {
                                limits.Add(this.policy.XsLine.XsGeneralLiabilityCoverage.PersonalAndAdvertisingLimit);
                            }

                            if (this.ShouldDisplayEmployeeBenefitsLimit(this.policy.XsLine.XsGeneralLiabilityCoverage.EmployeeBenefitsLimit))
                            {
                                limits.Add(this.policy.XsLine.XsGeneralLiabilityCoverage.EmployeeBenefitsLimit);
                            }
                        }
                        else if (controlVal == "AUTO")
                        {
                            underlyingCoverage = this.policy.XsLine.XsAutoLiabilityCoverage;
                            limits.Add(this.policy.XsLine.XsAutoLiabilityCoverage.CombinedSingleLimit);
                        }
                        else if (controlVal == "HNOA")
                        {
                            underlyingCoverage = this.policy.XsLine.XsHiredNonOwnedCoverage;
                            limits.Add(this.policy.XsLine.XsHiredNonOwnedCoverage.CombinedSingleLimit);
                        }
                        else if (controlVal == "EL")
                        {
                            underlyingCoverage = this.policy.XsLine.XsEmployersLiabilityCoverage;
                            limits.Add(this.policy.XsLine.XsEmployersLiabilityCoverage.EachAccident);
                            limits.Add(this.policy.XsLine.XsEmployersLiabilityCoverage.PolicyLimitAggregate);
                            limits.Add(this.policy.XsLine.XsEmployersLiabilityCoverage.EachEmployee);
                        }

                        if (underlyingCoverage.IsActive)
                        {
                            this.docManager.InsertUnderlyingCoverageTable(limits, underlyingCoverage, control);
                        }

                        control.Remove();
                        break;

                    case "_Warranties":
                        if (this.policy.RollUp.Warranties != null)
                        {
                            IEnumerable<Warranty> warranties = this.policy.RollUp.Warranties.FindAll(x => x.IsSelected || x.IsManuallySelected).OrderBy(x => x.WarrantyType);
                            if (warranties.Count() > 0)
                            {
                                this.docManager.InsertWarrantiesTable(warranties, control);
                            }
                        }

                        control.Remove();
                        break;

                    case "_xsAutoLiabilityVehicleType":
                        if (this.policy.XsLine.XsAutoLiabilityCoverage.VehicleTypes != null)
                        {
                            List<XsAutoLiabilityVehicleType> list = this.policy.XsLine.XsAutoLiabilityCoverage.VehicleTypes;
                            if (list.Count > 0)
                            {
                                this.docManager.InsertXSAutoLiabilityVehicleType(list.OrderBy(x => x.Order), control);
                            }
                        }

                        control.Remove();
                        break;

                    case "_UnderwriterToAgentNotes":
                        string note = string.Empty;
                        note = this.policy.Referral.DeclinationReasonText + ": " + this.policy.Referral.UnderwriterToAgentNotes;
                        this.docManager.InsertAgentToUnderwriterNotes(note, control);
                        control.Remove();

                        break;

                    case "_xsRisk":
                        controlVal = controlId.Substring(pointIndex + 1);
                        string xsVal = string.Empty;
                        string xsSuffix = string.Empty;
                        if (policyDict.ContainsKey(controlId))
                        {
                            xsVal = policyDict[controlId];
                        }
                        else
                        {
                            xsVal = StringExtension.GetPropValue<string>(this.policy.XsLine.RiskUnits[0], controlVal);
                            policyDict.Add(controlId, xsVal);
                        }

                        this.InsertValue(xsVal, control, xsSuffix);
                        break;

                    default:
                        string val = string.Empty;
                        string suffix = string.Empty;
                        if (policyDict.ContainsKey(controlId))
                        {
                            val = policyDict[controlId];
                        }
                        else
                        {
                            val = StringExtension.GetPropValue<string>(this.policy, controlId);
                            policyDict.Add(controlId, val);
                            if (controlId.EndsWith(".Premium"))
                            {
                                policyDict.Add(controlId + "MP", StringExtension.IsMinPremium.ToString());
                            }
                        }

                        if (controlId.EndsWith(".Premium"))
                        {
                            suffix = policyDict[controlId + "MP"];
                        }

                        this.InsertValue(val, control, suffix);
                        break;
                }
            }

            success = true;

            return success;
        }

        /// <summary>
        /// Keeps the first.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        private void KeepFirst(string controlId)
        {
            var controls = this.docManager.GetControls().FindAll(x => x.Tag == controlId);
            if (controls.Count > 1)
            {
                foreach (var control in controls)
                {
                    if (control != controls[0])
                    {
                        control.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Removes control unless Effective Date is before dateVal.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        /// <param name="dateVal">The date value.</param>
        /// <returns>if control was not removed</returns>
        private bool EffectiveBefore(string controlId, string dateVal)
        {
            var controls = this.docManager.GetControls().FindAll(x => x.Tag == controlId);
            if ((DateTime)this.policy.EffectiveDate >= DateTime.Parse(dateVal))
            {
                foreach (var control in controls)
                {
                    control.Remove();
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Removes control unless Effective Date is after or equal to dateVal.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        /// <param name="dateVal">The date value.</param>
        /// <returns>if control was not removed</returns>
        private bool EffectiveAfter(string controlId, string dateVal)
        {
            var controls = this.docManager.GetControls().FindAll(x => x.Tag == controlId);
            if ((DateTime)this.policy.EffectiveDate < DateTime.Parse(dateVal))
            {
                foreach (var control in controls)
                {
                    control.Remove();
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Gets the type of the im risk by.
        /// </summary>
        /// <param name="riskTypeStr">The risk type string.</param>
        /// <returns>Risk Unit</returns>
        private ImRiskUnit GetImRiskByType(string riskTypeStr)
        {
            IMClassType riskType = (IMClassType)Enum.Parse(typeof(IMClassType), riskTypeStr);
            return this.policy.ImLine.RiskUnits.Find(x => x.ImClassType == riskType);
        }

        /// <summary>
        /// Keeps the last.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        private void KeepLast(string controlId)
        {
            var controls = this.docManager.GetControls().FindAll(x => x.Tag == controlId);
            if (controls.Count > 1)
            {
                foreach (var control in controls)
                {
                    if (control != controls[controls.Count - 1])
                    {
                        control.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Handles the coverage controls.
        /// </summary>
        /// <param name="controlVal">The control value.</param>
        /// <param name="control">The control.</param>
        /// <param name="covObj">The coverage object.</param>
        /// <returns>boolean</returns>
        private bool HandleCoverageControls(string controlVal, Aspose.Words.Markup.StructuredDocumentTag control, List<DecisionModel.Models.Policy.OptionalCoverage> covObj)
        {
            if (covObj.Find(x => x.Code == controlVal) != null)
            {
                OptionalCoverage cov = covObj.Find(x => x.Code == controlVal);
                if (cov.IsSelected)
                {
                    string strAmount;
                    if (cov.IsIncluded)
                    {
                        strAmount = "Included";
                    }
                    else
                    {
                        strAmount = "$" + (cov.AdjustedPremium.HasValue ? cov.AdjustedPremium.Value.ToString("N0") : cov.Premium.ToString("N0"));
                    }

                    control.Range.Replace(new Regex("_" + controlVal + ".Premium"), strAmount, new FindReplaceOptions(FindReplaceDirection.Forward));
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// IM risk.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        /// <param name="control">The control.</param>
        /// <param name="policyDict">The policy dictionary.</param>
        private void ImRisk(string controlId, StructuredDocumentTag control, Dictionary<string, string> policyDict)
        {
            string controlVal = controlId.Substring(controlId.IndexOf(".") + 1);
            string riskTypeStr = (controlVal.IndexOf(".") == -1) ? controlVal : controlVal.Substring(0, controlVal.IndexOf("."));
            controlVal = controlVal.Substring(controlVal.IndexOf(".") + 1);
            string imVal = string.Empty;
            string imSuffix = string.Empty;
            if (policyDict.ContainsKey(controlId))
            {
                imVal = policyDict[controlId];
            }
            else
            {
                IMClassType riskType = (IMClassType)Enum.Parse(typeof(IMClassType), riskTypeStr);
                var riskUnit = this.policy.ImLine.RiskUnits.Find(x => x.ImClassType == riskType);
                if (riskUnit != null)
                {
                    imVal = StringExtension.GetPropValue<string>(riskUnit, controlVal);
                    imSuffix = StringExtension.IsMinPremium.ToString();
                }

                policyDict.Add(controlId, imVal);
            }

            this.InsertValue(imVal, control, imSuffix);
        }

        /// <summary>
        /// IM risk Value.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        /// <returns>value</returns>
        private string ImRiskValue(string controlId)
        {
            string controlVal = controlId.Substring(controlId.IndexOf(".") + 1);
            string riskTypeStr = (controlVal.IndexOf(".") == -1) ? controlVal : controlVal.Substring(0, controlVal.IndexOf("."));
            controlVal = controlVal.Substring(controlVal.IndexOf(".") + 1);
            string imVal = string.Empty;
            IMClassType riskType = (IMClassType)Enum.Parse(typeof(IMClassType), riskTypeStr);
            var riskUnit = this.policy.ImLine.RiskUnits.Find(x => x.ImClassType == riskType);
            if (riskUnit != null)
            {
                imVal = StringExtension.GetPropValue<string>(riskUnit, controlVal);
            }

            return imVal ?? string.Empty;
        }

        /// <summary>
        /// Returns the value for the IM item passed
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        /// <returns>Value string</returns>
        private string ImItemValue(string controlId)
        {
            string controlVal = controlId.Substring(controlId.IndexOf(".") + 1);
            string riskTypeStr = (controlVal.IndexOf(".") == -1) ? controlVal : controlVal.Substring(0, controlVal.IndexOf("."));
            controlVal = controlVal.Substring(controlVal.IndexOf(".") + 1);
            string itemTypeStr = (controlVal.IndexOf(".") == -1) ? controlVal : controlVal.Substring(0, controlVal.IndexOf("."));
            controlVal = controlVal.Substring(controlVal.IndexOf(".") + 1);
            string imVal = string.Empty;
            IMClassType riskType = (IMClassType)Enum.Parse(typeof(IMClassType), riskTypeStr);
            var riskUnit = this.policy.ImLine.RiskUnits.Find(x => x.ImClassType == riskType);
            if (riskUnit != null)
            {
                var imItem = riskUnit.Items.Find(x => x.ItemCode == itemTypeStr);
                if (imItem != null)
                {
                    imVal = StringExtension.GetPropValue<string>(imItem, controlVal);
                }
            }

            return imVal;
        }

        /// <summary>
        /// Inserts the value.
        /// </summary>
        /// <param name="val">The value.</param>
        /// <param name="control">The control.</param>
        /// <param name="suffix">The suffix.</param>
        private void InsertValue(string val, StructuredDocumentTag control, string suffix)
        {
            if (val != null && ((val.IsNumeric() && Convert.ToDecimal(val) > 0) || !val.IsNumeric()))
            {
                this.docManager.InsertValue(val, control, suffix);
            }
            else if (this.docManager.IsBlockControl(control))
            {
                control.Remove();
            }
            else if (val == null && control.GetText() == "$00")
            {
                this.docManager.InsertText("$0", control);
            }
            else if (control.GetText() == "00%")
            {
                this.docManager.InsertText("0%", control);
            }
        }

        /// <summary>
        /// Returns the coverages.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        /// <returns>List of OptionalCoverages</returns>
        private List<OptionalCoverage> ReturnCoverages(string controlId)
        {
            List<OptionalCoverage> polCovObj = null;
            string[] objArray = controlId.Split('.');
            if (objArray[0] == "GlLine")
            {
                if (this.policy.GlLine != null && this.policy.GlLine.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.GlLine.RollUp.OptionalCoverages;
                }
            }
            else if (objArray[0] == "OCPLine")
            {
                if (this.policy.OCPLine != null && this.policy.OCPLine.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.OCPLine.RollUp.OptionalCoverages;
                }
            }
            else if (objArray[0] == "CfLine")
            {
                if (this.policy.CfLine != null && this.policy.CfLine.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.CfLine.RollUp.OptionalCoverages;
                }
            }
            else if (objArray[0] == "LiquorLine")
            {
                if (this.policy.LiquorLine != null && this.policy.LiquorLine.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.LiquorLine.RollUp.OptionalCoverages;
                }
            }
            else if (objArray[0] == "XsLine")
            {
                if (this.policy.XsLine != null && this.policy.XsLine.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.XsLine.RollUp.OptionalCoverages;
                }
            }
            else if (objArray[0] == "ImLine")
            {
                if (this.policy.ImLine != null && this.policy.ImLine.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.ImLine.RollUp.OptionalCoverages;
                }
            }
            else if (objArray[0] == "SpecialEventLine")
            {
                if (this.policy.SpecialEventLine != null && this.policy.SpecialEventLine.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.SpecialEventLine.RollUp.OptionalCoverages;
                }
            }
            else if (objArray[0] == "Policy")
            {
                if (this.policy.RollUp.OptionalCoverages != null)
                {
                    polCovObj = this.policy.RollUp.OptionalCoverages;
                }
            }

            return polCovObj;
        }

        /// <summary>
        /// Gets the quote bind templates.
        /// </summary>
        /// <param name="isBind">if set to <c>true</c> [is bind].</param>
        /// <returns>string array</returns>
        private string[] GetQuoteBindTemplates(bool isBind = false)
        {
            List<string> docs = new List<string>();

            if (this.policy.LetterTemplates.Count > 0)
            {
                if (isBind)
                {
                    docs = this.policy.LetterTemplates.Where(x => x.AssignedLetter == LetterType.BindLetter || x.AssignedLetter == LetterType.QuoteAndBindLetter).OrderBy(x => x.Order).Select(x => this.dataPath + "/" + x.FileName).ToList();
                }
                else
                {
                    docs = this.policy.LetterTemplates.Where(x => x.AssignedLetter == LetterType.QuoteLetter || x.AssignedLetter == LetterType.QuoteAndBindLetter).OrderBy(x => x.Order).Select(x => this.dataPath + "/" + x.FileName).ToList();
                }
            }

            return docs.ToArray();
        }

        /// <summary>
        /// Returns the image.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="displayFile">The File.</param>
        /// <returns>An Image</returns>
        private HttpResponseMessage ReturnPNGImage(string fileName, object displayFile)
        {
            string savePath = System.Web.Hosting.HostingEnvironment.MapPath(Path.Combine(this.tempPath, this.folderId));
            HttpResponseMessage response = this.Request.CreateResponse();
            if (displayFile != null)
            {
                System.IO.MemoryStream docStream = new System.IO.MemoryStream();
                this.docManager.Save(docStream, DocumentManager.FileFormat.PNG);
                docStream.Position = 0;
                response.Content = new StreamContent(docStream);
                response.Content = new ByteArrayContent((byte[])displayFile);
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(MimeMapping.GetMimeMapping(fileName + ".png"));
                response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
                response.Content.Headers.ContentDisposition.FileName = fileName;
            }

            return response;
        }

        /// <summary>
        /// Returns the risk units.
        /// </summary>
        /// <param name="controlId">The control identifier.</param>
        /// <returns>numerable list of risk units</returns>
        private IEnumerable<IBaseGlRiskUnit> ReturnRiskUnits(string controlId)
        {
            IEnumerable<IBaseGlRiskUnit> riskObj = null;
            string[] objArray = controlId.Split('.');
            if (objArray[0] == "GlLine")
            {
                if (this.policy.GlLine != null && this.policy.GlLine.RiskUnits != null)
                {
                    riskObj = this.policy.GlLine.RiskUnits;
                }
            }
            else if (objArray[0] == "OCPLine")
            {
                if (this.policy.OCPLine != null && this.policy.OCPLine.RiskUnits != null)
                {
                    riskObj = this.policy.OCPLine.RiskUnits;
                }
            }
            else if (objArray[0] == "LiquorLine")
            {
                if (this.policy.LiquorLine != null && this.policy.LiquorLine.RiskUnits != null)
                {
                    riskObj = this.policy.LiquorLine.RiskUnits;
                }
            }
            else if (objArray[0] == "SpecialEventLine")
            {
                if (this.policy.SpecialEventLine != null && this.policy.SpecialEventLine.RiskUnits != null)
                {
                    riskObj = this.policy.SpecialEventLine.RiskUnits;
                }
            }

            return riskObj;
        }

        /// <summary>
        /// Parses if statement.
        /// </summary>
        /// <param name="condition">The condition.</param>
        /// <param name="control">The control.</param>
        /// <returns>string array of left value, operator, right value </returns>
        private string[] ParseIfStatement(string condition, StructuredDocumentTag control)
        {
            string option = string.Empty;
            if (condition.IndexOf('!') != -1)
            {
                option = "!";
                condition = condition.Replace("!", string.Empty);
            }

            if (condition.IndexOf('=') != -1)
            {
                option += "=";
            }
            else if (condition.IndexOf('>') != -1)
            {
                option += ">";
            }
            else if (condition.IndexOf('<') != -1)
            {
                option += "<";
            }

            char[] splitArray = new char[] { '=', '>', '<' };
            string[] formula = condition.Split(splitArray);
            string[] values = { formula[0], option, formula[1] };
            return values;
        }

        /// <summary>
        /// Equates the passed values base on the option and returns true of false
        /// </summary>
        /// <param name="value1">The value1.</param>
        /// <param name="option">The option.</param>
        /// <param name="value2">The value2.</param>
        /// <returns>whether the equation was true or false</returns>
        private bool Equation(string value1, string option, string value2)
        {
            if (option == "=")
            {
                if (value2 == "null")
                {
                    return value1 == null;
                }
                else
                {
                    return value1 == value2;
                }
            }
            else if (option == "!=")
            {
                if (value2 == "null")
                {
                    return value1 != null;
                }
                else
                {
                    return value1 != value2;
                }
            }
            else if (option == ">")
            {
                return Convert.ToDecimal(value1) > Convert.ToDecimal(value2);
            }
            else if (option == "!>")
            {
                return !(Convert.ToDecimal(value1) > Convert.ToDecimal(value2));
            }
            else if (option == "<")
            {
                return Convert.ToDecimal(value1) < Convert.ToDecimal(value2);
            }
            else if (option == "!<")
            {
                return !(Convert.ToDecimal(value1) < Convert.ToDecimal(value2));
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Gets results of the equation(true false) and passes it on for the results to be evaluated.
        /// </summary>
        /// <param name="compareLeft">The compare left.</param>
        /// <param name="option">The option.</param>
        /// <param name="compareRight">The compare right.</param>
        /// <param name="control">The control.</param>
        private void EvaluateEquation(string compareLeft, string option, string compareRight, StructuredDocumentTag control)
        {
            this.EvaluateEquation(this.Equation(compareLeft, option, compareRight), control);
        }

        /// <summary>
        /// Evaluates the expression and Inserts the appropriate value into the control
        /// </summary>
        /// <param name="result">if set to <c>true</c> [result].</param>
        /// <param name="control">The control.</param>
        private void EvaluateEquation(bool result, StructuredDocumentTag control)
        {
            string returnVal = string.Empty;
            string expression = string.Empty;
            List<string> expressions = new List<string>();
            string controlText = control.GetText();
            int j = 0;
            for (int i = 0; i < controlText.Length; ++i)
            {
                if (controlText[i] == ':' && (i == 0 || controlText[i - 1] != '\\'))
                {
                    expressions.Add(controlText.Substring(j, i - j));
                    j = i + 1;
                }
            }

            expressions.Add(controlText.Substring(j));
            if (result)
            {
                expression = expressions[0];
            }
            else
            {
                expression = (expressions.Count > 1) ? expressions[1] : string.Empty;
            }

            expression = expression.Replace("\\:", ":");
            if (expression.IndexOf("_") == 0)
            {
                expression = expression.Replace("_", string.Empty);
                string[] values = expression.Split(new char[] { ' ' }, 2);
                returnVal = this.GetPolicyValue(values[0]);
                if (values.Count() == 2)
                {
                    this.docManager.InsertValue(returnVal, control, string.Empty, values[1]);
                }
                else
                {
                    control.Range.Replace(new Regex("_" + values[0]), " ", new FindReplaceOptions(FindReplaceDirection.Forward));
                    this.docManager.InsertValue(returnVal, control);
                }
            }
            else
            {
                if (expression == string.Empty)
                {
                    control.Remove();
                }
                else
                {
                    this.docManager.InsertValue(returnVal, control, string.Empty, expression);
                }
            }
        }

        /// <summary>
        /// Gets the policy value.
        /// </summary>
        /// <param name="obj">The string indicating where in the policy object.</param>
        /// <returns>value</returns>
        private string GetPolicyValue(string obj)
        {
            if (obj.StartsWith("ImRisk"))
            {
                return this.ImRiskValue(obj);
            }
            else if (obj.StartsWith("ImItem"))
            {
                return this.ImItemValue(obj);
            }
            else if (obj.StartsWith("XsRisk"))
            {
                obj = obj.Substring(obj.IndexOf(".") + 1);
                return StringExtension.GetPropValue<string>(this.policy.XsLine.RiskUnits[0], obj);
            }
            else if (obj.StartsWith("SpecEventRisk"))
            {
                obj = obj.Substring(obj.IndexOf(".") + 1);
                return StringExtension.GetPropValue<string>(this.policy.SpecialEventLine.RiskUnits[0], obj);
            }
            else
            {
                return StringExtension.GetPropValue<string>(this.policy, obj);
            }
        }

        /// <summary>
        /// Saves the file to the temp/[folderId] folder.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="docFormat">The document format.</param>
        /// <returns>
        /// Boolean
        /// </returns>
        private bool SaveToFile(string fileName, DocumentManager.FileFormat docFormat)
        {
            if (Path.GetExtension(fileName).ToLower().TrimStart('.') != docFormat.ToString().ToLower())
            {
                fileName = fileName + '.' + docFormat.ToString().ToLower();
            }

            string savePath = System.Web.Hosting.HostingEnvironment.MapPath(Path.Combine(this.tempPath, this.folderId));
            if (this.docManager != null)
            {
                this.docManager.Save(savePath, fileName, docFormat);
            }

            // change this to use SaveParams as the parameters to utilize when saving and downloading
            DownloadParameters saveParams = new DownloadParameters()
            {
                SaveFormat = docFormat,
                TempFileName = System.IO.Path.Combine(savePath, fileName)
            };
            return true;
        }

        /// <summary>
        /// Determines if Employee benefits limit should be displayed on quote/bind letters.
        /// </summary>
        /// <param name="employeeBenefitsLimit">Employee benefits limit object</param>
        /// <returns>True if it should be displayed else false.</returns>
        private bool ShouldDisplayEmployeeBenefitsLimit(Limit employeeBenefitsLimit)
        {
            var result = false;

            if (employeeBenefitsLimit != null && !string.IsNullOrWhiteSpace(employeeBenefitsLimit.Code) && !(employeeBenefitsLimit.Code.Equals("N/A") || employeeBenefitsLimit.Code.Equals("Other")))
            {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// This method is for testing the Issuance document generation.
        /// This can be removed once the testing is complete
        /// </summary>
        private void GenerateIssuanceDocument()
        {
            ////Copy all forms that needs to be tested at this location
            var formsPath = @"C:\Forms";
            var mergedForms = new List<MergedForm>();

            string[] forms = File.ReadAllLines(Path.Combine(formsPath, "Forms.txt"));
            foreach (var form in forms)
            {
                mergedForms.Add(new MergedForm()
                {
                    FormNormalizedNumber = form.Replace(Environment.NewLine, string.Empty),
                    Name = form.Replace(Environment.NewLine, string.Empty),
                    Path = string.Empty,
                    Templated = true
                });
            }

            var mergeFormRequest = new DecisionModel.Representations.MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            this.issuanceController.SaveDocs(mergeFormRequest, DocumentManager.FileFormat.PDF, true);

            this.policy.WorkflowState = WorkflowState.PreIssue;

            this.issuanceController.SaveDocs(mergeFormRequest, DocumentManager.FileFormat.PDF, true);

            this.issuanceController.SaveConsolidatedPolicyDocument(this.policy);
        }

        /// <summary>
        /// Executes an api service request
        /// </summary>
        /// <param name="requestUri">The request uri</param>
        /// <param name="requestMethod">The request method</param>
        /// <param name="serviceUrl">The service endpoint</param>
        /// <returns>string</returns>
        /// <exception cref="System.Web.HttpException">Service Unavailable;500</exception>
        private string ExecuteRequest(string requestUri, Method requestMethod, string serviceUrl)
        {
            var request = new RestRequest(requestUri, requestMethod);

            RestClient client = new RestClient(serviceUrl);
            var response = (requestMethod == Method.DELETE) ?
                                client.ExecuteAsPost(request, "DELETE") :
                                client.Execute(request);

            if (response.StatusCode == HttpStatusCode.ServiceUnavailable)
            {
                throw new HttpException("Service Unavailable", 500);
            }

            return response.Content;
        }
    }
}
