// <copyright file="IssuanceController.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Controllers
{
    using System;
    using System.Diagnostics;
    using System.Linq;
    using System.Web.Http;
    using CorrespondenceServices.Classes;
    using CorrespondenceServices.Interfaces;
    using DecisionModel.Models.Policy;
    using DecisionModel.Representations;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.WebCore2;

    /// <summary>
    /// Issuance controller
    /// </summary>
    [RoutePrefix("api/Issuance")]
    public class IssuanceController : BaseApiController
    {
        /// <summary>
        /// issuance manager
        /// </summary>
        private IIssuanceManager issuanceManager;

        /// <summary>
        /// The policy
        /// </summary>
        private Policy policy;

        /// <summary>
        /// Initializes a new instance of the <see cref="IssuanceController" /> class.
        /// </summary>
        /// <param name="bootStrapper">IBootstrapper</param>
        /// <param name="jsonManager">IJSONManager</param>
        /// <param name="restHelper">The rest helper.</param>
        /// <param name="storageManager">The storage manager.</param>
        /// <param name="templateProcessor">The template processor.</param>
        public IssuanceController(
            IBootstrapper bootStrapper,
            IJsonManager jsonManager,
            IRestHelper restHelper,
            IStorageManager storageManager,
            ITemplateProcessor templateProcessor)
            : base(jsonManager, bootStrapper, restHelper, storageManager, templateProcessor)
        {
        }

        /// <summary>
        /// Generates and Saves the documents/merged forms with its field values to the store.
        /// </summary>
        /// <param name="mergedFormsrequest">MergeFormsRequest</param>
        /// <param name="fileFormat">fileFormat</param>
        /// <param name="removeDocs">bool</param>
        /// <returns>array of string</returns>
        /// <exception cref="ArgumentException">Thrown if the document parameters are invalid.</exception>
        [HttpPost]
        [Route("SaveDocs")]
        public MergeDocumentsResponse SaveDocs(
            DecisionModel.Representations.MergeFormsRequest mergedFormsrequest,
            [FromUri] DocumentManager.FileFormat fileFormat,
            [FromUri] bool removeDocs)
        {
            this.policy = mergedFormsrequest.Policy;

            Stopwatch stopWatch = new Stopwatch();
            try
            {
                stopWatch.Start();
                var policyCarrier = this.GetPolicyCarrier(mergedFormsrequest.Policy.CarrierName);
                this.issuanceManager = new IssuanceManager(this.RestHelper, this.policy, policyCarrier, this.TemplateProcessor, this.StorageProviderManager);
                MergeDocumentsResponse mergeDocumentsResponse = this.issuanceManager.ProcessMergedFormTemplates(this.AuthenticatedUser, mergedFormsrequest.MergedForms, mergedFormsrequest.Policy.EffectiveDate, fileFormat, removeDocs);

                return mergeDocumentsResponse;
            }
            catch (Exception e)
            {
                string error = "SaveDocs Issuance Error. Transaction: " + this.policy.SubmissionNumber + ". Error message: " + e.Message;
                LogEvent.Log.Error(error);
                LogEvent.TraceLog.Error(error, e);
                throw new ArgumentException(error, e);
            }
            finally
            {
                stopWatch.Stop();
                LogEvent.Log.Info("Time required to process templates and generate merged forms is " + stopWatch.Elapsed.TotalSeconds + " s");
            }
        }

        /// <summary>
        /// Generate the consolidated policy document and save it to storage provider.
        /// </summary>
        /// <param name="policy">Policy object</param>
        /// <returns>Path of the consolidated document policy.</returns>
        [HttpPost]
        [Route("SaveConsolidatedPolicyDocument")]
        public MergedForm SaveConsolidatedPolicyDocument(Policy policy)
        {
            this.policy = policy;

            try
            {
                var policyCarrier = this.GetPolicyCarrier(policy.CarrierName);
                this.issuanceManager = new IssuanceManager(this.RestHelper, this.policy, policyCarrier, this.TemplateProcessor, this.StorageProviderManager);
                var mergeDocumentsResponse = this.issuanceManager.GenerateConsolidatedPolicyDocument();
                return new MergedForm()
                {
                    CompressionLevel = 0,
                    Name = mergeDocumentsResponse.FileName,
                    Path = mergeDocumentsResponse.DirectoryPath,
                    Templated = false,
                    ErrorMessages = mergeDocumentsResponse.Errors
                };
            }
            catch (Exception e)
            {
                string error = "SaveConsolidatedPolicyDocument Issuance Error. Transaction: " + this.policy.SubmissionNumber + ". Error message: " + e.Message;
                LogEvent.Log.Error(error);
                LogEvent.TraceLog.Error(error, e);
                throw new ArgumentException(error, e);
            }
        }

        /// <summary>
        /// Gets the policy document.
        /// </summary>
        /// <param name="policyId">The policy identifier.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.Byte[].</returns>
        [HttpGet]
        [Route("GetPolicyDocument/{policyId}")]
        public byte[] GetPolicyDocument(string policyId, [FromUri] string fileName)
        {
            ////TODO: Fix this Grey Team
            return this.StorageProviderManager.GetIssuanceDocumentBinary(policyId, fileName);
        }

        /// <summary>
        /// Gets the carrier from carrier collection
        /// </summary>
        /// <param name="carrierName">Name of the carrier</param>
        /// <returns>Carrier</returns>
        private Carrier GetPolicyCarrier(string carrierName)
        {
            return this.Carriers?.FirstOrDefault(x => string.Compare(x.Name, carrierName, StringComparison.InvariantCultureIgnoreCase) == 0);
        }
    }
}