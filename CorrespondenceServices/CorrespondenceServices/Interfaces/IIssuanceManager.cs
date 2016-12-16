// <copyright file="IIssuanceManager.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Interfaces
{
    using System;
    using System.Collections.Generic;
    using DecisionModel.Models.Policy;
    using DecisionModel.Representations;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Models;
    using Mkl.WebTeam.SubmissionShared.Enumerations;

    /// <summary>
    /// Interface to IssuanceManager
    /// </summary>
    public interface IIssuanceManager
    {
        /// <summary>
        /// Create default template mapping
        /// </summary>
        /// <param name="authenticatedUser">User</param>
        /// <param name="effectiveDate">DateTime</param>
        void CreateDefaultTemplateMapping(IUser authenticatedUser, DateTime? effectiveDate);

        /// <summary>
        /// Get the generated documents out of merged forms and template mappings
        /// </summary>
        /// <param name="authenticatedUser">User</param>
        /// <param name="mergedForms">List of merged forms</param>
        /// <param name="effectiveDate">DateTime</param>
        /// <param name="fileFormat">fileFormat</param>
        /// <param name="removeDocs">bool</param>
        /// <returns>MergeDocumentsResponse</returns>
        MergeDocumentsResponse ProcessMergedFormTemplates(IUser authenticatedUser, List<MergedForm> mergedForms, DateTime? effectiveDate, DocumentManager.FileFormat fileFormat, bool removeDocs);

        /// <summary>
        /// Get the signature file name
        /// </summary>
        /// <param name="authenticatedUser">User</param>
        /// <param name="signatureType">CorporateSignatureType</param>
        /// <param name="effectiveDate">Effective date from the policy</param>
        /// <param name="stateCode">State code of the producer</param>
        /// <returns>Signature file name and path</returns>
        string GetSignatureFileName(IUser authenticatedUser, CorporateSignatureType signatureType, DateTime? effectiveDate, string stateCode = "");

        /// <summary>
        /// Consolidate all the issuance documents of the policy into one document and upload it to storage provider.
        /// </summary>
        /// <returns>Document Upload response</returns>
        UploadDocumentResponse GenerateConsolidatedPolicyDocument();
    }
}
