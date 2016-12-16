// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-11-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-16-2016
// ***********************************************************************
// <copyright file="StorageManagerStub.cs" company="Markel">
//     Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CorrespondenceServices.Tests.Stubs
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.StorageProvider.Models;

    /// <summary>
    /// Class StorageManagerStub.
    /// </summary>
    /// <seealso cref="Mkl.WebTeam.StorageProvider.Interfaces.IStorageManager" />
    [ExcludeFromCodeCoverage]
    public class StorageManagerStub : IStorageManager
    {
        /// <summary>
        /// Gets the agency signature binary.
        /// </summary>
        /// <param name="agencyId">The agency identifier.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.Byte[].</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public byte[] GetAgencySignatureBinary(string agencyId, string fileName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the corporate signature binary.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <returns>System.Byte[].</returns>
        public byte[] GetCorporateSignatureBinary(string filePath)
        {
            return File.ReadAllBytes(filePath);
        }

        /// <summary>
        /// Gets the document.
        /// </summary>
        /// <param name="attachmentId">The attachment identifier.</param>
        /// <returns>FileAttachmentBinary.</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public FileAttachmentBinary GetDocument(Guid attachmentId)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the document.
        /// </summary>
        /// <param name="policyId">The policy identifier.</param>
        /// <param name="filePath">The file path.</param>
        /// <param name="filename">The filename.</param>
        /// <param name="fileMimeType">Type of the file MIME.</param>
        /// <returns>System.Byte[].</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public byte[] GetDocument(Guid policyId, string filePath, string filename, string fileMimeType)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the issuance document binary.
        /// </summary>
        /// <param name="policyId">The policy identifier.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.Byte[].</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public byte[] GetIssuanceDocumentBinary(string policyId, string fileName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the signature list.
        /// </summary>
        /// <param name="agencyId">The agency identifier.</param>
        /// <param name="signatureType">Type of the signature.</param>
        /// <returns>List&lt;SignatureInfo&gt;.</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public List<SignatureInfo> GetSignatureList(string agencyId, SignatureType signatureType)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Removes the document.
        /// </summary>
        /// <param name="attachmentId">The attachment identifier.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public bool RemoveDocument(Guid attachmentId)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Saves the document.
        /// </summary>
        /// <param name="fileAttachmentBinary">The file attachment binary.</param>
        /// <param name="isCallFromMigrationTool">if set to <c>true</c> [is call from migration tool].</param>
        /// <returns>StorageResponse.</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public StorageResponse SaveDocument(FileAttachmentBinary fileAttachmentBinary, bool isCallFromMigrationTool = false)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Saves the document.
        /// </summary>
        /// <param name="fileAttachment">The file attachment.</param>
        /// <param name="isCallFromMigrationTool">if set to <c>true</c> [is call from migration tool].</param>
        /// <returns>StorageResponse.</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public StorageResponse SaveDocument(FileAttachment fileAttachment, bool isCallFromMigrationTool = false)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Uploads the issuance document.
        /// </summary>
        /// <param name="documentToUpload">The document to upload.</param>
        /// <returns>UploadDocumentResponse.</returns>
        public UploadDocumentResponse UploadIssuanceDocument(UploadDocumentRequest documentToUpload)
        {
            var filename = $@"Output\{documentToUpload.DestinationFileName}";
            File.WriteAllBytes(filename, documentToUpload.DocumentBinary);
            var response = new UploadDocumentResponse()
            {
                DirectoryPath = Directory.GetCurrentDirectory() + @"\Output",
                FileName = documentToUpload.DestinationFileName
            };

            return response;
        }

        /// <summary>
        /// Uploads the signature.
        /// </summary>
        /// <param name="signatureToUpload">The signature to upload.</param>
        /// <returns>UploadSignatureResponse.</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public UploadSignatureResponse UploadSignature(UploadSignatureRequest signatureToUpload)
        {
            throw new NotImplementedException();
        }
    }
}