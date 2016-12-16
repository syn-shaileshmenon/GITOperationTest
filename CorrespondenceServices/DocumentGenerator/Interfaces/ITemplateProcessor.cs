// <copyright file="ITemplateProcessor.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator
{
    using System.Collections.Generic;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Interface to template processor
    /// </summary>
    public interface ITemplateProcessor
    {
        /// <summary>
        /// Gets the default Policy dictionary
        /// </summary>
        Dictionary<string, MappingDetail> PolicyDictionary
        {
            get;
        }

        /// <summary>
        /// Gets the custom Policy dictionary
        /// </summary>
        Dictionary<string, MappingDetail> CustomPolicyDictionary
        {
            get;
        }

        /// <summary>
        /// Create the default dictionary
        /// </summary>
        /// <param name="tokens">JToken</param>
        void CreateDefaultDictionary(JToken tokens);

        /// <summary>
        /// Create the custom dictionary
        /// </summary>
        /// <param name="tokens">JToken</param>
        void CreateCustomDictionary(JToken tokens);

        /// <summary>
        /// Retrieve the field value based on bookmark's name from dictionary
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="policyJson">The policy json.</param>
        /// <param name="control">control</param>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <returns>string</returns>
        string GetFieldValue(DecisionModel.Models.Policy.Policy policy, JObject policyJson, Aspose.Words.Bookmark control, ref IPolicyDocumentManager docManager);

        /// <summary>
        /// Retrieve the field value based on bookmark's name from dictionary or invoke a custom function.
        /// </summary>
        /// <param name="keyName">custom mapping entry key</param>
        /// <param name="policy">Policy</param>
        /// <param name="policyJson">The policy json.</param>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <returns>string</returns>
        string GetFieldValueFromCustom(string keyName, DecisionModel.Models.Policy.Policy policy, JObject policyJson, ref IPolicyDocumentManager docManager);
    }
}
