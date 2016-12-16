// <copyright file="TemplateProcessor.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator
{
    using System.Collections.Generic;
    using Aspose.Words;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.DocumentGenerator.Helpers;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Template processor
    /// </summary>
    public class TemplateProcessor : ITemplateProcessor
    {
        /////// <summary>
        /////// Dictionary to contain the mapping details of a field
        /////// </summary>
        ////private Dictionary<string, MappingDetail> policyDict = null;

        /////// <summary>
        /////// Dictionary to contain custom mapping details of a field
        /////// </summary>
        ////private Dictionary<string, MappingDetail> CustomPolicyDict = null;

        /// <summary>
        /// Json manager, any json specific operations should go here
        /// </summary>
        private readonly IJsonManager jsonManager;

        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateProcessor" /> class.
        /// </summary>
        /// <param name="jsonMgr">IJSONManager</param>
        public TemplateProcessor(IJsonManager jsonMgr)
        {
            this.jsonManager = jsonMgr;
        }

        /// <summary>
        /// Gets the default Policy dictionary
        /// </summary>
        public Dictionary<string, MappingDetail> PolicyDictionary
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the custom Policy dictionary
        /// </summary>
        public Dictionary<string, MappingDetail> CustomPolicyDictionary
        {
            get;
            internal set;
        }

        /// <summary>
        /// Retrieve the field value based on bookmark's name from dictionary
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="policyJson">The policy json.</param>
        /// <param name="control">control</param>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <returns>string</returns>
        public string GetFieldValue(Policy policy, JObject policyJson, Bookmark control, ref IPolicyDocumentManager docManager)
        {
            string controlId = control.Name?.ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(controlId) &&
                this.PolicyDictionary.ContainsKey(controlId))
            {
                MappingDetail detail = this.PolicyDictionary[controlId];
                return this.ParseValue(detail, policy, policyJson, ref docManager);
            }

            return string.Empty;
        }

        /// <summary>
        /// Retrieve the field value based on bookmark's name from dictionary or invoke a custom function.
        /// </summary>
        /// <param name="keyName">custom mapping entry key</param>
        /// <param name="policy">Policy</param>
        /// <param name="policyJson">The policy json.</param>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <returns>string</returns>
        public string GetFieldValueFromCustom(string keyName, Policy policy, JObject policyJson, ref IPolicyDocumentManager docManager)
        {
            string suffix = string.Empty;
            if (this.CustomPolicyDictionary.ContainsKey(keyName.ToLowerInvariant()))
            {
                MappingDetail detail = this.CustomPolicyDictionary[keyName.ToLower()];
                return this.ParseValue(detail, policy, policyJson, ref docManager);
            }

            return string.Empty;
        }

        /// <summary>
        /// Create the default dictionary
        /// </summary>
        /// <param name="tokens">JToken</param>
        public void CreateDefaultDictionary(JToken tokens)
        {
            this.PolicyDictionary = this.jsonManager.GenerateMappingDictionary(tokens);
        }

        /// <summary>
        /// Create the custom dictionary
        /// </summary>
        /// <param name="tokens">JToken</param>
        public void CreateCustomDictionary(JToken tokens)
        {
            this.CustomPolicyDictionary = this.jsonManager.GenerateMappingDictionary(tokens);
        }

        /// <summary>
        /// Parse value
        /// </summary>
        /// <param name="detail">MappingDetail</param>
        /// <param name="policy">Policy</param>
        /// <param name="policyJson">The policy json.</param>
        /// <param name="docManager">PolicyDocumentManager</param>
        /// <returns>string</returns>
        private string ParseValue(MappingDetail detail, Policy policy, JObject policyJson, ref IPolicyDocumentManager docManager)
        {
            string val = string.Empty;
            if (detail != null && !detail.HasFunction)
            {
                if (!string.IsNullOrWhiteSpace(detail.Field))
                {
                    val = policy.GetPropValue<string>(detail.Field, detail.Format);
                }
                else if (!string.IsNullOrWhiteSpace(detail.JsonPath))
                {
                    var obj = JsonPathHelper.GetJsonPathValue(policyJson, detail);
                    if (obj != null)
                    {
                        val = obj.ToString();
                    }
                }
                else if (detail.Function == "SUM" ||
                    detail.Function == "IFTRUE" ||
                    detail.Function == "IFNULL" ||
                    detail.Function == "IFNOTNULL" ||
                    detail.Function == "IFNOTEQUAL" ||
                    detail.Function == "IFNOTNULLEMPTYWHITESPACE" ||
                    detail.Function == "JOIN" ||
                    detail.Function == "CONCAT" ||
                    detail.Function == "AND" ||
                    detail.Function == "OR" ||
                    detail.Function == "CONTAINS" ||
                    detail.Function == "DIVISION")
                {
                    var obj = JsonPathHelper.GetJsonPathValue(policyJson, detail);
                    if (obj != null)
                    {
                        val = obj.ToString();
                    }
                }
                else if (detail.Function == "JOINARRAY")
                {
                    string jsonPath = string.Empty;
                    string constant = string.Empty;

                    foreach (var param in detail.Params)
                    {
                        if (!string.IsNullOrWhiteSpace(param.JsonPath))
                        {
                            jsonPath = param.JsonPath;
                        }

                        if (!string.IsNullOrWhiteSpace(param.Constant))
                        {
                            constant = param.Constant;
                        }
                    }

                    var tokens = JsonPathHelper.GetJsonPathArrayToken(policyJson, jsonPath);
                    if (tokens != null)
                    {
                        val = string.Join(constant, tokens.Values());
                    }
                }

                if (string.IsNullOrWhiteSpace(val) || val == "0")
                {
                    val = string.IsNullOrWhiteSpace(detail.Constant) ? val : detail.Constant;
                }
            }
            else if (detail != null)
            {
                Dictionary<string, object> paramsOfCustomFunction = new Dictionary<string, object>();
                if (!string.IsNullOrWhiteSpace(detail.Format))
                {
                    paramsOfCustomFunction.Add("Format", detail.Format);
                }

                if (!string.IsNullOrWhiteSpace(detail.Constant))
                {
                    paramsOfCustomFunction.Add("Constant", detail.Constant);
                }

                if (detail.Params != null)
                {
                    foreach (MergedField mergedField in detail.Params)
                    {
                        if (!string.IsNullOrWhiteSpace(mergedField.Field))
                        {
                            paramsOfCustomFunction.Add("Field", mergedField.Field);
                        }
                    }
                }

                object[] parameters = new object[] { policy, docManager, paramsOfCustomFunction };
                object value = ReflectionHelper.InvokeFunction(detail.Function, parameters);
                if (value != null)
                {
                    return value.ToString();
                }
            }

            return val;
        }
    }
}
