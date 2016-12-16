// ***********************************************************************
// Assembly         : LossRunServices.Tests
// Author           : rsteelea
// Created          : 06-21-2016
//
// Last Modified By : rsteelea
// Last Modified On : 06-24-2016
// ***********************************************************************
// <copyright file="RestHelperStub.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CorrespondenceServices.Tests.Stubs
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.RestHelper.Enumerations;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using RestSharp;

    /// <summary>
    /// Class RestHelperStub.
    /// </summary>
    /// <seealso cref="Mkl.WebTeam.RestHelper.Interfaces.IRestHelper" />
    /// <seealso cref="IRestHelper" />
    [ExcludeFromCodeCoverage]
    public class RestHelperStub : IRestHelper
    {
        /// <summary>
        /// Gets or sets the custom mapping templates.
        /// </summary>
        /// <value>The custom mapping templates.</value>
        public static Dictionary<string, string> CustomMappingTemplates { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Gets or sets the document template bytes.
        /// </summary>
        /// <value>The document template bytes.</value>
        public static Dictionary<string, string> DocumentTemplateBase64 { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Get service result
        /// </summary>
        /// <param name="requestUri">requestUri as string</param>
        /// <param name="requestMethod">as RestSharp.Method</param>
        /// <param name="user">as User</param>
        /// <param name="requestBody">as object</param>
        /// <param name="endpoint">ServiceEndpoint</param>
        /// <param name="correlationId">correlationId as string</param>
        /// <returns>JToken</returns>
        /// <exception cref="ArgumentNullException">If request uri is not provided</exception>
        /// <exception cref="NotImplementedException">Not developed yet</exception>
        public JToken GetServiceResult(
            string requestUri,
            Method requestMethod,
            IUser user,
            object requestBody = null,
            WebTeamServiceEndpoint endpoint = WebTeamServiceEndpoint.BaseServices,
            string correlationId = null)
        {
            if (string.IsNullOrWhiteSpace(requestUri))
            {
                throw new ArgumentNullException(requestUri);
            }

            var request = requestUri.ToLowerInvariant();
            var dictArguments = GetRequestArguments(request);

            if (request.Contains("carrier?"))
            {
                var carriers = RestHelperStub.GetFakeCarrierList();
                var json = JsonConvert.SerializeObject(carriers);
                return JToken.Parse(json);
            }
            else if (request.Contains("defaultmappingtemplates"))
            {
                if (dictArguments.ContainsKey("effectivedate"))
                {
                    var defaultTemplate = RestHelperStub.GetDefaultMappingTemplate();
                    return JToken.Parse(defaultTemplate);
                }
            }
            else if (request.Contains("custommappingtemplates"))
            {
                if (dictArguments.ContainsKey("templateid") &&
                    dictArguments.ContainsKey("effectivedate"))
                {
                    if (RestHelperStub.CustomMappingTemplates.ContainsKey(dictArguments["templateid"]))
                    {
                        var customTemplate = RestHelperStub.CustomMappingTemplates[dictArguments["templateid"]];
                        if (customTemplate == null)
                        {
                            return JToken.Parse("{}");
                        }

                        return JToken.Parse(customTemplate);
                    }
                }

                return JToken.Parse("{}");
            }
            else if (request.Contains("document/"))
            {
                if (dictArguments.ContainsKey("documenttype"))
                {
                    // Value must be "=WordTemplate"
                    // Must extract the detailId out of the request
                    var detailId = RestHelperStub.GetRequestId(request);
                    var documentData = RestHelperStub.DocumentTemplateBase64[detailId];
                    var mappingData = $@"{{""WordTemplateBinaryData"":""{documentData}""}}";

                    return JToken.Parse(mappingData);
                }
            }
            else if (request.Contains("getpresidentsignature"))
            {
                var filename = @"""Files\\PreSig.jpg""";
                return JToken.Parse(filename);
            }
            else if (request.Contains("getsecretarysignature"))
            {
                var filename = @"""Files\\SecSig.jpg""";
                return JToken.Parse(filename);
            }
            else if (request.Contains("document"))
            {
                // Must return DocumentTemplateMappingId with the value used in the previous if block
                if (dictArguments.ContainsKey("headerid") && dictArguments.ContainsKey("asofdate"))
                {
                    ////headerid = new Guid(documentId)
                    ////asOfDate = effectiveDate.Value
                    var headerId = dictArguments["headerid"];
                    var mappingId = $@"{{""DocumentTemplateMappingId"":""{headerId}"", ""DetailId"":""{headerId}""}}";
                    return JToken.Parse(mappingId);
                }
            }

            throw new NotImplementedException();
        }

        /// <summary>
        /// Get service result
        /// </summary>
        /// <param name="requestUri">requestUri as string</param>
        /// <param name="requestMethod">as RestSharp.Method</param>
        /// <param name="objName">as string</param>
        /// <param name="user">as User</param>
        /// <param name="requestBody">as object</param>
        /// <param name="endpoint">ServiceEndpoint</param>
        /// <param name="correlationId">correlationId as string</param>
        /// <returns>JEnumerable of JToken</returns>
        /// <exception cref="ArgumentNullException">If request uri is not provided</exception>
        /// <exception cref="NotImplementedException">Not developed yet</exception>
        public JEnumerable<JToken> GetServiceResult(
            string requestUri,
            Method requestMethod,
            string objName,
            IUser user,
            object requestBody = null,
            WebTeamServiceEndpoint endpoint = WebTeamServiceEndpoint.BaseServices,
            string correlationId = null)
        {
            if (string.IsNullOrWhiteSpace(requestUri))
            {
                throw new ArgumentNullException(requestUri);
            }

            var request = requestUri.ToLowerInvariant();
            var dictArguments = GetRequestArguments(request);

            if (request.Contains("carrier?"))
            {
                var carriers = RestHelperStub.GetFakeCarrierList();
                var json = JsonConvert.SerializeObject(carriers);
                JToken jobject = JToken.Parse(json);

                if (objName != null &&
                    jobject[objName] != null)
                {
                    return jobject[objName].Children();
                }

                return default(JEnumerable<JToken>);
            }

            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the fake carrier list.
        /// </summary>
        /// <returns>List&lt;Carrier&gt;.</returns>
        internal static List<Carrier> GetFakeCarrierList()
        {
            return new List<Carrier>(new[]
                {
                    new Carrier()
                    {
                        Name = "Essex Insurance Company",
                        ShortName = "Essex",
                        Code = "ESS",
                        IsMarkel = true,
                        AddressLine1 = "Test Address Line 1",
                        AddressLine2 = "Test Address Line 2",
                        City = "Wilmington",
                        State = "DE",
                        ZipCode = "19801",
                    },

                    new Carrier()
                    {
                        Name = "Evanston Insurance Company",
                        ShortName = "Evanston",
                        Code = "EVA",
                        IsMarkel = true,
                        AddressLine1 = "Test Address Line 1",
                        AddressLine2 = string.Empty,
                        City = "Deerfield",
                        State = "IL",
                        ZipCode = "60015",
                    }
                });
        }

        /// <summary>
        /// Gets the request arguments.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <returns>Dictionary&lt;System.String, System.String&gt;.</returns>
        private static Dictionary<string, string> GetRequestArguments(string request)
        {
            var queryStringPos = request.IndexOf("?");
            var queryString = string.Empty;
            if (queryStringPos > 0 &&
                queryStringPos < request.Length)
            {
                queryString = request.Substring(queryStringPos + 1);
            }

            var arguments = queryString.Split(new char[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
            var dictArguments = new Dictionary<string, string>();
            foreach (var arg in arguments)
            {
                var splits = arg.Split('=');
                if (splits.Count() >= 2)
                {
                    dictArguments.Add(splits[0], splits[1]);
                }
                else if (splits.Count() == 1)
                {
                    dictArguments.Add(splits[0], string.Empty);
                }
            }

            return dictArguments;
        }

        /// <summary>
        /// Gets the request identifier.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <returns>System.String.</returns>
        private static string GetRequestId(string request)
        {
            var lastPos = request.LastIndexOf("/");
            if (lastPos > 0 && lastPos < request.Length - 1)
            {
                var questionIndex = request.IndexOf("?", lastPos);
                if (questionIndex > 0)
                {
                    return request.Substring(lastPos + 1, questionIndex - lastPos - 1);
                }
                else
                {
                    return request.Substring(lastPos + 1);
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Gets the default mapping template.
        /// </summary>
        /// <returns>System.String.</returns>
        private static string GetDefaultMappingTemplate()
        {
            var defaultTemplate = "\"{\\\"CurrentDate\\\":{\\\"function\\\"  : \\\"GetTodayDate\\\" ,\\\"format\\\": \\\"{0:MM/dd/yyyy}\\\"},\\\"ISSUINGCOMPANYP\\\" : { \\\"field\\\" : \\\"Policy.CarrierName\\\" },\\\"PolicyNumberP\\\" : { \\\"field\\\" :\\\"Policy.PolicyNumber\\\" },\\\"RenewalOfNumber\\\" : {\\\"field\\\" : \\\"Policy.ParentPolicy.PolicyNumber\\\", \\\"constant\\\":\\\"NEW\\\" },\\\"InsuredName\\\" : { \\\"field\\\" : \\\"Policy.PrimaryInsured.Name\\\" },\\\"InsuredName2P\\\" : { \\\"field\\\" : \\\"Policy.SecondaryInsured.Name\\\" },\\\"InsuredAddress1P\\\" : { \\\"field\\\" : \\\"Policy.SecondaryInsured.StreetAddress.Line1\\\" },\\\"InsuredAddress2P\\\" : { \\\"field\\\" : \\\"Policy.SecondaryInsured.StreetAddress.Line2\\\" },\\\"InsuredCityP\\\" : { \\\"field\\\" : \\\"Policy.SecondaryInsured.StreetAddress.City\\\" },\\\"InsuredStateP\\\" : { \\\"field\\\" : \\\"Policy.SecondaryInsured.StreetAddress.StateCode\\\" },\\\"InsuredZIP\\\" : { \\\"field\\\" : \\\"Policy.SecondaryInsured.StreetAddress.ZipCode\\\" },\\\"EffectiveDateP\\\" : { \\\"field\\\" : \\\"Policy.EffectiveDate\\\" , \\\"format\\\": \\\"{0:MM/dd/yyyy}\\\" },\\\"ExpirationDateP\\\" : { \\\"field\\\" : \\\"Policy.ExpirationDate\\\", \\\"format\\\": \\\"{0:MM/dd/yyyy}\\\" },\\\"EXEXOccLimit\\\" : { \\\"field\\\" : \\\"Policy.XsLine.SelectedLimit\\\", \\\"format\\\": \\\"{0:n0}\\\" },\\\"EXEXAggLimit\\\" : { \\\"field\\\" : \\\"Policy.XsLine.SelectedLimit\\\", \\\"format\\\": \\\"{0:n0}\\\"},\\\"PremiumOnly\\\" : { \\\"field\\\" : \\\"Policy.XsLine.RollUp.Premium\\\", \\\"format\\\": \\\"{0:n0}\\\" },\\\"TerrorismPremiums\\\" : { \\\"function\\\" : \\\"GetTerrorismPremium\\\", \\\"format\\\": \\\"{0:n0}\\\",\\\"constant\\\":\\\"Rejected\\\"},\\\"TotalPremiums\\\" : { \\\"field\\\" : \\\"Policy.TotalAmountWithTaxes\\\", \\\"format\\\": \\\"{0:n2}\\\" },\\\"AgentCode\\\" : { \\\"field\\\" : \\\"Policy.Agency.Code\\\" },\\\"AgentName1\\\" : { \\\"field\\\" : \\\"Policy.Agency.Name\\\" },\\\"AgentAddress1\\\" : { \\\"field\\\" : \\\"Policy.Agency.StreetAddress.Line1\\\" },\\\"AgentAddress2\\\" : { \\\"field\\\" : \\\"Policy.Agency.StreetAddress.Line2\\\" },\\\"AgentCity\\\" : { \\\"field\\\" : \\\"Policy.Agency.StreetAddress.City\\\" },\\\"AgentState\\\" : { \\\"field\\\" : \\\"Policy.Agency.StreetAddress.StateCode\\\" },\\\"AgentZipCode\\\" : { \\\"field\\\" : \\\"Policy.Agency.StreetAddress.ZipCode\\\" },\\\"ISSUINGCOMPANYP_1\\\" : { \\\"field\\\" : \\\"Policy.CarrierName\\\" }, \\\"ISSUINGCOMPANYP_2\\\" : { \\\"field\\\" : \\\"Policy.CarrierName\\\" }, \\\"ISSUINGCOMPANYP_3\\\" : { \\\"field\\\" : \\\"Policy.CarrierName\\\" }, \\\"ISSUINGCOMPANYP_4\\\" : { \\\"field\\\" : \\\"Policy.CarrierName\\\" }, \\\"MEP\\\" : { \\\"field\\\" : \\\"Policy.MinimumEarnedPercentage.ApplicableValue\\\" }}\"";

            return defaultTemplate;
        }
    }
}
