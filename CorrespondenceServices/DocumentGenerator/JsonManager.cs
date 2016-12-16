// <copyright file="JsonManager.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Json manager
    /// </summary>
    public class JsonManager : IJsonManager
    {
        /// <summary>
        /// GenerateMappingDictionary method generates a dictionary based on passed in JToken object
        /// </summary>
        /// <param name="jsonMappingTemplate">JToken</param>
        /// <returns>Dictionary</returns>
        public Dictionary<string, MappingDetail> GenerateMappingDictionary(JToken jsonMappingTemplate)
        {
            JObject jsonVal = JObject.Parse(jsonMappingTemplate.Value<JValue>().Value.ToString());
            Dictionary<string, MappingDetail> tempDict = new Dictionary<string, MappingDetail>();

            foreach (JToken child in jsonVal.Children())
            {
                var childProperty = child as JProperty;
                if (childProperty == null)
                {
                    continue;
                }

                MappingDetail detail = new MappingDetail();
                this.LoopChild(child, ref detail);
                tempDict.Add(childProperty.Name.ToLower(), detail);
            }

            return tempDict;
        }

        /// <summary>
        /// Create mapping detail using the field option
        /// </summary>
        /// <param name="mappingDetail">MappingDetail</param>
        /// <param name="propName">string</param>
        /// <param name="propValue">propValue</param>
        private static void CreateMappingDetail(ref MappingDetail mappingDetail, string propName, string propValue)
        {
            switch (propName)
            {
                case "field":
                    mappingDetail.Field = propValue;
                    break;

                case "image":
                    mappingDetail.Image = propValue;
                    break;

                case "function":
                    mappingDetail.Function = propValue;
                    mappingDetail.HasFunction = propValue != "SUM" &&
                                                propValue != "IFTRUE" &&
                                                propValue != "CONCAT" &&
                                                propValue != "STRINGFORMAT" &&
                                                propValue != "JOIN" &&
                                                propValue != "IFNOTNULLEMPTYWHITESPACE" &&
                                                propValue != "IFNULL" &&
                                                propValue != "IFNOTNULL" &&
                                                propValue != "IFNOTEQUAL" &&
                                                propValue != "AND" &&
                                                propValue != "OR" &&
                                                propValue != "CONDITION" &&
                                                propValue != "CONTAINS" &&
                                                propValue != "JOINARRAY" &&
                                                propValue != "DIVISION";
                    break;

                case "constant":
                    mappingDetail.Constant = propValue;
                    break;

                case "format":
                    mappingDetail.Format = propValue;
                    break;

                case "jsonPath":
                    mappingDetail.JsonPath = propValue;
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Loop the child elements of a node and devise a mapping detail object
        /// </summary>
        /// <param name="child">JToken</param>
        /// <param name="detail">MappingDetail</param>
        private void LoopChild(JToken child, ref MappingDetail detail)
        {
            foreach (JToken grandChild in child)
            {
                // The first time through this function for each item that has children will be an empty item
                // and will go one level further down recursively
                var property = grandChild as JProperty;

                if (property != null)
                {
                    if (property.Name == "params")
                    {
                        List<MappingDetail> paramDetails = this.LoopParamChild(grandChild);
                        detail.Params = paramDetails;
                    }
                    else
                    {
                        CreateMappingDetail(ref detail, property.Name, property.Value.ToString());
                    }
                }
                else
                {
                    this.LoopChild(grandChild, ref detail);
                }
            }

            // Perform some error logic around the parameter counts if a predefined function
            if (detail.HasFunction || string.IsNullOrWhiteSpace(detail.Function))
            {
                return;
            }

            switch (detail.Function)
            {
                case "OR":
                case "AND":
                    var error = false;
                    if (detail.Params == null || detail.Params.Count < 3)
                    {
                        error = true;
                    }
                    else
                    {
                        if (detail.Params.Count == 3)
                        {
                            if (!detail.Params[0].HasFunction &&
                                detail.Params[0].Function == "CONDITION")
                            {
                                if (detail.Params[0].Params == null || detail.Params[0].Params.Count < 2)
                                {
                                    error = true;
                                }
                            }
                            else
                            {
                                error = true;
                            }
                        }
                    }

                    if (error)
                    {
                        throw new ArgumentException($"Three (with conditions) or Four parameters must be provided to {detail.Function} function.  {detail.Params?.Count ?? 0} provided. => {child.ToString()}");
                    }

                    break;

                case "IFNOTEQUAL":
                case "CONTAINS":
                    if (detail.Params == null || detail.Params.Count != 4)
                    {
                        throw new ArgumentException($"Four parameters must be provided to {detail.Function} function.  {detail.Params?.Count ?? 0} provided. => {child.ToString()}");
                    }

                    break;

                case "IFTRUE":
                case "IFNULL":
                case "IFNOTNULLEMPTYWHITESPACE":
                    if (detail.Params == null || detail.Params.Count != 3)
                    {
                        throw new ArgumentException($"Three parameters must be provided to {detail?.Function} function.  {detail.Params?.Count ?? 0} provided. => {child.ToString()}");
                    }

                    break;

                case "DIVISION":
                    if (detail.Params == null || detail.Params.Count != 2)
                    {
                        throw new ArgumentException($"Two parameters must be provided to {detail?.Function} function.  {detail.Params?.Count ?? 0} provided. => {child.ToString()}");
                    }

                    break;

                case "JOIN":
                case "JOINARRAY":
                    if (detail.Params == null || detail.Params.Count < 2)
                    {
                        throw new ArgumentException($"At least two parameters must be provided to {detail?.Function} function.  {detail.Params?.Count ?? 0} provided. => {child.ToString()}");
                    }

                    break;

                case "CONCAT":
                case "SUM":
                    if (detail.Params == null || detail.Params.Count < 1)
                    {
                        throw new ArgumentException($"At least one parameter must be provided to {detail?.Function} function. => {child.ToString()}");
                    }

                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Loop through the parameter node retrieve the list of parameters required for a function
        /// </summary>
        /// <param name="child">JToken</param>
        /// <returns>List of MergedField</returns>
        private List<MappingDetail> LoopParamChild(JToken child)
        {
            List<MappingDetail> paramDetails = new List<MappingDetail>();
            var grandChildren = child.Values();

            foreach (JToken grandChild in grandChildren)
            {
                MappingDetail param = new MappingDetail();
                var property = grandChild as JProperty;

                if (property != null)
                {
                    CreateMappingDetail(ref param, property.Name, property.Value.ToString());
                }
                else
                {
                    this.LoopChild(grandChild, ref param);
                }

                paramDetails.Add(param);
            }

            return paramDetails;
        }
    }
}
