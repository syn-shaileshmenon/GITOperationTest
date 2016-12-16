// <copyright file="JsonPathHelper.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Helpers
{
    using System.Collections.Generic;
    using System.Linq;
    using Entities;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Class JsonPathHelper - For evaluating JsonPath
    /// </summary>
    public static class JsonPathHelper
    {
        /// <summary>
        /// Returns the value for the specified mapping based on JsonPath.
        /// </summary>
        /// <param name="policyJson">policy</param>
        /// <param name="mappingDetail">mapping</param>
        /// <returns>value for the mapping</returns>
        public static object GetJsonPathValue(JObject policyJson, MappingDetail mappingDetail)
        {
            var parameters = new List<object>();
            if (mappingDetail.Params != null && mappingDetail.Params.Any())
            {
                foreach (var param in mappingDetail.Params)
                {
                    if (!string.IsNullOrWhiteSpace(param.JsonPath))
                    {
                        var value = GetJsonPathValue(policyJson, param);
                        parameters.Add(value);
                    }
                    else if (!param.HasFunction && (param.Function == "IFTRUE" ||
                                                    param.Function == "IFNULL" ||
                                                    param.Function == "IFNOTNULL" ||
                                                    param.Function == "IFNOTEQUAL" ||
                                                    param.Function == "STRINGFORMAT" ||
                                                    param.Function == "IFNOTNULLEMPTYWHITESPACE" ||
                                                    param.Function == "JOIN" ||
                                                    param.Function == "SUM" ||
                                                    param.Function == "CONCAT" ||
                                                    param.Function == "OR" ||
                                                    param.Function == "AND" || param.Function == "CONTAINS" || param.Function == "DIVISION"))
                    {
                        var value = GetJsonPathValue(policyJson, param);
                        parameters.Add(value);
                    }
                    else if (!param.HasFunction && param.Function == "CONDITION" && param.Params != null && param.Params.Count > 0)
                    {
                        var conditions = new List<object>();
                        foreach (var condition in param.Params)
                        {
                            var conditionValue = GetJsonPathValue(policyJson, condition);
                            conditions.Add(conditionValue);
                        }

                        parameters.Add(conditions);
                    }
                    else if (param.Constant != null)
                    {
                        parameters.Add(param.Constant);
                    }
                }
            }

            var jsonPath = mappingDetail.JsonPath;
            if (!string.IsNullOrWhiteSpace(jsonPath))
            {
                if (parameters.Any())
                {
                    jsonPath = string.Format(jsonPath, parameters.ToArray());
                }
            }
            else if (mappingDetail.Function == "SUM")
            {
                var sumValue = parameters.Select(p => p.ConvertDecimal()).Sum().ToString();
                return FormatTypedValue(sumValue, mappingDetail.Format);
            }
            else if (mappingDetail.Function == "STRINGFORMAT")
            {
                var formatString = parameters[0].ToString();
                var formatArray = parameters.GetRange(1, parameters.Count - 1).ToArray();
                var formatValue = string.Format(formatString, formatArray);

                return formatValue;
            }
            else if (mappingDetail.Function == "CONDITION")
            {
                return string.Empty;
            }
            else if (mappingDetail.Function == "AND")
            {
                if (parameters.Count >= 4)
                {
                    bool left = false;
                    bool right = false;

                    bool.TryParse(parameters[0].ToString(), out left);
                    bool.TryParse(parameters[1].ToString(), out right);
                    return left && right ? parameters[2] : parameters[3];
                }
                else if (parameters.Count == 3 && parameters[0] != null)
                {
                    var conditions = parameters[0] as List<object>;
                    if (conditions != null)
                    {
                        foreach (var item in conditions)
                        {
                            bool itemBool = false;
                            bool.TryParse(item.ToString(), out itemBool);
                            if (!itemBool)
                            {
                                return parameters[2];
                            }
                        }

                        return parameters[1];
                    }

                    return parameters[2];
                }
            }
            else if (mappingDetail.Function == "OR")
            {
                if (parameters.Count >= 4)
                {
                    bool left = false;
                    bool right = false;

                    bool.TryParse(parameters[0].ToString(), out left);
                    bool.TryParse(parameters[1].ToString(), out right);
                    return left || right ? parameters[2] : parameters[3];
                }
                else if (parameters.Count == 3 && parameters[0] != null)
                {
                    var conditions = parameters[0] as List<object>;
                    if (conditions != null)
                    {
                        foreach (var item in conditions)
                        {
                            bool itemBool = false;
                            bool.TryParse(item.ToString(), out itemBool);
                            if (itemBool)
                            {
                                return parameters[1];
                            }
                        }
                    }

                    return parameters[2];
                }
            }
            else if (mappingDetail.Function == "IFNOTEQUAL")
            {
                if (!parameters[0].Equals(parameters[1]))
                {
                    return parameters[2];
                }
                else
                {
                    return parameters[3];
                }
            }
            else if (mappingDetail.Function == "IFNULL")
            {
                if (parameters[0] == null)
                {
                    return parameters[1];
                }
                else
                {
                    return parameters[2];
                }
            }
            else if (mappingDetail.Function == "IFNOTNULL")
            {
                if (parameters[0] != null)
                {
                    return parameters[1];
                }
                else
                {
                    return parameters[2];
                }
            }
            else if (mappingDetail.Function == "IFTRUE")
            {
                var checkPassed = false;

                if (parameters[0] != null)
                {
                    var boolString = parameters[0].ToString().ToLowerInvariant();
                    bool.TryParse(boolString, out checkPassed);
                }

                if (checkPassed)
                {
                    return parameters[1];
                }
                else
                {
                    return parameters[2];
                }
            }
            else if (mappingDetail.Function == "IFNOTNULLEMPTYWHITESPACE")
            {
                if (parameters[0] != null && !string.IsNullOrWhiteSpace(parameters[0].ToString()))
                {
                    return parameters[1];
                }
                else
                {
                    return parameters[2];
                }
            }
            else if (mappingDetail.Function == "JOIN")
            {
                object[] joinParameters = new object[parameters.Count - 1];
                parameters.CopyTo(1, joinParameters, 0, parameters.Count - 1);
                var nonEmptyItems = parameters.Where(a => a.ToString() != string.Empty).Skip(1).Select(a => a).ToArray();
                return string.Join(parameters[0].ToString(), nonEmptyItems.Cast<string>());
            }
            else if (mappingDetail.Function == "CONCAT")
            {
                return string.Concat(parameters);
            }
            else if (mappingDetail.Function == "CONTAINS")
            {
                if (parameters[0] != null && JArray.Parse(parameters[0].ToString()).ToList().Contains(parameters[1].ToString()))
                {
                    return parameters[2].ToString();
                }
                else
                {
                    return parameters[3].ToString();
                }
            }
            else if (mappingDetail.Function == "DIVISION")
            {
                decimal divideValue = 0;
                if (parameters[0] != null && parameters[1] != null)
                {
                    if (parameters[1].ConvertDecimal() != 0)
                    {
                        divideValue = parameters[0].ConvertDecimal() / parameters[1].ConvertDecimal();
                    }
                }

                return FormatTypedValue(divideValue, mappingDetail.Format);
            }
            else if (mappingDetail.Constant != null)
            {
                return FormatTypedValue(mappingDetail.Constant, mappingDetail.Format);
            }

            return GetJsonPathValue(policyJson, jsonPath, mappingDetail.Format);
        }

        /// <summary>
        /// Gets the JToken list from policy, based on specified JsonPath
        /// </summary>
        /// <param name="policyJson">Policy</param>
        /// <param name="jsonPath">JsonPath</param>
        /// <returns>JToken list after evaluating JsonPath</returns>
        public static IEnumerable<JToken> GetJsonPathArrayToken(JObject policyJson, string jsonPath)
        {
            var tokens = policyJson.SelectTokens(jsonPath);

            return tokens;
        }

        /// <summary>
        /// Converts the decimal.
        /// </summary>
        /// <param name="s">The s.</param>
        /// <returns>System.Decimal.</returns>
        private static decimal ConvertDecimal(this object s)
        {
            decimal output = 0;
            decimal.TryParse(s.ToString(), out output);

            return output;
        }

        /// <summary>
        /// Gets the value from policy, based on specified JsonPath
        /// </summary>
        /// <param name="policyJson">Policy</param>
        /// <param name="jsonPath">JsonPath</param>
        /// <param name="format">Format</param>
        /// <returns>Value after evaluating JsonPath</returns>
        private static object GetJsonPathValue(JObject policyJson, string jsonPath, string format)
        {
            JToken token = policyJson.SelectToken(jsonPath);

            object returnValue = null;
            if (token != null && token.Type != JTokenType.Null)
            {
                returnValue = token;
            }

            ////if (returnValue == null)
            ////{
            ////    returnValue = string.Empty;
            ////}

            return FormatTypedValue(returnValue, format);
        }

        /// <summary>
        /// Formats the typed value.
        /// </summary>
        /// <param name="returnValue">The return value.</param>
        /// <param name="format">The format.</param>
        /// <returns>System.String.</returns>
        private static object FormatTypedValue(object returnValue, string format)
        {
            if (returnValue == null)
            {
                return null;
            }

            var typedValue = GetTypedValue(returnValue);

            if (!string.IsNullOrWhiteSpace(format))
            {
                return FormatHelper.FormatValue(typedValue, format);
            }
            else
            {
                return returnValue;
            }
        }

        /// <summary>
        /// Gets the typed value
        /// </summary>
        /// <param name="value">string value</param>
        /// <returns>Typed value</returns>
        private static object GetTypedValue(object value)
        {
            object result = value;

            decimal decimalValue = 0;
            if (decimal.TryParse(value.ToString(), out decimalValue))
            {
                result = decimalValue;
            }

            return result;
        }
    }
}
