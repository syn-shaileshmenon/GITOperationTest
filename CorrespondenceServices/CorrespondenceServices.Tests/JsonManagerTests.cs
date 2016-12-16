// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-15-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-16-2016
// ***********************************************************************
// <copyright file="JsonManagerTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Tests
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Class JsonManagerTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class JsonManagerTests
    {
        /// <summary>
        /// Tests the generate mapping dictionary can handle empty parameters.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleEmptyParams()
        {
            var templateString = "{\"TestKey\": {\"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(0, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle parameters format.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleParamsFormat()
        {
            var templateString = "{\"TestKey\": {\"params\":[{\"format\":\"{0:0n}\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(1, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle parameters json path.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleParamsJsonPath()
        {
            var templateString = "{\"TestKey\": {\"params\":[{\"jsonPath\":\"$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='ET')].TotalInsuredValue\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(1, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle parameters custom function.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleParamsCustomFunction()
        {
            var templateString = "{\"TestKey\": {\"params\":[{\"function\":\"CalculateTaxesFees\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(1, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle parameters standard function sum.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleParamsStandardFunctionSum()
        {
            var templateString = "{\"TestKey\": {\"params\":[{\"function\":\"SUM\", \"params\":[{\"constant\":0}]}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(1, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle constant.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleConstant()
        {
            var templateString = "{\"TestKey\": {\"constant\":\"MyValue\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Constant);
            Assert.AreEqual("MyValue", dictionary.ElementAt(0).Value.Constant);
            Assert.IsNull(dictionary.ElementAt(0).Value.Params);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle format.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleFormat()
        {
            var templateString = "{\"TestKey\": {\"format\":\"{0:0n}\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNull(dictionary.ElementAt(0).Value.Params);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Format);
            Assert.AreEqual("{0:0n}", dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle json path.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleJsonPath()
        {
            var templateString = "{\"TestKey\": {\"jsonPath\":\"$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='ET')].TotalInsuredValue\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNull(dictionary.ElementAt(0).Value.Params);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.AreEqual("$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='ET')].TotalInsuredValue", dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNull(dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle custom function.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleCustomFunction()
        {
            var templateString = "{\"TestKey\": {\"function\":\"CalculateTaxesFees\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNull(dictionary.ElementAt(0).Value.Params);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("CalculateTaxesFees", dictionary.ElementAt(0).Value.Function);
            Assert.IsTrue(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function sum.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionSum()
        {
            var templateString = "{\"TestKey\": {\"function\":\"SUM\", \"params\":[{\"constant\":0}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(1, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("SUM", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function if true.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionIfTrue()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFTRUE\", \"params\":[{\"constant\":\"a\"},{\"constant\":\"b\"},{\"constant\":\"c\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(3, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("IFTRUE", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function if not null.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionIfNotNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFNOTNULL\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNull(dictionary.ElementAt(0).Value.Params);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("IFNOTNULL", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function if not equal.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionIfNotEqual()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFNOTEQUAL\", \"params\":[{\"constant\":\"a\"},{\"constant\":\"b\"},{\"constant\":\"c\"},{\"constant\":\"d\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(4, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("IFNOTEQUAL", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function if not null empty whitespace.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionIfNotNullEmptyWhitespace()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFNOTNULLEMPTYWHITESPACE\", \"params\":[{\"constant\":\"a\"},{\"constant\":\"b\"},{\"constant\":\"c\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(3, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("IFNOTNULLEMPTYWHITESPACE", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function join.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionJoin()
        {
            var templateString = "{\"TestKey\": {\"function\":\"JOIN\", \"params\":[{\"constant\":\"a\"},{\"constant\":\"b\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(2, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("JOIN", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function CONCAT.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionConcat()
        {
            var templateString = "{\"TestKey\": {\"function\":\"CONCAT\", \"params\":[{\"constant\":\"a\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(1, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("CONCAT", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function and.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionAnd4Parameters()
        {
            var templateString = "{\"TestKey\": {\"function\":\"AND\", \"params\":[{\"constant\":\"a\"}, {\"constant\":\"b\"}, {\"constant\":\"true\"}, {\"constant\":\"not true\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(4, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("AND", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function and.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionAnd3Parameters()
        {
            var templateString = "{\"TestKey\": {\"function\":\"AND\", \"params\":[{\"function\":\"CONDITION\", \"params\":[{\"constant\":\"a\"}, {\"constant\":\"b\"}]}, {\"constant\":\"true\"}, {\"constant\":\"not true\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(3, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("AND", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function or.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionOr()
        {
            var templateString = "{\"TestKey\": {\"function\":\"OR\", \"params\":[{\"constant\":\"a\"}, {\"constant\":\"b\"}, {\"constant\":\"true\"}, {\"constant\":\"not true\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(4, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("OR", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function string format.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionStringFormat()
        {
            var templateString = "{\"TestKey\": {\"function\":\"STRINGFORMAT\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNull(dictionary.ElementAt(0).Value.Params);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("STRINGFORMAT", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function CONCAT parameters null.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionConcatParamsNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"CONCAT\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'At least one parameters must be provided to CONCAT function.'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function if true parameters null.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionIfTrueParamsNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFTRUE\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to IFTRUE function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function if not equal parameters null.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionIfNotEqualParamsNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFNOTEQUAL\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Four parameters must be provided to IFNOTEQUAL function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function if not null empty whitespace parameters null.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionIfNotNullEmptyWhitespaceParamsNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFNOTNULLEMPTYWHITESPACE\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to IFNOTNULLEMPTYWHITESPACE function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function join with parameters null.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionJoinPararmsNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"JOIN\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'At least two parameters must be provided to JOIN function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function CONCAT parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionConcatParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"CONCAT\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'At least one parameters must be provided to CONCAT function.'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function if true parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionIfTrueParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFTRUE\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to IFTRUE function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function if not equal parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionIfNotEqualParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFNOTEQUAL\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Four parameters must be provided to IFNOTEQUAL function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function if not null empty whitespace parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionIfNotNullEmptyWhitespaceParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"IFNOTNULLEMPTYWHITESPACE\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to IFNOTNULLEMPTYWHITESPACE function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function join parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionJoinParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"JOIN\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'At least two parameters must be provided to JOIN function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function and parameters no condition.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionAnd3ParamsNoCondition()
        {
            var templateString = "{\"TestKey\": {\"function\":\"AND\", \"params\":[{\"constant\":\"a\"}, {\"constant\":\"true\"}, {\"constant\":\"not true\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three (with conditions) or Four parameters must be provided to AND function.  3 provided. =>...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function and parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionAndParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"AND\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three (with conditions) or Four parameters must be provided to AND function.  0 provided. =>...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function or3 parameters no condition.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionOr3ParamsNoCondition()
        {
            var templateString = "{\"TestKey\": {\"function\":\"OR\", \"params\":[{\"constant\":\"a\"}, {\"constant\":\"true\"}, {\"constant\":\"not true\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three (with conditions) or Four parameters must be provided to AND function.  3 provided. =>...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function or parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionOrParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"OR\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three (with conditions) or Four parameters must be provided to AND function.  0 provided. =>...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary can handle standard function contains.
        /// </summary>
        [TestMethod]
        public void TestGenerateMappingDictionaryCanHandleStandardFunctionCONTAINS()
        {
            var templateString = "{\"TestKey\": {\"function\":\"CONTAINS\", \"params\":[{\"jsonPath\":\"$.ImLine.RiskUnits[?(@.ClassCode == '445')].RollUp.SelectedClauses\"},{\"constant\":\"a\"},{\"constant\":\"a\"},{\"constant\":\"b\"}]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.IsNotNull(dictionary);
            Assert.AreEqual(1, dictionary.Keys.Count);
            Assert.AreEqual("testkey", dictionary.ElementAt(0).Key);
            Assert.IsNotNull(dictionary.ElementAt(0).Value);
            Assert.IsNull(dictionary.ElementAt(0).Value.Constant);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Params);
            Assert.AreEqual(4, dictionary.ElementAt(0).Value.Params.Count);
            Assert.IsNull(dictionary.ElementAt(0).Value.Field);
            Assert.IsNull(dictionary.ElementAt(0).Value.Format);
            Assert.IsNull(dictionary.ElementAt(0).Value.Image);
            Assert.IsNull(dictionary.ElementAt(0).Value.JsonPath);
            Assert.IsNotNull(dictionary.ElementAt(0).Value.Function);
            Assert.AreEqual("CONTAINS", dictionary.ElementAt(0).Value.Function);
            Assert.IsFalse(dictionary.ElementAt(0).Value.HasFunction);
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function contains parameters null.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionCONTAINSParamsNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"CONTAINS\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to CONTAINS function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function contains parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionCONTAINSParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"CONTAINS\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to CONTAINS function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function join array parameters null.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionJOINARRAYParamsNull()
        {
            var templateString = "{\"TestKey\": {\"function\":\"JOINARRAY\"}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to JOINARRAY function.  0 provided...'");
        }

        /// <summary>
        /// Tests the generate mapping dictionary fails standard function join array parameters empty.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestGenerateMappingDictionaryFailsStandardFunctionJOINARRAYParamsEmpty()
        {
            var templateString = "{\"TestKey\": {\"function\":\"JOINARRAY\", \"params\":[]}}";
            var tokenizedString = JsonManagerTests.GetJTokenizableString(templateString);
            var template = JToken.Parse(tokenizedString);
            var manager = new JsonManager();
            var dictionary = manager.GenerateMappingDictionary(template);

            Assert.Fail("We should fail before this line with error 'Three parameters must be provided to JOINARRAY function.  0 provided...'");
        }

        /// <summary>
        /// Gets the json tokenized string.
        /// </summary>
        /// <param name="json">The json.</param>
        /// <returns>System.String.</returns>
        internal static string GetJTokenizableString(string json)
        {
            var serialized = JsonConvert.SerializeObject(json);
            return serialized;
        }
    }
}
