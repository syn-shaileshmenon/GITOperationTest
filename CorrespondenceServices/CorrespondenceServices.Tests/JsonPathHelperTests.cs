// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-12-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-16-2016
// ***********************************************************************
// <copyright file="JsonPathHelperTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Tests
{
    using System.Diagnostics.CodeAnalysis;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.DocumentGenerator.Helpers;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Class JsonPathHelperTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class JsonPathHelperTests
    {
        /// <summary>
        /// Gets or sets the policy json.
        /// </summary>
        /// <value>The policy json.</value>
        private JObject PolicyJson { get; set; }

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        [TestInitialize]
        public void Initialize()
        {
            this.SetPolicyJson();
        }

        /// <summary>
        /// Determines whether this instance [can get first level value].
        /// </summary>
        [TestMethod]
        public void CanGetFirstLevelValue()
        {
            var mapping = new MappingDetail()
            {
                JsonPath = "$.Item1.Field1"
            };

            var result = JsonPathHelper.GetJsonPathValue(this.PolicyJson, mapping);
            Assert.IsNotNull(result);
            Assert.AreNotEqual(string.Empty, result);
            Assert.AreEqual("Value 1.1", result.ToString());
        }

        /// <summary>
        /// Determines whether this instance [can get deep level value].
        /// </summary>
        [TestMethod]
        public void CanGetDeepLevelValue()
        {
            var mapping = new MappingDetail()
            {
                JsonPath = "$.Item1.Field4.ItemA"
            };

            var result = JsonPathHelper.GetJsonPathValue(this.PolicyJson, mapping);
            Assert.IsNotNull(result);
            Assert.AreNotEqual(string.Empty, result);
            Assert.AreEqual("Value 1.4.A", result.ToString());
        }

        /// <summary>
        /// Determines whether this instance [can get array value].
        /// </summary>
        [TestMethod]
        public void CanGetArrayValue()
        {
            var mapping = new MappingDetail()
            {
                JsonPath = "$.Item4[?(@.Field2 == 30)].Field1"
            };

            var result = JsonPathHelper.GetJsonPathValue(this.PolicyJson, mapping);
            Assert.IsNotNull(result);
            Assert.AreNotEqual(string.Empty, result.ToString());
            Assert.AreEqual("Value Array 1", result.ToString());
        }

        /// <summary>
        /// Invalids the json path throws exception.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(JsonException))]
        public void InvalidJsonPathThrowsException()
        {
            var mapping = new MappingDetail()
            {
                JsonPath = "SomeJson.String[that is invaid]"
            };

            var result = JsonPathHelper.GetJsonPathValue(this.PolicyJson, mapping);
            Assert.Fail("We should fail before this line with error 'Unexpected character '.' found trying to parse Json'");
        }

        /// <summary>
        /// Tests when a json path is not found that it returns null
        /// </summary>
        [TestMethod]
        public void JsonPathElementNotFoundReturnsNull()
        {
            var mapping = new MappingDetail()
            {
                JsonPath = "$.Item33[?(@.Field2 == 30)].Field1"
            };

            var result = JsonPathHelper.GetJsonPathValue(this.PolicyJson, mapping);
            Assert.IsNull(result);
            ////Assert.AreEqual(string.Empty, result);
        }

        /// <summary>
        /// Tests that when a json path sub-query is not found that it returns Null
        /// </summary>
        [TestMethod]
        public void JsonPathArrayQueryNotFoundReturnsNull()
        {
            var mapping = new MappingDetail()
            {
                JsonPath = "$.Item4[?(@.Field2 == '30')].Field1"
            };

            var result = JsonPathHelper.GetJsonPathValue(this.PolicyJson, mapping);
            Assert.IsNull(result);
            ////Assert.AreEqual(string.Empty, result);
        }

        /// <summary>
        /// Sets the policy json.
        /// </summary>
        private void SetPolicyJson()
        {
            dynamic json = new JObject();

            json.Item1 = new JObject();
            json.Item1.Field1 = "Value 1.1";
            json.Item1.Field2 = 10;
            json.Item1.Field3 = 10.0m;
            json.Item1.Field4 = new JObject();
            json.Item1.Field4.ItemA = "Value 1.4.A";
            json.Item2 = "Value 2";
            json.Item3 = new JObject();
            json.Item3.Field1 = "Value 3.1";
            json.Item3.Field3 = 30;
            json.Item3.Field5 = 30m;

            dynamic arrayItem1 = new JObject();
            arrayItem1.Field1 = "Value Array 1";
            arrayItem1.Field2 = 30;

            dynamic arrayItem2 = new JObject();
            arrayItem2.Field1 = "Value Array 2";
            arrayItem2.Field3 = "Value Array 2.3";
            arrayItem2.Field4 = 30;

            json.Item4 = new JArray("Value 4.1", "Value 4.2", arrayItem1, arrayItem2);

            var jsonString = json.ToString();
            this.PolicyJson = json;
        }
    }
}
