// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-17-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-17-2016
// ***********************************************************************
// <copyright file="TemplateProcessorTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using CorrespondenceServices.Tests.Stubs;
    using DecisionModel.Models.Policy;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.SubmissionShared.Enumerations;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using PolicyGeneration.Classes;

    /// <summary>
    /// Class TemplateProcessorTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class TemplateProcessorTests
    {
        /// <summary>
        /// Gets or sets the policy.
        /// </summary>
        /// <value>The policy.</value>
        private Policy Policy { get; set; }

        /// <summary>
        /// Gets or sets the policy json.
        /// </summary>
        /// <value>The policy json.</value>
        private JObject PolicyJson { get; set; }

        /// <summary>
        /// Gets or sets the custom mapping json.
        /// </summary>
        /// <value>The custom mapping json.</value>
        private string CustomMappingJson { get; set; }

        /// <summary>
        /// Setups this instance.
        /// </summary>
        [TestInitialize]
        public void Setup()
        {
            this.CreatePolicy();
            this.PolicyJson = JObject.Parse(JsonConvert.SerializeObject(this.Policy));
            this.CreateMappingJson();
        }

        /// <summary>
        /// Determines whether this instance [can get constant test].
        /// </summary>
        [TestMethod]
        public void CanGetConstantTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestConstant", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ABCD", val);
        }

        /// <summary>
        /// Determines whether this instance [can get sum test].
        /// </summary>
        [TestMethod]
        public void CanGetSumTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestSum", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("10000", val);
        }

        /// <summary>
        /// Determines whether this instance [can get format test].
        /// </summary>
        [TestMethod]
        public void CanGetFormatTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestSumFormatPosition1", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("$10,000", val, "TestSumFormatPosition1");

            val = processor.GetFieldValueFromCustom("TestSumFormatPosition2", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("$10,000", val, "TestSumFormatPosition2");

            val = processor.GetFieldValueFromCustom("TestSumFormatPosition3", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("$10,000", val, "TestSumFormatPosition3");
        }

        /// <summary>
        /// Determines whether this instance [can get json path test].
        /// </summary>
        [TestMethod]
        public void CanGetJsonPathTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestJsonPath", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("$500", val);
        }

        /// <summary>
        /// Determines whether this instance [can get empty on not found json path test].
        /// </summary>
        [TestMethod]
        public void CanGetEmptyOnNotFoundJsonPathTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestJsonPathReturnsEmpty", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual(string.Empty, val);
        }

        /// <summary>
        /// Determines whether this instance [can get if null value is null test].
        /// </summary>
        [TestMethod]
        public void CanGetIfNullValueIsNullTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfNullTrue", this.Policy, this.PolicyJson, ref docManager);

            // TODO: This should really return NULL but doesn't due to the fact that we don't have a way currently to send a null
            //          value in, but a custom function could return null as a value.
            Assert.AreEqual("NOTNULL", val);
        }

        /// <summary>
        /// Determines whether this instance [can get if null value is not null test].
        /// </summary>
        [TestMethod]
        public void CanGetIfNullValueIsNotNullTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfNullFalseString", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTNULL", val, "TestIfNullFalseString");

            val = processor.GetFieldValueFromCustom("TestIfNullFalseNumber", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTNULL", val, "TestIfNullFalseNumber");
        }

        /// <summary>
        /// Determines whether this instance [can get if true value is true test].
        /// </summary>
        [TestMethod]
        public void CanGetIfTrueValueIsTrueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfTrueTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("TRUE", val);
        }

        /// <summary>
        /// Determines whether this instance [can get if true value is not true test].
        /// </summary>
        [TestMethod]
        public void CanGetIfTrueValueIsNotTrueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfTrueFalseFalse", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTTRUE", val, "TestIfTrueFalseFalse");

            val = processor.GetFieldValueFromCustom("TestIfTrueFalseOther", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTTRUE", val, "TestIfNullFalseOther");

            val = processor.GetFieldValueFromCustom("TestIfTrueFalseNull", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTTRUE", val, "TestIfNullFalseNull");
        }

        /// <summary>
        /// Determines whether this instance [can get if not equal value is not equal test].
        /// </summary>
        [TestMethod]
        public void CanGetIfNotEqualValueIsNotEqualTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfNotEqualTrueString", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTEQUAL", val, "TestIfNotEqualTrueString");

            val = processor.GetFieldValueFromCustom("TestIfNotEqualTrueNumber", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTEQUAL", val, "TestIfNotEqualTrueNumber");

            val = processor.GetFieldValueFromCustom("TestIfNotEqualTrueNull", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NOTEQUAL", val, "TestIfNotEqualTrueNull");
        }

        /// <summary>
        /// Determines whether this instance [can get if not equal value is equal test].
        /// </summary>
        [TestMethod]
        public void CanGetIfNotEqualValueIsEqualTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfNotEqualFalseString", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("EQUAL", val, "TestIfNotEqualFalseString");

            val = processor.GetFieldValueFromCustom("TestIfNotEqualFalseNumber", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("EQUAL", val, "TestIfNotEqualFalseNumber");
        }

        /// <summary>
        /// Determines whether this instance [can get if not null empty whitespace value is not test].
        /// </summary>
        [TestMethod]
        public void CanGetIfNotNullEmptyWhitespaceValueIsNotTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfNotNullEmptyWhitespaceTrueString", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ISNOTNULLEMPTYWHITESPACE", val, "TestIfNotNullEmptyWhitespaceTrueString");

            val = processor.GetFieldValueFromCustom("TestIfNotNullEmptyWhitespaceTrueNumber", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ISNOTNULLEMPTYWHITESPACE", val, "TestIfNotNullEmptyWhitespaceTrueNumber");
        }

        /// <summary>
        /// Determines whether this instance [can join test].
        /// </summary>
        [TestMethod]
        public void CanJoinTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestJoin", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("A,B,C,D", val);
        }

        /// <summary>
        /// Determines whether this instance [can concatenate test].
        /// </summary>
        [TestMethod]
        public void CanConcatTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestConcat", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ABCD", val);
        }

        /// <summary>
        /// Determines whether this instance [can get if not null empty whitespace value is test].
        /// </summary>
        [TestMethod]
        public void CanGetIfNotNullEmptyWhitespaceValueIsTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestIfNotNullEmptyWhitespaceFalseEmptyString", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ISNULLEMPTYWHITESPACE", val, "TestIfNotNullEmptyWhitespaceFalseEmptyString");

            val = processor.GetFieldValueFromCustom("TestIfNotNullEmptyWhitespaceFalseWhitespace", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ISNULLEMPTYWHITESPACE", val, "TestIfNotNullEmptyWhitespaceFalseWhitespace");

            val = processor.GetFieldValueFromCustom("TestIfNotNullEmptyWhitespaceFalseNull", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ISNULLEMPTYWHITESPACE", val, "TestIfNotNullEmptyWhitespaceFalseNull");
        }

        /// <summary>
        /// Determines whether this instance [can get if simple and true test].
        /// </summary>
        [TestMethod]
        public void CanGetIfSimpleAndTrueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestSimpleAndTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ANDTRUE", val);
        }

        /// <summary>
        /// Determines whether this instance [can get if complex and true test].
        /// </summary>
        [TestMethod]
        public void CanGetIfComplexAndTrueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestComplexAnd2ParametersTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ANDTRUE", val, "TestComplexAnd2ParametersTrue");

            val = processor.GetFieldValueFromCustom("TestComplexAndManyParametersTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ANDTRUE", val, "TestComplexAndManyParametersTrue");
        }

        /// <summary>
        /// Determines whether this instance [can get if simple and false test].
        /// </summary>
        [TestMethod]
        public void CanGetIfSimpleAndFalseTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestSimpleAndFalse", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ANDFALSE", val);
        }

        /// <summary>
        /// Determines whether this instance [can get if complex and true test].
        /// </summary>
        [TestMethod]
        public void CanGetIfComplexAndFalseTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestComplexAnd2ParametersFalse", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ANDFALSE", val, "TestComplexAnd2ParametersFalse");

            val = processor.GetFieldValueFromCustom("TestComplexAndManyParametersFalse", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ANDFALSE", val, "TestComplexAndManyParametersFalse");
        }

        /// <summary>
        /// Determines whether this instance [can get if simple or true test].
        /// </summary>
        [TestMethod]
        public void CanGetIfSimpleOrTrueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestSimpleOrTrueAllTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ORTRUE", val, "TestSimpleOrTrueAllTrue");

            val = processor.GetFieldValueFromCustom("TestSimpleOrTrueSomeTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ORTRUE", val, "TestSimpleOrTrueSomeTrue");
        }

        /// <summary>
        /// Determines whether this instance [can get if complex or true test].
        /// </summary>
        [TestMethod]
        public void CanGetIfComplexOrTrueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestComplexOr2ParametersTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ORTRUE", val, "TestComplexOr2ParametersTrue");

            val = processor.GetFieldValueFromCustom("TestComplexOrManyParametersTrue", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ORTRUE", val, "TestComplexOrManyParametersTrue");
        }

        /// <summary>
        /// Determines whether this instance [can get if simple or false test].
        /// </summary>
        [TestMethod]
        public void CanGetIfSimpleOrFalseTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestSimpleOrFalse", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ORFALSE", val);
        }

        /// <summary>
        /// Determines whether this instance [can get if complex or false test].
        /// </summary>
        [TestMethod]
        public void CanGetIfComplexOrFalseTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestComplexOr2ParametersFalse", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ORFALSE", val, "TestComplexOr2ParametersFalse");

            val = processor.GetFieldValueFromCustom("TestComplexOrManyParametersFalse", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ORFALSE", val, "TestComplexOrManyParametersFalse");
        }

        /// <summary>
        /// Determines whether this instance [can get contains value].
        /// </summary>
        [TestMethod]
        public void CanGetContainsValueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestCONTAINS", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("NF", val, "NotFoundValue");
        }

        /// <summary>
        /// Determines whether this instance [can get join array value].
        /// </summary>
        [TestMethod]
        public void CanGetJOINARRAYValueTest()
        {
            IPolicyDocumentManager docManager = new PolicyDocumentManagerStub();
            var processor = new TemplateProcessor(new JsonManager());

            processor.CreateCustomDictionary(this.CustomMappingJson);

            string val = processor.GetFieldValueFromCustom("TestJOINARRAY", this.Policy, this.PolicyJson, ref docManager);
            Assert.AreEqual("ET,MUT,SE,ELRO", val, "JoinArrayValue");
        }

        /// <summary>
        /// Creates the policy.
        /// </summary>
        private void CreatePolicy()
        {
            // TODO: This should be moved to the PolicyGeneration piece whenever that is actually completed and
            //          can generate a policy matching your defined criteria.  For now lets just hard code a
            //          (not really complete) policy really quickly.  You will probably need to add additional
            //          fields and values to this as you write tests.
            var policy = PolicyHelper.GeneratePolicy(true);
            policy.IsInternal = true;
            policy.IsUnderwriterInitiated = true;
            policy.WorkingUserId = "workinguser@markeltest.com";
            policy.WorkingUserName = "My Working User Name";
            policy.PrimaryInsured = new Insured { Name = "Trans123" };
            policy.Agency = new Agency { Code = "210657", Name = "CRC Insurance Services, Inc." };
            policy.ProducerContact = new Producer { Email = "jim@210657_simoklahoma.com", Name = "Jim Sim" };
            policy.Term = new Term
            {
                Code = "A",
                Description = "12-month"
            };
            policy.LobOrder.Add(new LobType { Code = LineOfBusiness.IM });
            policy.ImLine = new ImLine();
            policy.ImLine.RollUp.Premium = 500;
            policy.ImLine.RiskUnits = new List<ImRiskUnit>();

            var riskUnit = new ImRiskUnit();
            policy.ImLine.RiskUnits.Add(riskUnit);
            riskUnit.State = "Virginia";
            riskUnit.Territory = "503";
            riskUnit.StateCode = "VA";
            riskUnit.ClassCode = "205";
            riskUnit.ClassDescription = "Contractor's Equipment";
            riskUnit.ClassSubCode = "00";
            riskUnit.Id = Guid.Parse("2ea0966b-ee7e-4b9b-bdb9-39c5d06169e2");
            riskUnit.PremiumBaseUnitMeasureQuantity = 0;
            riskUnit.Term = new Term { Code = "A", Description = "12-month" };

            riskUnit.PerOccurrenceDeductible = new ImDeductible();
            riskUnit.PerOccurrenceDeductible.Amount = 500;
            riskUnit.PerOccurrenceDeductible.UwAdjustedAmount = 600;
            riskUnit.PerOccurrenceDeductible.AgentMinimumAmount = 700;

            riskUnit.Items = new List<ImRatingItem>(4);
            riskUnit.Items.Add(
                new ImRatingItem()
                {
                    ItemCode = "ET",
                    ItemType = ImRateItemType.EmployeesTools,
                    IsSelected = false,
                    Premium = 175,
                    BaseMinimumPremium = 150,
                    BaseRate = 1.00m,
                    BaseRateUnscheduled = 1.05m,
                    BaseRateWithScheduled = 1.15m,
                    DevelopedRate = 1.30m,
                    UnderwriterAdjustedRate = 1.25m,
                    AgentAdjustedRate = 1.75m,
                    CauseOfLoss = null,
                    CauseOfLossCode = null,
                    Description = string.Empty,
                    TotalInsuredValue = 1000,
                    MaxLimit = 15000,
                    MaxSingleInsuredValue = 150,
                });
            riskUnit.Items.Add(
                new ImRatingItem()
                {
                    ItemCode = "MUT",
                    ItemType = ImRateItemType.MiscUnscheduledTools,
                    IsSelected = false,
                    Premium = 275,
                    BaseMinimumPremium = 250,
                    BaseRate = 2.00m,
                    BaseRateUnscheduled = 2.05m,
                    BaseRateWithScheduled = 2.15m,
                    DevelopedRate = 2.30m,
                    UnderwriterAdjustedRate = 2.25m,
                    AgentAdjustedRate = 2.75m,
                    CauseOfLoss = null,
                    CauseOfLossCode = null,
                    Description = string.Empty,
                    TotalInsuredValue = 2000,
                    MaxLimit = 25000,
                    MaxSingleInsuredValue = 250,
                });

            riskUnit.Items.Add(
                new ImRatingItem()
                {
                    ItemCode = "SE",
                    ItemType = ImRateItemType.ScheduledEquipment,
                    IsSelected = false,
                    Premium = 175,
                    BaseMinimumPremium = 350,
                    BaseRate = 3.00m,
                    BaseRateUnscheduled = 3.05m,
                    BaseRateWithScheduled = 3.15m,
                    DevelopedRate = 3.30m,
                    UnderwriterAdjustedRate = 3.25m,
                    AgentAdjustedRate = 3.75m,
                    CauseOfLoss = null,
                    CauseOfLossCode = null,
                    Description = string.Empty,
                    TotalInsuredValue = 3000,
                    MaxLimit = 35000,
                    MaxSingleInsuredValue = 350,
                });

            riskUnit.Items.Add(
                new ImRatingItem()
                {
                    ItemCode = "ELRO",
                    ItemType = ImRateItemType.EquipmentLeasedRented,
                    IsSelected = false,
                    Premium = 475,
                    BaseMinimumPremium = 450,
                    BaseRate = 4.00m,
                    BaseRateUnscheduled = 4.05m,
                    BaseRateWithScheduled = 4.15m,
                    DevelopedRate = 4.30m,
                    UnderwriterAdjustedRate = 4.25m,
                    AgentAdjustedRate = 4.75m,
                    CauseOfLoss = null,
                    CauseOfLossCode = null,
                    Description = string.Empty,
                    TotalInsuredValue = 4000,
                    MaxLimit = 45000,
                    MaxSingleInsuredValue = 450,
                });

            riskUnit.ContractorEquipmentRates = new List<ImContractorEquipmentRates>();
            riskUnit.ContractorEquipmentRates.Add(
                new ImContractorEquipmentRates
                {
                    ItemType = ImRateItemType.EmployeesTools,
                    ItemCode = "ET",
                    MinLimit = 0,
                    MaxLimit = 0,
                    Rate = 3,
                    RateUnscheduled = 4,
                    NonPackagedMinimumPremium = 500,
                    PackagedMinimumPremium = 350,
                });
            riskUnit.ContractorEquipmentRates.Add(
                new ImContractorEquipmentRates
                {
                    ItemType = ImRateItemType.MiscUnscheduledTools,
                    ItemCode = "MUT",
                    MinLimit = 0,
                    MaxLimit = 0,
                    Rate = 3,
                    RateUnscheduled = 4,
                    NonPackagedMinimumPremium = 500,
                    PackagedMinimumPremium = 350
                });
            riskUnit.ContractorEquipmentRates.Add(
                new ImContractorEquipmentRates
                {
                    ItemType = ImRateItemType.ScheduledEquipment,
                    ItemCode = "SE",
                    MinLimit = 100000,
                    MaxLimit = 0,
                    Rate = Convert.ToDecimal(1.25),
                    RateUnscheduled = Convert.ToDecimal(1.25),
                    NonPackagedMinimumPremium = 500,
                    PackagedMinimumPremium = 350
                });
            riskUnit.ContractorEquipmentRates.Add(
                new ImContractorEquipmentRates
                {
                    ItemType = ImRateItemType.EquipmentLeasedRented,
                    ItemCode = "ELRO",
                    MinLimit = 0,
                    MaxLimit = 0,
                    Rate = Convert.ToDecimal(1.5),
                    RateUnscheduled = Convert.ToDecimal(1.5),
                    NonPackagedMinimumPremium = 500,
                    PackagedMinimumPremium = 350
                });

            this.Policy = policy;
        }

        /// <summary>
        /// Creates the mapping json.
        /// </summary>
        private void CreateMappingJson()
        {
            var mapping =
                @"{
                    ""TestConstant"": {
                        ""constant"": ""ABCD""
                    },
                    ""TestSum"": {
                        ""function"": ""SUM"",
                        ""params"": [
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode == 'ET')].TotalInsuredValue""},
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='ELRO')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='MUT')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='SE')].TotalInsuredValue""}
                        ]
                    },
                    ""TestSumFormatPosition1"": {
		                ""format"": ""{0:C0}"",
                        ""function"": ""SUM"",
                        ""params"": [
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode == 'ET')].TotalInsuredValue""},
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='ELRO')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='MUT')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='SE')].TotalInsuredValue""}
                        ]
                    },
                    ""TestSumFormatPosition2"": {
                        ""function"": ""SUM"",
		                ""format"": ""{0:C0}"",
                        ""params"": [
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode == 'ET')].TotalInsuredValue""},
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='ELRO')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='MUT')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='SE')].TotalInsuredValue""}
                        ]
                    },
                    ""TestSumFormatPosition3"": {
                        ""function"": ""SUM"",
                        ""params"": [
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode == 'ET')].TotalInsuredValue""},
                            {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='ELRO')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='MUT')].TotalInsuredValue""},
			                {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[?(@.ItemCode=='SE')].TotalInsuredValue""}
                        ],
		                ""format"": ""{0:C0}"",
                    },
	                ""TestJsonPath"" : {
                        ""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].PerOccurrenceDeductible.Amount"",
		                ""format"": ""{0:C0}"",
	                },
	                ""TestJsonPathReturnsNull"" : {
                        ""jsonPath"": ""$.NotFoundReturnsNull"",
	                },
	                ""TestIfNullTrue"" : { 
                        ""function"": ""IFNULL"",
		                ""params"": [
			                {""constant"": """"},
			                {""constant"": ""NULL""},
			                {""constant"": ""NOTNULL""}
		                ]
	                },
	                ""TestIfNullFalseString"" : { 
                        ""function"": ""IFNULL"",
		                ""params"": [
			                {""constant"": ""A""},
			                {""constant"": ""NULL""},
			                {""constant"": ""NOTNULL""}
		                ]
	                },
	                ""TestIfNullFalseNumber"" : { 
                        ""function"": ""IFNULL"",
		                ""params"": [
			                {""constant"": 0},
			                {""constant"": ""NULL""},
			                {""constant"": ""NOTNULL""}
		                ]
	                },
	                ""TestIfTrueTrue"" : { 
                        ""function"": ""IFTRUE"",
		                ""params"": [
			                {""constant"": ""True""},
			                {""constant"": ""TRUE""},
			                {""constant"": ""NOTTRUE""}
		                ]
	                },
	                ""TestIfTrueFalseFalse"" : { 
                        ""function"": ""IFTRUE"",
		                ""params"": [
			                {""constant"": ""False""},
			                {""constant"": ""TRUE""},
			                {""constant"": ""NOTTRUE""}
		                ]
	                },
	                ""TestIfTrueFalseOther"" : { 
                        ""function"": ""IFTRUE"",
		                ""params"": [
			                {""constant"": ""ABC""},
			                {""constant"": ""TRUE""},
			                {""constant"": ""NOTTRUE""}
		                ]
	                },
	                ""TestIfTrueFalseNull"" : { 
                        ""function"": ""IFTRUE"",
		                ""params"": [
			                {""jsonPath"": ""$.IfTrue.TestReturnsNull""},
			                {""constant"": ""TRUE""},
			                {""constant"": ""NOTTRUE""}
		                ]
	                },
	                ""TestIfNotEqualTrueString"" : { 
                        ""function"": ""IFNOTEQUAL"",
		                ""params"": [
			                {""constant"": ""A""},
			                {""constant"": """"},
			                {""constant"": ""NOTEQUAL""},
			                {""constant"": ""EQUAL""}
		                ]
	                },
	                ""TestIfNotEqualFalseString"" : { 
                        ""function"": ""IFNOTEQUAL"",
		                ""params"": [
			                {""constant"": ""A""},
			                {""constant"": ""A""},
			                {""constant"": ""NOTEQUAL""},
			                {""constant"": ""EQUAL""}
		                ]
	                },
	                ""TestIfNotEqualTrueNumber"" : { 
                        ""function"": ""IFNOTEQUAL"",
		                ""params"": [
			                {""constant"": 0},
			                {""constant"": 1},
			                {""constant"": ""NOTEQUAL""},
			                {""constant"": ""EQUAL""}
		                ]
	                },
	                ""TestIfNotEqualFalseNumber"" : { 
                        ""function"": ""IFNOTEQUAL"",
		                ""params"": [
			                {""constant"": 0},
			                {""constant"": 0},
			                {""constant"": ""NOTEQUAL""},
			                {""constant"": ""EQUAL""}
		                ]
	                },
	                ""TestIfNotEqualTrueNull"" : { 
                        ""function"": ""IFNOTEQUAL"",
		                ""params"": [
			                {""constant"": 0},
			                {""jsonPath"": ""$.ReturnsANullValue""},
			                {""constant"": ""NOTEQUAL""},
			                {""constant"": ""EQUAL""}
		                ]
	                },
	                ""TestIfNotNullEmptyWhitespaceFalseEmptyString"" : { 
                        ""function"": ""IFNOTNULLEMPTYWHITESPACE"",
		                ""params"": [
			                {""constant"": """"},
			                {""constant"": ""ISNOTNULLEMPTYWHITESPACE""},
			                {""constant"": ""ISNULLEMPTYWHITESPACE""}
		                ]
	                },
	                ""TestIfNotNullEmptyWhitespaceFalseWhitespace"" : { 
                        ""function"": ""IFNOTNULLEMPTYWHITESPACE"",
		                ""params"": [
			                {""constant"": ""    ""},
			                {""constant"": ""ISNOTNULLEMPTYWHITESPACE""},
			                {""constant"": ""ISNULLEMPTYWHITESPACE""}
		                ]
	                },
	                ""TestIfNotNullEmptyWhitespaceFalseNull"" : { 
                        ""function"": ""IFNOTNULLEMPTYWHITESPACE"",
		                ""params"": [
			                {""jsonPath"": ""$.isnotnullemptywhitspace.testreturnsnull""},
			                {""constant"": ""ISNOTNULLEMPTYWHITESPACE""},
			                {""constant"": ""ISNULLEMPTYWHITESPACE""}
		                ]
	                },
	                ""TestIfNotNullEmptyWhitespaceTrueString"" : { 
                        ""function"": ""IFNOTNULLEMPTYWHITESPACE"",
		                ""params"": [
			                {""constant"": ""ABC""},
			                {""constant"": ""ISNOTNULLEMPTYWHITESPACE""},
			                {""constant"": ""ISNULLEMPTYWHITESPACE""}
		                ]
	                },
	                ""TestIfNotNullEmptyWhitespaceTrueNumber"" : { 
                        ""function"": ""IFNOTNULLEMPTYWHITESPACE"",
		                ""params"": [
			                {""constant"": 0},
			                {""constant"": ""ISNOTNULLEMPTYWHITESPACE""},
			                {""constant"": ""ISNULLEMPTYWHITESPACE""}
		                ]
	                },
	                ""TestJoin"" : { 
                        ""function"": ""JOIN"",
		                ""params"": [
			                {""constant"": "",""},
			                {""constant"": ""A""},
			                {""constant"": ""B""},
			                {""constant"": ""C""},
			                {""constant"": ""D""}
		                ]
	                },
	                ""TestConcat"" : { 
                        ""function"": ""CONCAT"",
		                ""params"": [
			                {""constant"": ""A""},
			                {""constant"": ""B""},
			                {""constant"": ""C""},
			                {""constant"": ""D""}
		                ]
	                },
	                ""TestSimpleAndTrue"" : { 
                        ""function"": ""AND"",
		                ""params"": [
			                {""constant"": ""true""},
			                {""constant"": ""true""},
			                {""constant"": ""ANDTRUE""},
			                {""constant"": ""ANDFALSE""}
		                ]
	                },
	                ""TestComplexAnd2ParametersTrue"" : { 
                        ""function"": ""AND"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""true""},
			                        {""constant"": ""true""}
                            ]},
			                {""constant"": ""ANDTRUE""},
			                {""constant"": ""ANDFALSE""}
		                ]
	                },
	                ""TestComplexAndManyParametersTrue"" : { 
                        ""function"": ""AND"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""}
                            ]},
			                {""constant"": ""ANDTRUE""},
			                {""constant"": ""ANDFALSE""}
		                ]
	                },
	                ""TestSimpleAndFalse"" : { 
                        ""function"": ""AND"",
		                ""params"": [
			                {""constant"": ""true""},
			                {""constant"": ""false""},
			                {""constant"": ""ANDTRUE""},
			                {""constant"": ""ANDFALSE""}
		                ]
	                },
	                ""TestComplexAnd2ParametersFalse"" : { 
                        ""function"": ""AND"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""false""},
			                        {""constant"": ""true""}
                            ]},
			                {""constant"": ""ANDTRUE""},
			                {""constant"": ""ANDFALSE""}
		                ]
	                },
	                ""TestComplexAndManyParametersFalse"" : { 
                        ""function"": ""AND"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""false""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""},
			                        {""constant"": ""true""}
                            ]},
			                {""constant"": ""ANDTRUE""},
			                {""constant"": ""ANDFALSE""}
		                ]
	                },
	                ""TestSimpleOrTrueAllTrue"" : { 
                        ""function"": ""OR"",
		                ""params"": [
			                {""constant"": ""true""},
			                {""constant"": ""true""},
			                {""constant"": ""ORTRUE""},
			                {""constant"": ""ORFALSE""}
		                ]
	                },
	                ""TestSimpleOrTrueSomeTrue"" : { 
                        ""function"": ""OR"",
		                ""params"": [
			                {""constant"": ""false""},
			                {""constant"": ""true""},
			                {""constant"": ""ORTRUE""},
			                {""constant"": ""ORFALSE""}
		                ]
	                },
	                ""TestComplexOr2ParametersTrue"" : { 
                        ""function"": ""OR"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""false""},
			                        {""constant"": ""true""}
                            ]},
			                {""constant"": ""ORTRUE""},
			                {""constant"": ""ORFALSE""}
		                ]
	                },
	                ""TestComplexOrManyParametersTrue"" : { 
                        ""function"": ""OR"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""true""}
                            ]},
			                {""constant"": ""ORTRUE""},
			                {""constant"": ""ORFALSE""}
		                ]
	                },
	                ""TestSimpleOrFalse"" : { 
                        ""function"": ""OR"",
		                ""params"": [
			                {""constant"": ""false""},
			                {""constant"": ""false""},
			                {""constant"": ""ORTRUE""},
			                {""constant"": ""ORFALSE""}
		                ]
	                },
	                ""TestComplexOr2ParametersFalse"" : { 
                        ""function"": ""OR"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""false""},
			                        {""constant"": ""false""}
                            ]},
			                {""constant"": ""ORTRUE""},
			                {""constant"": ""ORFALSE""}
		                ]
	                },
	                ""TestComplexOrManyParametersFalse"" : { 
                        ""function"": ""OR"",
		                ""params"": [
                            {""function"": ""CONDITION"",
                                ""params"": [
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""},
			                        {""constant"": ""false""}
                            ]},
			                {""constant"": ""ORTRUE""},
			                {""constant"": ""ORFALSE""}
		                ]
	                },
                    ""TestCONTAINS"" : { 
                        ""function"": ""CONTAINS"",
		                ""params"": [
                                      {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].RollUp.SelectedClauses""},
                                      {""constant"": ""C""},
                                      {""constant"": ""F""},
                                      {""constant"": ""NF""}
                        ]
	                },
                    ""TestJOINARRAY"" : { 
                        ""function"": ""JOINARRAY"",
		                ""params"": [
                                      {""constant"": "",""},
                                      {""jsonPath"": ""$.ImLine.RiskUnits[?(@.ClassCode == '205')].Items[*].ItemCode""}
                        ]
	                },
                }";

            ////this.CustomMappingJson = JsonManagerTests.GetJTokenizableString(mapping);
            this.CustomMappingJson = mapping;
        }
    }
}
