// <copyright file="IssuanceControllerTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Web.Http;
    using Common.Logging;
    using CorrespondenceServices.Controllers;
    using CorrespondenceServices.Tests.Stubs;
    using DecisionModel.Models.Policy;
    using DecisionModel.Representations;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Implementors;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.SubmissionShared.Enumerations;
    using Mkl.WebTeam.WebCore2;
    using Newtonsoft.Json;
    using SimpleInjector;
    using SimpleInjector.Integration.WebApi;

    /// <summary>
    /// Issuance controller tests
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class IssuanceControllerTests : IssuanceController
    {
        /// <summary>
        /// policy Id that's present in policy json
        /// </summary>
        private const string PolicyID = "e4a02c5a-8a81-48d7-a768-39d3df2b16ba";

        /// <summary>
        /// The file directory
        /// </summary>
        private const string FileDir = @"Files\";

        /// <summary>
        /// The container
        /// </summary>
        private Container container;

        /// <summary>
        /// Policy object
        /// </summary>
        private Policy policy;

        /// <summary>
        /// Initializes a new instance of the <see cref="IssuanceControllerTests" /> class
        /// </summary>
        public IssuanceControllerTests()
            : base(new Bootstrapper(), new JsonManager(), new RestHelperStub(), new StorageManagerStub(), new TemplateProcessor(new JsonManager()))
        {
            Aspose.Words.License license = new Aspose.Words.License();
            license.SetLicense("Aspose.Words.lic");
        }

        /// <summary>
        /// Gets override Carriers for testing
        /// </summary>
        protected override List<Carrier> Carriers
        {
            get
            {
                return new List<Carrier>()
                {
                    new Carrier()
                    {
                        Name = "Essex Insurance Company"
                    }
                };
            }
        }

        /// <summary>
        /// Test initializer
        /// </summary>
        [TestInitialize]
        public void TestInitialize()
        {
            this.ConfigureContainer();
            string policyText = File.ReadAllText(this.GetSourceFilePath() + "PolicyJson.txt");
            this.policy = JsonConvert.DeserializeObject<Policy>(policyText);
        }

        /// <summary>
        /// Generate MDIL 1002 form and save the file uploaded on storage provider to local "Output" folder
        /// </summary>
        ////[TestMethod]
        public void TEstMDIL1002Generated()
        {
            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    CompressionLevel = 0,
                    FormNormalizedNumber = "MDIL10020110",
                    Name = "MDIL 1002 01 10",
                    Path = @"E:\Consolidated DB\MDIL 1002 01 10 rev.docx",
                    Templated = true
                }
            };

            DecisionModel.Representations.MergeFormsRequest mergedFormsRequest = new DecisionModel.Representations.MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
            foreach (MergedForm mergedForm in arrFiles.UploadedFilePathAndName)
            {
                byte[] fileContents = this.StorageProviderManager.GetIssuanceDocumentBinary(
                    PolicyID,
                    Path.Combine(mergedForm.Path, mergedForm.Name));
                File.WriteAllBytes(this.GetOutputFilePath() + mergedForm.Name, fileContents);
            }
        }

        /// <summary>
        /// Generate MADUB 100 and 1003 together and save the files uploaded on storage provider to local "Output" folder
        /// </summary>
        ////[TestMethod]
        public void TEstMADUB1000And1003Generated()
        {
            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    CompressionLevel = 0,
                    FormNormalizedNumber = "MADUB10000115",
                    Name = "MADUB 1000",
                    Path = @"E:\Consolidated DB\MADUB 1000 01 15 101215.docx",
                    Templated = true
                },
                new MergedForm()
                {
                    CompressionLevel = 0,
                    FormNormalizedNumber = "MADUB10030115",
                    Name = "MADUB 1003 01 15",
                    Path = @"E:\Consolidated DB\MADUB 1003 01 15 REV 102915 (3).docx",
                    Templated = true
                },
            };

            DecisionModel.Representations.MergeFormsRequest mergedFormsRequest = new DecisionModel.Representations.MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
            foreach (MergedForm mergedForm in arrFiles.UploadedFilePathAndName)
            {
                byte[] fileContents = this.StorageProviderManager.GetIssuanceDocumentBinary(
                    PolicyID,
                    Path.Combine(mergedForm.Path, mergedForm.Name));
                File.WriteAllBytes(this.GetOutputFilePath() + mergedForm.Name, fileContents);
            }
        }

        /// <summary>
        /// Generate MADUB 100 and 1003 together and save the files uploaded on storage provider to local "Output" folder
        /// </summary>
        ////[TestMethod]
        public void FlattenExceptionTest()
        {
            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    CompressionLevel = 0,
                    FormNormalizedNumber = "MADB1000011",
                    Name = "MADUB 1000",
                    Path = @"E:\Consolidated DB\MADUB 1000 01 15 101215.docx",
                    Templated = true
                },
                new MergedForm()
                {
                    CompressionLevel = 0,
                    FormNormalizedNumber = "MAUB100301",
                    Name = "MADUB 1003 01 15",
                    Path = @"E:\Consolidated DB\MADUB 1003 01 15 REV 102915 (3).docx",
                    Templated = true
                },
                new MergedForm()
                {
                    CompressionLevel = 0,
                    FormNormalizedNumber = "MADUB10030115",
                    Name = "MADUB 1003 01 15",
                    Path = @"E:\Consolidated DB\MADUB 1003 01 15 REV 102915 (3).docx",
                    Templated = true
                },
            };

            DecisionModel.Representations.MergeFormsRequest mergedFormsRequest = new DecisionModel.Representations.MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
            foreach (MergedForm mergedForm in arrFiles.UploadedFilePathAndName)
            {
                byte[] fileContents = this.StorageProviderManager.GetIssuanceDocumentBinary(
                    PolicyID,
                    Path.Combine(mergedForm.Path, mergedForm.Name));
                File.WriteAllBytes(this.GetOutputFilePath() + mergedForm.Name, fileContents);
                Assert.Equals(mergedForm.ErrorMessages.Count, 0);
            }
        }

        /// <summary>
        /// Replacements the question test.
        /// </summary>
        [TestMethod]
        public void ReplacementQuestionTest()
        {
            Guid fakeId = Guid.NewGuid();

            string policyText = string.Empty;
            this.policy = new Policy()
            {
                CarrierName = "Essex Insurance Company",
                RollUp = new RollUp()
                {
                    Documents = new List<Document>(1)
                }
            };

            var doc = new Document()
            {
                NormalizedNumber = "IMMTC0030807REV",
                Questions = new List<Question>(),
            };

            // Replacing the value
            var q = new Question()
            {
                MergeFieldName = "BlankMergeField_12",
                Answers = new List<Answer>(),
                Display = true,
                DisplayFormat = DisplayFormat.DropDown,
                Order = 1,
                AnswerValue = "Other",
                Code = "BodyType"
            };

            var a = new Answer()
            {
                IsSelected = false,
                Order = 0,
                Value = "Car",
                Verbiage = "Car"
            };

            q.Answers.Add(a);

            a = new Answer()
            {
                IsSelected = false,
                Order = 1,
                Value = "Truck",
                Verbiage = "Truck"
            };

            q.Answers.Add(a);

            a = new Answer()
            {
                IsSelected = false,
                Order = 2,
                Value = "Other",
                Verbiage = "Other"
            };

            q.Answers.Add(a);
            doc.Questions.Add(q);

            q = new Question()
            {
                MergeFieldName = "BlankMergeField_12",
                Answers = new List<Answer>(),
                Display = true,
                DisplayFormat = DisplayFormat.TextBox,
                Order = 1,
                AnswerValue = "Monster Truck",
                Code = "OtherBodyType",
                ControllingQuestionCode = "BodyType",
                AnswerCodeReplacement = "Other"
            };

            doc.Questions.Add(q);

            q = new Question()
            {
                MergeFieldName = "BlankMergeField_10",
                Answers = new List<Answer>(),
                Display = true,
                DisplayFormat = DisplayFormat.DropDown,
                Order = 1,
                AnswerValue = "Ford",
                Code = "Manufacturer"
            };

            a = new Answer()
            {
                IsSelected = false,
                Order = 0,
                Value = "Ford",
                Verbiage = "Ford"
            };

            q.Answers.Add(a);

            a = new Answer()
            {
                IsSelected = false,
                Order = 1,
                Value = "Dodge",
                Verbiage = "Dodge"
            };

            q.Answers.Add(a);

            a = new Answer()
            {
                IsSelected = false,
                Order = 2,
                Value = "Other",
                Verbiage = "Other"
            };

            q.Answers.Add(a);
            doc.Questions.Add(q);

            q = new Question()
            {
                MergeFieldName = "BlankMergeField_10",
                Answers = new List<Answer>(),
                Display = true,
                DisplayFormat = DisplayFormat.TextBox,
                Order = 1,
                AnswerValue = null,
                Code = "OtherManufacturer",
                ControllingQuestionCode = "Manufacturer",
                AnswerCodeReplacement = "Other"
            };

            doc.Questions.Add(q);

            q = new Question()
            {
                MergeFieldName = "BlankMergeField_9",
                Answers = new List<Answer>(),
                Display = true,
                DisplayFormat = DisplayFormat.TextBox,
                Order = 1,
                AnswerValue = "1990",
                Code = "Year"
            };

            doc.Questions.Add(q);
            this.policy.RollUp.Documents.Add(doc);

            byte[] documentBytes = File.ReadAllBytes(@"Files\IM MTC 003 08 07 REV.docx");
            string documentBase64 = Convert.ToBase64String(documentBytes);
            RestHelperStub.DocumentTemplateBase64[fakeId.ToString()] = documentBase64;

            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    DocumentId = fakeId,
                    CompressionLevel = 0,
                    FormNormalizedNumber = "IMMTC0030807REV",
                    Name = "IMMTC0030807REV",
                    Path = @"Files\IM MTC 003 08 07 REV.docx",
                    Templated = true
                }
            };

            MergeFormsRequest mergedFormsRequest = new MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void SimpleReplacementTest()
        {
            Guid fakeId = Guid.NewGuid();

            string policyText = string.Empty;
            this.policy = new Policy()
            {
                CarrierName = "Essex Insurance Company",
                RollUp = new RollUp()
                {
                    Documents = new List<Document>(1)
                }
            };

            var doc = new Document()
            {
                NormalizedNumber = "TestForm",
                Questions = new List<Question>(),
            };

            var q = new Question()
            {
                MergeFieldName = "BlankMergeFieldTest",
                Answers = new List<Answer>(),
                Display = true,
                DisplayFormat = DisplayFormat.DropDown,
                Order = 1,
                AnswerValue = "Yes"
            };

            var a = new Answer()
            {
                IsSelected = false,
                Order = 1,
                Value = "Yes",
                Verbiage = "Yes"
            };

            q.Answers.Add(a);

            a = new Answer()
            {
                IsSelected = false,
                Order = 2,
                Value = "No",
                Verbiage = "No"
            };

            q.Answers.Add(a);

            a = new Answer()
            {
                IsSelected = false,
                Order = 3,
                Value = "Other",
                Verbiage = "Other"
            };

            q.Answers.Add(a);
            doc.Questions.Add(q);

            q = new Question()
            {
                MergeFieldName = "BlankMergeFieldTest",
                Display = true,
                DisplayFormat = DisplayFormat.TextBox,
                Order = 2
            };

            doc.Questions.Add(q);

            this.policy.RollUp.Documents.Add(doc);

            byte[] documentBytes = File.ReadAllBytes(@"Files\BlankMergeFieldTest.docx");
            string documentBase64 = Convert.ToBase64String(documentBytes);
            RestHelperStub.DocumentTemplateBase64[fakeId.ToString()] = documentBase64;

            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    DocumentId = fakeId,
                    CompressionLevel = 0,
                    FormNormalizedNumber = "TestForm",
                    Name = "TestForm",
                    Path = @"Files\BlankMergeFieldTest.docx",
                    Templated = true
                }
            };

            MergeFormsRequest mergedFormsRequest = new MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void CG2116ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{4CB22588-9329-40B2-A4EE-C8909AF4064F}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "CG2116", "CG 21 16 04 13.doc", fileFormat: DocumentManager.FileFormat.PDF);
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void CP1218ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{87B747C7-233E-4965-84CE-4FB0CFF1F3AC}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "CP1218", "CP 12 18 10 12.doc", fileFormat: DocumentManager.FileFormat.PDF);
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void CP1470ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{0929143B-C07F-4A75-A1C6-09D240CC3371}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "CP1470", "CP 14 70 10 12.doc", fileFormat: DocumentManager.FileFormat.PDF);
            ////arrFiles = this.RunFormGenerationTest(fakeId, "CP1470", "CP 14 70 10 12.doc", fileFormat: DocumentManager.FileFormat.DOC);
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void CG2144ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{FD0BD102-4AC2-404D-BF7B-AE6C2A3D1FCA}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "CG2144", "CG 21 44 07 98.doc", fileFormat: DocumentManager.FileFormat.PDF);
            ////arrFiles = this.RunFormGenerationTest(fakeId, "CG2144", "CG 21 44 07 98.doc", fileFormat: DocumentManager.FileFormat.PDF);
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void IMBMTCDEManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{04106B07-CFEA-4C2D-9750-3028A2020D50}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "IMBMTCDE", "IMB MTC-DE 07 03.doc", "IMBMTCDE2_JSON.txt");
        }

        /// <summary>
        /// Form IM MTC 003 generation test.
        /// </summary>
        [TestMethod]
        public void IMMTC003ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{43d0eb2b-f542-4c57-8f76-24ebdfd62ea4}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "IMMTC003", "IM MTC 003 08 07 REV.docx", "IMMTC003_JSON.txt", "IMMTC003_0807_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Form MEGL0009-01 generation test.
        /// </summary>
        [TestMethod]
        public void MEGL000901ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{3b94d4fd-326c-4a6e-bc1c-160c9088819b}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "MEGL000901", "MEGL 0009-01 05 16.doc", "MEGL0009_01_0516_JSON.txt", "MEGL0009_01_0516_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MEIM5216ManualFormGenerationTest()
        {
            var fakeId = new Guid("{A24D3DF1-4F16-4A49-A7AF-7E3A38CD61BC}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MEIM5216", "MEIM 5216 07 12.doc");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MDIL1001ManualFormGenerationTest()
        {
            var fakeId = new Guid("{B1AF384A-5992-459C-B21B-EBCBFAEA2266}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MDIL1001", "MDIL 1001 08 11.docx", "MDIL1001_JSON.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MEGL0263ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{16C619F7-0594-49BB-B5CB-7153C716AEBF}");

            ////var arrFiles = this.RunFormGenerationTest(fakeId, "MEGL0263", "MEGL 0263 03 14.doc", "MEGL0263_JSON.txt",  null, DocumentManager.FileFormat.DOC);
            var arrFiles = this.RunFormGenerationTest(fakeId, "MEGL0263", "MEGL 0263 03 14.doc", "MEGL0263_JSON.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MEGL0217ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{A0C1AEB0-9419-4A32-BEA3-A37B5A826CAB}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MEGL0217", "MEGL 0217 11 16.docx", "MEGL0263_JSON.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void CG2037ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{AA8E881A-E758-4254-8047-9E4EB6FFD3E4}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "CG2037", "CG 20 37 04 13.doc", "CG2037_JSON.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MDGL1008ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{5186F9D4-A7A5-497E-9622-73382A75D686}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MDGL1008", "MDGL 1008 08 11-rev.docx", "MDGL1008_JSON.txt", "MDGL1008_0811_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MECP1244ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{2fc01df5-4dcf-4d77-b0df-9c4418626b26}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MECP1244", "MECP 1244 04 16.docx", "MECP1244_JSON.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MECP1321ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{448F57DF-5C28-4F95-95C6-554757BB74A6}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MECP1321", "MECP 1321 02 16.docx", "MECP1321_JSON.txt");
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MDIL1005ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{8D44BA4A-6C78-4226-935F-006E8E377484}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MDIL1005", "MDIL 1005 08 14.doc", "MDIL1005_JSON.txt");
            ////arrFiles = this.RunFormGenerationTest(fakeId, "MDIL1005", "MDIL 1005 08 14.doc", "MDIL1005_JSON.txt", null, DocumentManager.FileFormat.DOC);
        }

        /// <summary>
        /// Simples the replacement test.
        /// </summary>
        [TestMethod]
        public void MEIL1205ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{F057D26C-A7F4-42D7-BE37-53D9796FD886}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MEIL1205", "MEIL 1205 07 13.doc", "MEIL1205_JSON.txt");
        }

        /// <summary>
        /// Runs the following form tests:
        ///     CP1033
        ///     CP1054
        ///     CP1055
        ///     CP1056
        /// </summary>
        [TestMethod]
        public void CP10_33_54_55_56_ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{2E43AD7F-F48D-46BB-82EE-595877A12769}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "CP1033", "CP 10 33 10 12.doc", "CP10-33-54-55-56_JSON.txt");

            fakeId = new Guid("{28CB9E88-E74A-48FE-BEEA-048C4DA9429A}");
            arrFiles = this.RunFormGenerationTest(fakeId, "CP1054", "CP 10 54 06 07.doc", "CP10-33-54-55-56_JSON.txt");

            fakeId = new Guid("{D1B87FE0-722E-4DF9-90FD-5EED47D0D117}");
            arrFiles = this.RunFormGenerationTest(fakeId, "CP1055", "CP 10 55 06 07.doc", "CP10-33-54-55-56_JSON.txt");

            fakeId = new Guid("{F0672340-B1FA-4A03-845D-0E9673FC6A9D}");
            arrFiles = this.RunFormGenerationTest(fakeId, "CP1056", "CP 10 56 06 07.doc", "CP10-33-54-55-56_JSON.txt");
        }

        /// <summary>
        /// Manuals the form generation test.
        /// </summary>
        [TestMethod]
        public void IMCEARManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{04106B07-CFEA-4C2D-9750-3028A2020D50}");
            string policyText = File.ReadAllText(@"Files\IMCEAR_JSON.txt");
            this.policy = JsonConvert.DeserializeObject<Policy>(policyText);

            string customMapping = File.ReadAllText(@"Files\IMCEAR_CUSTOMMAPPING.txt").Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Replace("  ", " ");
            Dictionary<Guid, string> responseObject = new Dictionary<Guid, string>();
            responseObject.Add(fakeId, customMapping);
            var jsonResponse = JsonConvert.SerializeObject(responseObject);

            RestHelperStub.CustomMappingTemplates[fakeId.ToString()] = jsonResponse;

            byte[] documentBytes = File.ReadAllBytes(@"Files\IM CEAR 05 09 REV.docx");
            string documentBase64 = Convert.ToBase64String(documentBytes);
            RestHelperStub.DocumentTemplateBase64[fakeId.ToString()] = documentBase64;

            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    DocumentId = fakeId,
                    CompressionLevel = 0,
                    FormNormalizedNumber = "IMCEAR0509REV",
                    Name = "IM CEAR 05 09 REV",
                    Path = @"Files\IM CEAR 05 09 REV.docx",
                    Templated = true
                }
            };

            DecisionModel.Representations.MergeFormsRequest mergedFormsRequest = new DecisionModel.Representations.MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
        }

        /// <summary>
        /// Form IL1201 generation test.
        /// </summary>
        [TestMethod]
        public void IL1201ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{3e9e0409-8f84-4a97-86e3-29ac460c0296}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "IL1201", "IL 12 01 11 85.docx", "IL1201_JSON.txt", "IL1201_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Form MDIL 1000 generation test.
        /// </summary>
        [TestMethod]
        public void MDIL1000ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{a138a7b5-46aa-4e08-a1a5-0a20490d5ed0}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "MDIL1000", "MDIL 1000 08 11SL.doc", "MDIL1000_JSON.txt", "MDIL1000_0811_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Form MEGL0257 0516 generation test.
        /// </summary>
        [TestMethod]
        public void MEGL0257_0516_ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{d807db87-9933-4b85-8e44-8e675755eb0e}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "MEGL0257", "MEGL 0257 05 16.docx", "MEGL0257_0516_JSON.txt", "MEGL0257_0516_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Form MECP1200 0914 generation test.
        /// </summary>
        [TestMethod]
        public void MECP1200_0914_ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{7ece2d51-f022-4371-ad2c-2d7b728ff85a}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "MECP1200", "MECP 1200 09 14.docx", "MECP1200_0914_JSON.txt");
        }

        /// <summary>
        /// Manuals the form generation test for MEGL1547.
        /// </summary>
        [TestMethod]
        public void TestFormGeneration_MEGL1547()
        {
            Guid fakeId = new Guid("{04106B07-CFEA-4C2D-9750-3028A2020D50}");
            string policyText = File.ReadAllText(@"Policies\PolicyJson - MEGL1547.txt");
            this.policy = JsonConvert.DeserializeObject<Policy>(policyText);

            string customMapping = File.ReadAllText(@"Files\MEGL1547_CUSTOMMAPPING.txt").Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Replace("  ", " ");
            Dictionary<Guid, string> responseObject = new Dictionary<Guid, string>();
            responseObject.Add(fakeId, customMapping);
            var jsonResponse = JsonConvert.SerializeObject(responseObject);

            RestHelperStub.CustomMappingTemplates[fakeId.ToString()] = jsonResponse;

            byte[] documentBytes = File.ReadAllBytes(@"Files\MEGL 1547 05 16.docx");
            string documentBase64 = Convert.ToBase64String(documentBytes);
            RestHelperStub.DocumentTemplateBase64[fakeId.ToString()] = documentBase64;

            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    DocumentId = fakeId,
                    CompressionLevel = 0,
                    FormNormalizedNumber = "MEGL1547",
                    Name = "MEGL 1547",
                    Path = @"Files\MEGL 1547 05 16.docx",
                    Templated = true
                }
            };

            DecisionModel.Representations.MergeFormsRequest mergedFormsRequest = new DecisionModel.Representations.MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
        }

        /// <summary>
        /// Form MEGL1844 0614 generation test.
        /// </summary>
        [TestMethod]
        public void MEGL1844_0611_ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{307ebda6-46af-4ceb-9f22-00b00877f51b}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "MEGL1844", "MEGL 1844 06 14.docx", "MEGL1844_0614_JSON.txt", "MEGL1844_0614_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Form MPIL1039 generation test.
        /// </summary>
        [TestMethod]
        public void MPIL1039ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{4a02f9e8-596c-4dd2-9f7a-15c4804e0b28}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "MPIL1039", "MPIL 1039 01 12.doc", "MPIL1039_JSON.txt", "MPIL1039_0112_CUSTOMMAPPING.txt");
        }

        /// <summary>
        /// Form CP1036 generation test.
        /// </summary>
        [TestMethod]
        public void CP1036_ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{307ebda6-46af-4ceb-9f22-00b00877f51b}");
            var arrFiles = this.RunFormGenerationTest(fakeId, "CP1036", "CP 10 36 10 12.doc", "CP1036_JSON.txt");
        }

        /// <summary>
        /// Manuals the form generation test for IMCENP.
        /// </summary>
        [TestMethod]
        public void TestFormGeneration_IMCENP()
        {
            Guid fakeId = new Guid("{04106B07-CFEA-4C2D-9750-3028A2020D50}");
            string policyText = File.ReadAllText(@"Policies\PolicyJson - IMCENP.txt");
            this.policy = JsonConvert.DeserializeObject<Policy>(policyText);

            string customMapping = File.ReadAllText(@"Files\IMCENP_CUSTOMMAPPING.txt").Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').Replace("  ", " ");
            Dictionary<Guid, string> responseObject = new Dictionary<Guid, string>();
            responseObject.Add(fakeId, customMapping);
            var jsonResponse = JsonConvert.SerializeObject(responseObject);

            RestHelperStub.CustomMappingTemplates[fakeId.ToString()] = jsonResponse;

            byte[] documentBytes = File.ReadAllBytes(@"Files\IM CENP 05 09.doc");
            string documentBase64 = Convert.ToBase64String(documentBytes);
            RestHelperStub.DocumentTemplateBase64[fakeId.ToString()] = documentBase64;

            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    DocumentId = fakeId,
                    CompressionLevel = 0,
                    FormNormalizedNumber = "IMCENP",
                    Name = "IM CENP",
                    Path = @"Files\IM CENP 05 09.doc",
                    Templated = true
                }
            };

            DecisionModel.Representations.MergeFormsRequest mergedFormsRequest = new DecisionModel.Representations.MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            MergeDocumentsResponse arrFiles = this.SaveDocs(mergedFormsRequest, DocumentManager.FileFormat.PDF, true);
        }

        /// <summary>
        /// Form CP0450 generation test.
        /// </summary>
        [TestMethod]
        public void CP0450_ManualFormGenerationTest()
        {
            Guid fakeId = new Guid("{B8F69152-8E9D-4A44-99BB-39DB83673B99}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "CP0450", "CP 04 50 07 88.docx", "CP0450_JSON.txt");
            arrFiles = this.RunFormGenerationTest(fakeId, "CP0450", "CP 04 50 07 88.docx", "CP0450_JSON.txt", null, DocumentManager.FileFormat.DOCX);
        }

        /// <summary>
        /// Test form generation for MEIM 5202.
        /// </summary>
        [TestMethod]
        public void MEIM5202ManualFormGenerationTest()
        {
            var fakeId = new Guid("{f4a1bf9e-70b3-4d38-9330-a9a32a7e86e6}");

            var arrFiles = this.RunFormGenerationTest(fakeId, "MEIM5202", "MEIM 5202 11 09.doc");
        }

        /// <summary>
        /// get source file path
        /// </summary>
        /// <returns>string</returns>
        private string GetSourceFilePath()
        {
            string dir = Directory.GetCurrentDirectory();
            return dir.Remove(dir.IndexOf(@"\bin")) + @"\Policies\";
        }

        /// <summary>
        /// Get output file path
        /// </summary>
        /// <returns>string</returns>
        private string GetOutputFilePath()
        {
            string dir = Directory.GetCurrentDirectory();
            return dir.Remove(dir.IndexOf(@"\bin")) + @"\Output\";
        }

        /// <summary>
        /// Configures the container.
        /// </summary>
        private void ConfigureContainer()
        {
            this.container = new Container();

            this.container.Options.AllowOverridingRegistrations = true;
            this.container.RegisterWebApiRequest<IBootstrapper, Bootstrapper>();
            this.container.RegisterWebApiRequest<IJsonManager, JsonManager>();
            this.container.RegisterWebApiRequest<IRestHelper, RestHelperStub>();
            this.container.RegisterWebApiRequest<IStorageManager, StorageManagerStub>();
            this.container.RegisterSingleton<ILog>(() => LogEvent.Log);

            this.container.Verify();

            GlobalConfiguration.Configuration.DependencyResolver = new SimpleInjectorWebApiDependencyResolver(this.container);
        }

        /// <summary>
        /// Runs the form generation test.
        /// </summary>
        /// <param name="testGuid">The test unique identifier.</param>
        /// <param name="normalizedNumber">The form normalized number.</param>
        /// <param name="formName">Name of the form.</param>
        /// <param name="policyJsonFilename">The policy json filename.</param>
        /// <param name="customMappingFilename">The custom mapping filename.</param>
        /// <param name="fileFormat">The file format.</param>
        /// <returns>MergeDocumentsResponse.</returns>
        private MergeDocumentsResponse RunFormGenerationTest(Guid testGuid, string normalizedNumber, string formName, string policyJsonFilename = null, string customMappingFilename = null, DocumentManager.FileFormat fileFormat = DocumentManager.FileFormat.PDF)
        {
            if (string.IsNullOrWhiteSpace(policyJsonFilename))
            {
                policyJsonFilename = $"{normalizedNumber}_JSON.txt";
            }

            var policyText = File.ReadAllText($@"Files\{policyJsonFilename}");
            this.policy = JsonConvert.DeserializeObject<Policy>(policyText);

            if (string.IsNullOrWhiteSpace(customMappingFilename))
            {
                customMappingFilename = $@"Files\{normalizedNumber}_CUSTOMMAPPING.txt";
            }
            else if (!customMappingFilename.Contains($@"Files\"))
            {
                customMappingFilename = FileDir + customMappingFilename;
            }

            if (File.Exists(customMappingFilename))
            {
                var customMapping = File.ReadAllText(customMappingFilename)
                    .Replace('\t', ' ')
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace("  ", " ");

                Dictionary<Guid, string> responseObject = new Dictionary<Guid, string>
                {
                    { testGuid, customMapping }
                };
                var jsonResponse = JsonConvert.SerializeObject(responseObject);
                RestHelperStub.CustomMappingTemplates[testGuid.ToString()] = jsonResponse;
            }
            else
            {
                RestHelperStub.CustomMappingTemplates[testGuid.ToString()] = null;
            }

            var filename = $@"Files\{formName}";
            var documentBytes = File.ReadAllBytes(filename);
            var documentBase64 = Convert.ToBase64String(documentBytes);
            RestHelperStub.DocumentTemplateBase64[testGuid.ToString()] = documentBase64;

            List<MergedForm> mergedForms = new List<MergedForm>()
            {
                new MergedForm()
                {
                    DocumentId = testGuid,
                    CompressionLevel = 0,
                    FormNormalizedNumber = normalizedNumber,
                    Name = normalizedNumber,
                    Path = filename,
                    Templated = true
                }
            };

            var mergedFormsRequest = new MergeFormsRequest()
            {
                MergedForms = mergedForms,
                Policy = this.policy
            };

            return this.SaveDocs(mergedFormsRequest, fileFormat, true);
        }
    }
}
