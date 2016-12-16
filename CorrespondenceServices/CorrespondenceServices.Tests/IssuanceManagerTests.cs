// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-16-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-16-2016
// ***********************************************************************
// <copyright file="IssuanceManagerTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CorrespondenceServices.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using Aspose.Words;
    using Common.Logging;
    using Common.Logging.Simple;
    using CorrespondenceServices.Classes;
    using CorrespondenceServices.Tests.Stubs;
    using DecisionModel.Models.Policy;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.RestHelper.Classes;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Implementors;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.SubmissionShared.Enumerations;
    using Moq;
    using Newtonsoft.Json.Linq;
    using PolicyGeneration.Classes;
    using SimpleInjector;
    using pol = DecisionModel.Models.Policy;

    /// <summary>
    /// Class IssuanceManagerTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class IssuanceManagerTests
    {
        /// <summary>
        /// Gets or sets the authenticated user.
        /// </summary>
        /// <value>The authenticated user.</value>
        private static IUser AuthenticatedUser { get; set; }

        /// <summary>
        /// Gets or sets the carrier.
        /// </summary>
        /// <value>The carrier.</value>
        private static Carrier Carrier { get; set; }

        /// <summary>
        /// Gets or sets the container.
        /// </summary>
        /// <value>The container.</value>
        private Container Container { get; set; }

        /// <summary>
        /// Gets or sets the custom mappings.
        /// </summary>
        /// <value>The custom mappings.</value>
        private Dictionary<string, JToken> CustomMappings { get; set; }

        /// <summary>
        /// Initializes this instance.
        /// </summary>
        /// <param name="context">The context.</param>
        [ClassInitialize]
        public static void Initialize(TestContext context)
        {
            SetAuthenticatedUser();
            Carrier = new Carrier()
            {
                Name = "Essex Insurance Company"
            };
        }

        /// <summary>
        /// Tests the initialize.
        /// </summary>
        [TestInitialize]
        public void TestInitialize()
        {
            this.ConfigureContainer();
        }

        #region GetWorkingQuestion

        /// <summary>
        /// Determines whether the original question is returned if there are no other questions that are overriding this question
        /// and there are not multiple row groupings (single instance of this particular question)
        /// </summary>
        [TestMethod]
        public void CanGetWorkingQuestionReturnsSameQuestionNoMultipleRowGroupingNumber()
        {
            var answerOther = CreateAnswer("MergeField_1", "Other");
            var answerSomething = CreateAnswer("MergeField_2", "Something");

            var controllingQuestion = CreateQuestion(
                DisplayFormat.DropDown,
                "Test1",
                "This is a controlling question",
                "MergeField_1",
                answerValue: "Something",
                answers: new List<Answer>(new[] { answerSomething, answerOther }));

            var overridingQuestion = CreateQuestion(
                DisplayFormat.TextBox,
                "Test2",
                "This is an overriding question",
                "MergeField_1",
                controllingCode: "Test1",
                answerCodeReplacement: "Other");

            var doc = new pol.Document()
            {
                Questions = new List<Question>()
            };
            doc.Questions.Add(controllingQuestion);
            doc.Questions.Add(overridingQuestion);

            var po = new PrivateType(typeof(IssuanceManager));
            var args = new object[] { controllingQuestion, doc };
            var replacementQuestion = (Question)po.InvokeStatic("GetWorkingQuestion", args);

            Assert.AreEqual(replacementQuestion, controllingQuestion, $"Question {replacementQuestion.Verbiage}");
        }

        /// <summary>
        /// Determines whether the overriding question that has the correct answer selected is returned
        /// and there are not multiple row groupings (single instance of this particular question)
        /// </summary>
        [TestMethod]
        public void CanGetWorkingQuestionReturnsOverridingQuestionCorrectAnswerNoMultipleRowGroupingNumber()
        {
            var answerOther = CreateAnswer("MergeField_2", "Other", true);
            var answerSomething = CreateAnswer("MergeField_2", "Something");

            var controllingQuestion = CreateQuestion(
                DisplayFormat.DropDown,
                "Test1",
                "This is a controlling question",
                "MergeField_1",
                answerValue: "Other",
                answers: new List<Answer>(new[] { answerSomething, answerOther }));

            var overridingQuestion = CreateQuestion(
                DisplayFormat.TextBox,
                "Test2",
                "This is an overriding question",
                "MergeField_2",
                answerValue: "I have answered Other",
                controllingCode: "Test1",
                answerCodeReplacement: "Other");

            var doc = new pol.Document()
            {
                Questions = new List<Question>(new[] { controllingQuestion, overridingQuestion })
            };

            var po = new PrivateType(typeof(IssuanceManager));
            object[] args = new object[] { controllingQuestion, doc };
            var replacementQuestion = (Question)po.InvokeStatic("GetWorkingQuestion", args);

            Assert.AreEqual(replacementQuestion, overridingQuestion, message: $"Question {replacementQuestion.Verbiage}");
        }

        /// <summary>
        /// Determines whether the controlling question that has the incorrect answer selected is returned when there is an overriding
        /// question and there are not multiple row groupings (single instance of this particular question)
        /// </summary>
        [TestMethod]
        public void CanGetWorkingQuestionReturnsControllingQuestionWithOverridingQuestionWrongAnswerNoMultipleRowGroupingNumber()
        {
            var answerOther = CreateAnswer("MergeField_2", "Other");
            var answerSomething = CreateAnswer("MergeField_2", "Something", true);

            var controllingQuestion = CreateQuestion(
                DisplayFormat.DropDown,
                "Test1",
                "This is a controlling question",
                "MergeField_1",
                answerValue: "Something",
                answers: new List<Answer>(new[] { answerOther, answerSomething }));

            var overridingQuestion = CreateQuestion(
                DisplayFormat.TextBox,
                "Test2",
                "This is an overriding question",
                "MergeField_2",
                controllingCode: "Test1",
                answerCodeReplacement: "Other",
                answerValue: "I have answered Other");

            var doc = new pol.Document()
            {
                Questions = new List<Question>(new[] { controllingQuestion, overridingQuestion })
            };

            var po = new PrivateType(typeof(IssuanceManager));
            object[] args = new object[] { controllingQuestion, doc };
            var replacementQuestion = (Question)po.InvokeStatic("GetWorkingQuestion", args);

            Assert.AreEqual(replacementQuestion, controllingQuestion, message: $"Question {replacementQuestion.Verbiage}");
        }

        /// <summary>
        /// Determines whether the overriding question that has the correct answer selected is returned
        /// when there are multiple row groupings
        /// </summary>
        [TestMethod]
        public void CanGetWorkingQuestionReturnsOverridingQuestionCorrectAnswerWithMultiplRowGroupingNumber()
        {
            var answerOther = CreateAnswer("MergeField_2", "Other", true);
            var answerSomething = CreateAnswer("MergeField_2", "Something");

            var controllingQuestion = CreateQuestion(
                DisplayFormat.DropDown,
                "Test1",
                "This is a controlling question",
                "MergeField_2",
                2,
                answerValue: "Other",
                answers: new List<Answer>(new[] { answerOther, answerSomething }));

            var overridingQuestion = CreateQuestion(
                DisplayFormat.TextBox,
                "Test2",
                "This is an overriding question",
                "MergeField_1",
                2,
                0,
                "I have answered Other",
                "Test1",
                "Other");

            var answerOtherAdditional = CreateAnswer("MergeField_2", "Other", true);
            var answerSomethingAdditional = CreateAnswer("MergeField_2", "Something");

            var controllingQuestionAdditional = CreateQuestion(
                DisplayFormat.DropDown,
                "Test1",
                "This is an additional controlling question",
                "MergeField_2",
                1,
                answerValue: "Other",
                answers: new List<Answer>(new[] { answerOtherAdditional, answerSomethingAdditional }));

            var overridingQuestionAdditional = CreateQuestion(
                DisplayFormat.TextBox,
                "Test2",
                "This is an additional overriding question",
                "MergeField_1",
                1,
                0,
                "I have answered Other additional",
                "Test1",
                "Other");

            var doc = new pol.Document()
            {
                Questions = new List<Question>(new[] { controllingQuestionAdditional, overridingQuestionAdditional, controllingQuestion, overridingQuestion })
            };

            var po = new PrivateType(typeof(IssuanceManager));
            object[] args = new object[] { controllingQuestion, doc };
            var replacementQuestion = (Question)po.InvokeStatic("GetWorkingQuestion", args);

            Assert.AreEqual(replacementQuestion, overridingQuestion, message: $"Question {replacementQuestion.Verbiage}");
        }

        /// <summary>
        /// Determines whether the overriding question that has the correct answer selected is returned
        /// when there are multiple row groupings
        /// </summary>
        [TestMethod]
        public void CanGetWorkingQuestionReturnsSameOverridingQuestionQuestionWrongAnswerWithMultiplRowGroupingNumber()
        {
            var answerOther = CreateAnswer("MergeField_2", "Other");
            var answerSomething = CreateAnswer("MergeField_2", "Something", true);

            var controllingQuestion = CreateQuestion(
                DisplayFormat.DropDown,
                "Test1",
                "This is a controlling question",
                "MergeField_2",
                2,
                0,
                "Something",
                answers: new List<Answer>(new[] { answerSomething, answerOther }));

            var overridingQuestion = CreateQuestion(
                DisplayFormat.TextBox,
                "Test2",
                "This is an overriding question",
                "MergeField_1",
                2,
                controllingCode: "Test1",
                answerCodeReplacement: "Other");

            var answerOtherAdditional = CreateAnswer("MergeField_2", "Other", true);
            var answerSomethingAdditional = CreateAnswer("MergeField_2", "Something");

            var controllingQuestionAdditional = CreateQuestion(
                DisplayFormat.DropDown,
                "Test1",
                "This is an additional controlling question",
                "MergeField_2",
                1,
                0,
                "Something",
                answers: new List<Answer>(new[] { answerSomethingAdditional, answerOtherAdditional }));

            var overridingQuestionAdditional = CreateQuestion(
                DisplayFormat.TextBox,
                "Test2",
                "This is an additional overriding question",
                "MergeField_2",
                1,
                controllingCode: "Test1",
                answerCodeReplacement: "Other",
                answerValue: "Additional answer value");

            var doc = new pol.Document()
            {
                Questions = new List<Question>(new[] { controllingQuestionAdditional, overridingQuestionAdditional, controllingQuestion, overridingQuestion })
            };

            var po = new PrivateType(typeof(IssuanceManager));
            object[] args = new object[] { controllingQuestion, doc };
            var replacementQuestion = (Question)po.InvokeStatic("GetWorkingQuestion", args);

            Assert.AreEqual(replacementQuestion, controllingQuestion, message: $"Question {replacementQuestion.Verbiage}");
        }

        #endregion

        #region GetNumberOfFormInstancesRequired

        /// <summary>
        /// Determines whether 1 instance of the form is required if no questions
        /// and no maximum multiple row count are assigned.
        /// </summary>
        [TestMethod]
        public void CanGetNumberOfFormInstancesRequiredWithNoQuestions()
        {
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var mergedForm = new MergedForm()
            {
                FormNormalizedNumber = normalizedNumber
            };
            var args = new object[] { mergedForm };
            var retVal = (int)po.Invoke("GetNumberOfFormInstancesRequired", args);
            Assert.AreEqual(1, retVal);
        }

        /// <summary>
        /// Determines whether 1 instance of the form is required if a question with no
        /// maximum multiple row count
        /// </summary>
        [TestMethod]
        public void CanGetNumberOfFormInstancesRequiredWithOneQuestionsAndNoMaximumMultipleRowCount()
        {
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
                Questions = new List<Question>(new[]
                {
                    new Question()
                    {
                        Code = "Test1"
                    }
                })
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var mergedForm = new MergedForm()
            {
                FormNormalizedNumber = normalizedNumber
            };
            var args = new object[] { mergedForm };
            var retVal = (int)po.Invoke("GetNumberOfFormInstancesRequired", args);
            Assert.AreEqual(1, retVal);
        }

        /// <summary>
        /// Determines whether 1 instance of the form is required if one questions is assigned and there are less instances
        /// than MaximumMultipleRowCount and MultipleRowGroupingNumber starts at 1
        /// </summary>
        [TestMethod]
        public void CanGetNumberOfFormInstancesRequiredWithOneQuestionIncremental()
        {
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
                Questions = new List<Question>(new[]
                {
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 1
                    }
                })
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var mergedForm = new MergedForm()
            {
                FormNormalizedNumber = normalizedNumber
            };
            var args = new object[] { mergedForm };
            var retVal = (int)po.Invoke("GetNumberOfFormInstancesRequired", args);
            Assert.AreEqual(1, retVal);
        }

        /// <summary>
        /// Determines whether 1 instance of the form is required if one questions is assigned and there are less instances
        /// than MaximumMultipleRowCount and the MultipleRowGroupingNumber does not start at 1
        /// </summary>
        [TestMethod]
        public void CanGetNumberOfFormInstancesRequiredWithOneQuestionNotIncremental()
        {
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
                Questions = new List<Question>(new[]
                {
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 2
                    }
                })
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var mergedForm = new MergedForm()
            {
                FormNormalizedNumber = normalizedNumber
            };
            var args = new object[] { mergedForm };
            var retVal = (int)po.Invoke("GetNumberOfFormInstancesRequired", args);
            Assert.AreEqual(1, retVal);
        }

        /// <summary>
        /// Determines whether 1 instance of the form is required if one questions is assigned and there are less instances
        /// than MaximumMultipleRowCount
        /// </summary>
        [TestMethod]
        public void CanGetNumberOfFormInstancesRequiredWithOneQuestionGroupingGreaterThanMaximum()
        {
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
                Questions = new List<Question>(new[]
                {
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 5
                    }
                })
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var mergedForm = new MergedForm()
            {
                FormNormalizedNumber = normalizedNumber
            };
            var args = new object[] { mergedForm };
            var retVal = (int)po.Invoke("GetNumberOfFormInstancesRequired", args);
            Assert.AreEqual(1, retVal);
        }

        /// <summary>
        /// Determines whether 1 instance of the form is required if the same number questions is assigned
        /// as the MaximumMultipleRowCount
        /// </summary>
        [TestMethod]
        public void CanGetNumberOfFormInstancesRequiredWithNumberOfQuestionsEqualsMaximum()
        {
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
                Questions = new List<Question>(new[]
                {
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 5
                    },
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 1
                    },
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 7
                    },
                })
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var mergedForm = new MergedForm()
            {
                FormNormalizedNumber = normalizedNumber
            };
            var args = new object[] { mergedForm };
            var retVal = (int)po.Invoke("GetNumberOfFormInstancesRequired", args);
            Assert.AreEqual(1, retVal);
        }

        /// <summary>
        /// Determines whether 1 instance of the form is required if the more questions are assigned
        /// than the MaximumMultipleRowCount
        /// </summary>
        [TestMethod]
        public void CanGetNumberOfFormInstancesRequiredWithNumberOfQuestionsGreaterThanMaximum()
        {
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
                Questions = new List<Question>(new[]
                {
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 5
                    },
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 1
                    },
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 7
                    },
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 2
                    },
                })
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var mergedForm = new MergedForm()
            {
                FormNormalizedNumber = normalizedNumber
            };
            var args = new object[] { mergedForm };
            var retVal = (int)po.Invoke("GetNumberOfFormInstancesRequired", args);
            Assert.AreEqual(2, retVal);
        }

        #endregion

        #region ReplaceSingleInstanceDocumentQuestionMergeFields

        /// <summary>
        /// Determines whether 1 instance of the form is required if no questions
        /// and no maximum multiple row count are assigned.
        /// </summary>
        [TestMethod]
        public void CanReplaceSingleInstanceDocumentQuestionMergeFieldsWithOnlyMultipleQuestionsDoesNothing()
        {
            Dictionary<string, Bookmark> mockedBookmarks = new Dictionary<string, Bookmark>();
            Dictionary<string, int> calledBookmarks = new Dictionary<string, int>();
            Dictionary<string, string> valueBookmarks = new Dictionary<string, string>();
            var normalizedNumber = "TestForm";
            var doc = new pol.Document()
            {
                NormalizedNumber = normalizedNumber,
                Questions = new List<Question>(new[]
                {
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 5,
                        MergeFieldName = "MergeField_1"
                    },
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 1,
                        MergeFieldName = "MergeField_1"
                    },
                    new Question()
                    {
                        Code = "Test1",
                        MaximumMultipleRowCount = 3,
                        MultipleRowGroupingNumber = 7,
                        MergeFieldName = "MergeField_1"
                    }
                })
            };

            var policy = new Policy()
            {
                RollUp = new RollUp()
                {
                    Documents = new List<pol.Document>(new[] { doc })
                }
            };

            var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
            var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

            var policyDocumentMock = new Mock<IPolicyDocumentManager>();
            policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>()))
                .Returns((string bookmarkName) =>
                {
                    if (!mockedBookmarks.ContainsKey(bookmarkName))
                    {
                        var bookmark = MockBookmark(bookmarkName);
                        mockedBookmarks.Add(bookmarkName, bookmark);
                    }

                    return mockedBookmarks[bookmarkName];
                });
            policyDocumentMock.Setup(s => s.ReplaceNodeValue(It.IsAny<Bookmark>(), It.IsAny<string>()))
                .Callback((Bookmark b, string s) =>
                {
                    if (!calledBookmarks.ContainsKey(b.Name))
                    {
                        calledBookmarks.Add(b.Name, 0);
                        valueBookmarks.Add(b.Name, string.Empty);
                    }

                    calledBookmarks[b.Name]++;
                    valueBookmarks[b.Name] = s;
                });
            policyDocumentMock.Setup(s => s.ReplaceQuestionBookmarkValue(It.IsAny<string>(), It.IsAny<Question>()))
                .Callback((string s, Question q) =>
                {
                    if (!calledBookmarks.ContainsKey(s))
                    {
                        calledBookmarks.Add(s, 0);
                        valueBookmarks.Add(s, string.Empty);
                    }

                    calledBookmarks[s]++;
                    valueBookmarks[s] = s;
                });

            var args = new object[] { policyDocumentMock.Object, normalizedNumber };
            po.Invoke("ReplaceSingleInstanceDocumentQuestionMergeFields", args);
            Assert.AreEqual(0, mockedBookmarks.Count);
            Assert.AreEqual(0, calledBookmarks.Count);
            Assert.AreEqual(0, valueBookmarks.Count);
        }

        /////// <summary>
        /////// Determines whether 1 instance of the form is required if no questions
        /////// and no maximum multiple row count are assigned.
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceSingleInstanceDocumentQuestionMergeFields()
        ////{
        ////    Dictionary<string, Bookmark> mockedBookmarks = new Dictionary<string, Bookmark>();
        ////    Dictionary<string, int> calledBookmarks = new Dictionary<string, int>();
        ////    Dictionary<string, string> valueBookmarks = new Dictionary<string, string>();
        ////    var normalizedNumber = "TestForm";
        ////    var doc = new pol.Document()
        ////    {
        ////        NormalizedNumber = normalizedNumber,
        ////        Questions = new List<Question>(new[]
        ////        {
        ////            CreateQuestion(DisplayFormat.TextBox, "Test1", "This is text box question 1", "MergeField_1", answerValue: "Answer1"),
        ////            CreateQuestion(DisplayFormat.TextBox, "Test2", "This is text box question 2", "MergeField_2", answerValue: "Answer2"),
        ////            CreateQuestion(DisplayFormat.TextBox, "Test3", "This is text box question 3", "MergeField_3", answerValue: "Answer3"),
        ////            CreateQuestion(
        ////                DisplayFormat.DropDown,
        ////                "Test4",
        ////                "This is a drop down question 4",
        ////                answerValue: "Other",
        ////                answers: new List<Answer>(new[] { CreateAnswer("MergeField_4", "Something"), CreateAnswer("MergeField_5", "Other", true) }))
        ////        })
        ////    };

        ////    var policy = new Policy()
        ////    {
        ////        RollUp = new RollUp()
        ////        {
        ////            Documents = new List<pol.Document>(new[] { doc })
        ////        }
        ////    };

        ////    var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
        ////    var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>()))
        ////        .Returns((string bookmarkName) =>
        ////        {
        ////            if (!mockedBookmarks.ContainsKey(bookmarkName))
        ////            {
        ////                var bookmark = MockBookmark(bookmarkName);
        ////                mockedBookmarks.Add(bookmarkName, bookmark);
        ////            }

        ////            return mockedBookmarks[bookmarkName];
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(It.IsAny<Bookmark>(), It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(b.Name))
        ////            {
        ////                calledBookmarks.Add(b.Name, 0);
        ////                valueBookmarks.Add(b.Name, string.Empty);
        ////            }

        ////            calledBookmarks[b.Name]++;
        ////            valueBookmarks[b.Name] = s;
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceQuestionBookmarkValue(It.IsAny<string>(), It.IsAny<Question>()))
        ////        .Callback((string s, Question q) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(s))
        ////            {
        ////                calledBookmarks.Add(s, 0);
        ////                valueBookmarks.Add(s, string.Empty);
        ////            }

        ////            calledBookmarks[s]++;
        ////            valueBookmarks[s] = q.AnswerValue;
        ////        });

        ////    var args = new object[] { policyDocumentMock.Object, normalizedNumber };
        ////    po.Invoke("ReplaceSingleInstanceDocumentQuestionMergeFields", args);
        ////    ////Assert.AreEqual(5, mockedBookmarks.Count);
        ////    Assert.AreEqual(5, calledBookmarks.Count);
        ////    Assert.AreEqual(5, valueBookmarks.Count);
        ////    Assert.AreEqual("Answer1", valueBookmarks["MergeField_1"]);
        ////    Assert.AreEqual("Answer2", valueBookmarks["MergeField_2"]);
        ////    Assert.AreEqual("Answer3", valueBookmarks["MergeField_3"]);
        ////    Assert.AreEqual("Other", valueBookmarks["MergeField_5"]);
        ////    Assert.AreEqual(1, calledBookmarks["MergeField_1"]);
        ////    Assert.AreEqual(1, calledBookmarks["MergeField_2"]);
        ////    Assert.AreEqual(1, calledBookmarks["MergeField_3"]);
        ////    Assert.AreEqual(1, calledBookmarks["MergeField_5"]);
        ////}

        #endregion

        #region ReplaceDocumentQuestionMergeFields

        /////// <summary>
        /////// Test whether replacement of a repeated bookmark gets filled correctly.
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceDocumentQuestionMergeFieldsWithMultipleQuestionsSameBookmarkNameUpdatesBookmarkSingleTime()
        ////{
        ////    Dictionary<string, Bookmark> mockedBookmarks = new Dictionary<string, Bookmark>();
        ////    Dictionary<string, int> calledBookmarks = new Dictionary<string, int>();
        ////    Dictionary<string, string> valueBookmarks = new Dictionary<string, string>();
        ////    var normalizedNumber = "TestForm";
        ////    var doc = new pol.Document()
        ////    {
        ////        NormalizedNumber = normalizedNumber,
        ////        Questions = new List<Question>(new[]
        ////        {
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 5,
        ////                MergeFieldName = "MergeField_1",
        ////                AnswerValue = "Answer 5"
        ////            },
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 1,
        ////                MergeFieldName = "MergeField_1",
        ////                AnswerValue = "Answer 1"
        ////            },
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 7,
        ////                MergeFieldName = "MergeField_1",
        ////                AnswerValue = "Answer 7"
        ////            }
        ////        })
        ////    };

        ////    var policy = new Policy()
        ////    {
        ////        RollUp = new RollUp()
        ////        {
        ////            Documents = new List<pol.Document>(new[] { doc })
        ////        }
        ////    };

        ////    var wordDoc = new Mock<Aspose.Words.Document>();
        ////    var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
        ////    var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>()))
        ////        .Returns((string bookmarkName) =>
        ////        {
        ////            if (!mockedBookmarks.ContainsKey(bookmarkName))
        ////            {
        ////                var bookmark = MockBookmark(bookmarkName);
        ////                mockedBookmarks.Add(bookmarkName, bookmark);
        ////            }

        ////            return mockedBookmarks[bookmarkName];
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(It.IsAny<Bookmark>(), It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(b.Name))
        ////            {
        ////                calledBookmarks.Add(b.Name, 0);
        ////                valueBookmarks.Add(b.Name, string.Empty);
        ////            }

        ////            calledBookmarks[b.Name]++;
        ////            valueBookmarks[b.Name] = s;
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceQuestionBookmarkValue(It.IsAny<string>(), It.IsAny<Question>()))
        ////        .Callback((string s, Question q) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(s))
        ////            {
        ////                calledBookmarks.Add(s, 0);
        ////                valueBookmarks.Add(s, string.Empty);
        ////            }

        ////            calledBookmarks[s]++;
        ////            valueBookmarks[s] = s;
        ////        });
        ////policyDocumentMock.Setup(s => s.Document)
        ////        .Returns(wordDoc.Object);

        ////    var args = new object[] { policyDocumentMock.Object, normalizedNumber, this.CustomMappings, 1 };
        ////    po.Invoke("ReplaceDocumentQuestionMergeFields", args);
        ////    ////Assert.AreEqual(1, mockedBookmarks.Count);
        ////    Assert.AreEqual(1, calledBookmarks.Count);
        ////    Assert.AreEqual(1, calledBookmarks["MergeField_1"]);
        ////    Assert.AreEqual(1, valueBookmarks.Count);
        ////}

        /////// <summary>
        /////// Test whether replacement of a repeated bookmark gets filled correctly.
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceDocumentQuestionMergeFieldsWithOnlyMultipleQuestionsUpdatesBookmarkMultipleTimes()
        ////{
        ////    Dictionary<string, Bookmark> mockedBookmarks = new Dictionary<string, Bookmark>();
        ////    Dictionary<string, int> calledBookmarks = new Dictionary<string, int>();
        ////    Dictionary<string, string> valueBookmarks = new Dictionary<string, string>();
        ////    var normalizedNumber = "TestForm";
        ////    var doc = new pol.Document()
        ////    {
        ////        NormalizedNumber = normalizedNumber,
        ////        Questions = new List<Question>(new[]
        ////        {
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 5,
        ////                MergeFieldName = "MergeField_1,MergeField_2,MergeField3",
        ////                AnswerValue = "Answer 5"
        ////            },
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 1,
        ////                MergeFieldName = "MergeField_1,MergeField_2,MergeField3",
        ////                AnswerValue = "Answer 1"
        ////            },
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 7,
        ////                MergeFieldName = "MergeField_1,MergeField_2,MergeField3",
        ////                AnswerValue = "Answer 7"
        ////            }
        ////        })
        ////    };

        ////    var policy = new Policy()
        ////    {
        ////        RollUp = new RollUp()
        ////        {
        ////            Documents = new List<pol.Document>(new[] { doc })
        ////        }
        ////    };

        ////    var wordDoc = new Mock<Aspose.Words.Document>();
        ////    var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
        ////    var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>()))
        ////        .Returns((string bookmarkName) =>
        ////        {
        ////            if (!mockedBookmarks.ContainsKey(bookmarkName))
        ////            {
        ////                var bookmark = MockBookmark(bookmarkName);
        ////                mockedBookmarks.Add(bookmarkName, bookmark);
        ////            }

        ////            return mockedBookmarks[bookmarkName];
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(It.IsAny<Bookmark>(), It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(b.Name))
        ////            {
        ////                calledBookmarks.Add(b.Name, 0);
        ////                valueBookmarks.Add(b.Name, string.Empty);
        ////            }

        ////            calledBookmarks[b.Name]++;
        ////            valueBookmarks[b.Name] = s;
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceQuestionBookmarkValue(It.IsAny<string>(), It.IsAny<Question>()))
        ////        .Callback((string s, Question q) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(s))
        ////            {
        ////                calledBookmarks.Add(s, 0);
        ////                valueBookmarks.Add(s, string.Empty);
        ////            }

        ////            calledBookmarks[s]++;
        ////            valueBookmarks[s] = s;
        ////        });
        ////    policyDocumentMock.Setup(s => s.Document)
        ////        .Returns(wordDoc.Object);

        ////    var args = new object[] { policyDocumentMock.Object, normalizedNumber, this.CustomMappings, 1 };
        ////    po.Invoke("ReplaceDocumentQuestionMergeFields", args);
        ////    ////Assert.AreEqual(3, mockedBookmarks.Count);
        ////    Assert.AreEqual(3, calledBookmarks.Count);
        ////    Assert.AreEqual(1, calledBookmarks["MergeField_1"]);
        ////    Assert.AreEqual(3, valueBookmarks.Count);
        ////}

        /////// <summary>
        /////// Test whether replacement of a repeated bookmark gets filled correctly.
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceDocumentQuestionMergeFieldsWithOnlyMultipleInstancesUpdatesBookmarkMultipleTimes()
        ////{
        ////    Dictionary<string, Bookmark> mockedBookmarks = new Dictionary<string, Bookmark>();
        ////    Dictionary<string, int> calledBookmarks = new Dictionary<string, int>();
        ////    Dictionary<string, string> valueBookmarks = new Dictionary<string, string>();
        ////    var normalizedNumber = "TestForm";
        ////    var doc = new pol.Document()
        ////    {
        ////        NormalizedNumber = normalizedNumber,
        ////        Questions = new List<Question>(new[]
        ////        {
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 5,
        ////                MergeFieldName = "MergeField_1,MergeField_2,MergeField3",
        ////                AnswerValue = "Answer 5"
        ////            },
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 1,
        ////                MergeFieldName = "MergeField_1,MergeField_2,MergeField3",
        ////                AnswerValue = "Answer 1"
        ////            },
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 7,
        ////                MergeFieldName = "MergeField_1,MergeField_2,MergeField3",
        ////                AnswerValue = "Answer 7"
        ////            },
        ////            new Question()
        ////            {
        ////                Code = "Test1",
        ////                MaximumMultipleRowCount = 3,
        ////                MultipleRowGroupingNumber = 2,
        ////                MergeFieldName = "MergeField_1,MergeField_2,MergeField3",
        ////                AnswerValue = "Answer 2"
        ////            }
        ////        })
        ////    };

        ////    var policy = new Policy()
        ////    {
        ////        RollUp = new RollUp()
        ////        {
        ////            Documents = new List<pol.Document>(new[] { doc })
        ////        }
        ////    };

        ////    var wordDoc = new Mock<Aspose.Words.Document>();
        ////    var manager = new IssuanceManager(this.Container.GetInstance<IRestHelper>(), policy, Carrier, this.Container.GetInstance<ITemplateProcessor>(), this.Container.GetInstance<IStorageManager>());
        ////    var po = new PrivateObject(manager, new PrivateType(typeof(IssuanceManager)));

        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>()))
        ////        .Returns((string bookmarkName) =>
        ////        {
        ////            if (!mockedBookmarks.ContainsKey(bookmarkName))
        ////            {
        ////                var bookmark = MockBookmark(bookmarkName);
        ////                mockedBookmarks.Add(bookmarkName, bookmark);
        ////            }

        ////            return mockedBookmarks[bookmarkName];
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(It.IsAny<Bookmark>(), It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(b.Name))
        ////            {
        ////                calledBookmarks.Add(b.Name, 0);
        ////                valueBookmarks.Add(b.Name, string.Empty);
        ////            }

        ////            calledBookmarks[b.Name]++;
        ////            valueBookmarks[b.Name] = s;
        ////        });
        ////    policyDocumentMock.Setup(s => s.ReplaceQuestionBookmarkValue(It.IsAny<string>(), It.IsAny<Question>()))
        ////        .Callback((string s, Question q) =>
        ////        {
        ////            if (!calledBookmarks.ContainsKey(s))
        ////            {
        ////                calledBookmarks.Add(s, 0);
        ////                valueBookmarks.Add(s, string.Empty);
        ////            }

        ////            calledBookmarks[s]++;
        ////            valueBookmarks[s] = s;
        ////        });
        ////    policyDocumentMock.Setup(s => s.Document)
        ////        .Returns(wordDoc.Object);

        ////    var args = new object[] { policyDocumentMock.Object, normalizedNumber, this.CustomMappings, 2 };
        ////    po.Invoke("ReplaceDocumentQuestionMergeFields", args);
        ////    ////Assert.AreEqual(3, mockedBookmarks.Count);
        ////    Assert.AreEqual(3, calledBookmarks.Count);
        ////    Assert.AreEqual(1, calledBookmarks["MergeField_1"]);
        ////    Assert.AreEqual(3, valueBookmarks.Count);
        ////}

        #endregion

        /// <summary>
        /// Determines whether this instance [can run test].
        /// </summary>
        public void CanRunInlandMarineTest()
        {
            var rest = this.Container.GetInstance<IRestHelper>();
            Assert.IsNotNull(rest);

            var template = this.Container.GetInstance<ITemplateProcessor>();
            Assert.IsNotNull(template);
            ////template.CreateCustomDictionary(customMapping);

            var storage = this.Container.GetInstance<IStorageManager>();
            Assert.IsNotNull(storage);

            IPolicyDocumentManager document = new PolicyDocumentManagerStub();
            Assert.IsNotNull(document);

            Policy jsonPolicy = CreateInlandMarinePolicy();
            var manager = new IssuanceManager(rest, jsonPolicy, Carrier, template, storage);
            Assert.IsNotNull(manager);
            ////manager.ApplyCustomMapping(ref document, this.CustomMappings);

            var result = document.Document.Range.Bookmarks[0].Text;
            Assert.IsNotNull(result);

            var response = manager.GenerateConsolidatedPolicyDocument();
            Assert.IsNotNull(manager);
            ////string val = this.templateProcessor.GetFieldValueFromCustom(customFieldMapping.Key, this.Policy, this.PolicyJson, ref docManager);
        }

        /// <summary>
        /// Creates a question.
        /// </summary>
        /// <param name="displayFormat">The display format.</param>
        /// <param name="code">The code.</param>
        /// <param name="verbiage">The verbiage.</param>
        /// <param name="mergeFieldName">Name of the merge field.</param>
        /// <param name="groupingNumber">The grouping number.</param>
        /// <param name="maxMultipleRowCount">The maximum number of rows allowed on the doc before a new instance is created</param>
        /// <param name="answerValue">The answer value.</param>
        /// <param name="controllingCode">The controlling code.</param>
        /// <param name="answerCodeReplacement">The answer code replacement.</param>
        /// <param name="answers">The answers.</param>
        /// <returns>Question.</returns>
        internal static Question CreateQuestion(
            DisplayFormat displayFormat,
            string code,
            string verbiage,
            string mergeFieldName = null,
            int groupingNumber = 0,
            int maxMultipleRowCount = 0,
            string answerValue = null,
            string controllingCode = null,
            string answerCodeReplacement = null,
            List<Answer> answers = null)
        {
            var question = new Question()
            {
                DisplayFormat = displayFormat,
                Code = code,
                Verbiage = verbiage,
                MergeFieldName = mergeFieldName,
                MultipleRowGroupingNumber = groupingNumber,
                AnswerValue = answerValue,
                ControllingQuestionCode = controllingCode,
                AnswerCodeReplacement = answerCodeReplacement,
                Answers = answers,
                MaximumMultipleRowCount = maxMultipleRowCount,

                CustomAnswersMethod = null,
                Column = 1,
                Display = true,
                Indention = 0,
                IsAllowMultipleRows = false,
                IsHiddenOnBindLetter = false,
                IsHiddenOnQuoteLetter = false,
                IsIssuanceOptional = false,
                IsReadOnly = false,
                IsRequired = true,
                Messages = null,
                Order = 1,
                OverrideEligibility = Eligibility.Agent,
                QuestionTypeId = null,
                Row = 1,
                Span = 1,
                ValidationMessage = null,
                ValidationRegex = null,
            };
            return question;
        }

        /// <summary>
        /// Creates an answer.
        /// </summary>
        /// <param name="mergeFieldName">Name of the merge field.</param>
        /// <param name="value">The value.</param>
        /// <param name="isSelected">if set to <c>true</c> [is selected].</param>
        /// <returns>Answer.</returns>
        internal static Answer CreateAnswer(string mergeFieldName, string value, bool isSelected = false)
        {
            return new Answer()
            {
                OverrideEligibility = Eligibility.Agent,
                IsDefault = false,
                IsSelected = isSelected,
                MergeFieldName = mergeFieldName,
                Order = 1,
                Value = value,
                Verbiage = value
            };
        }

        /// <summary>
        /// Mocks the bookmark.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns>Bookmark.</returns>
        internal static Bookmark MockBookmark(string bookmarkName)
        {
            var documentMock = new Mock<Aspose.Words.Document>();
            var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
            var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
            var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
            var bookmark = (Bookmark)bookmarkMock.Target;
            return bookmark;
        }

        /// <summary>
        /// Sets the authenticated user.
        /// </summary>
        private static void SetAuthenticatedUser()
        {
            AuthenticatedUser = new User()
            {
                AgencyCode = string.Empty,
                AgencyName = string.Empty,
                FirstName = "Test",
                IsBrokerageUser = false,
                IsInternalUser = true,
                LastName = "User",
                Region = "SE",
                Username = "testuser@markeltestuser.com"
            };
        }

        /// <summary>
        /// Creates the policy.
        /// </summary>
        /// <returns>Policy.</returns>
        private static Policy CreateInlandMarinePolicy()
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

            return policy;
        }

        /// <summary>
        /// Configures the container.
        /// </summary>
        private void ConfigureContainer()
        {
            this.Container = new Container();

            this.Container.Options.AllowOverridingRegistrations = true;
            this.Container.Register<IBootstrapper, Bootstrapper>();
            this.Container.Register<IJsonManager, JsonManager>();
            this.Container.Register<IRestHelper, RestHelperStub>();
            this.Container.Register<IStorageManager, StorageManagerStub>();
            this.Container.Register<ITemplateProcessor, TemplateProcessor>();
            this.Container.RegisterSingleton<ILog>(() => new NoOpLogger());

            this.Container.Verify();

            ////GlobalConfiguration.Configuration.DependencyResolver = new SimpleInjectorWebApiDependencyResolver(this.Container);
        }
    }
}