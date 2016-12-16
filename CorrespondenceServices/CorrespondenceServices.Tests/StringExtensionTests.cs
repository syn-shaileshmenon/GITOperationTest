// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 11-02-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-02-2016
// ***********************************************************************
// <copyright file="StringExtensionTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CorrespondenceServices.Tests
{
    using System.Diagnostics.CodeAnalysis;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator.Helpers;

    /// <summary>
    /// Class StringExtensionTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class StringExtensionTests
    {
        /// <summary>
        /// The last value
        /// </summary>
        private static object lastValue = string.Empty;

        /// <summary>
        /// Determines whether this instance [can test is minimum premium property string get set].
        /// </summary>
        [TestMethod]
        public void CanTestIsMinPremiumPropertyStringGetSet()
        {
            const string MinPremium = "MinPremium";
            Assert.AreEqual(lastValue, StringExtension.IsMinPremium);
            this.SetIsMinPremium(MinPremium);
            Assert.AreEqual(MinPremium, StringExtension.IsMinPremium);
        }

        /// <summary>
        /// Determines whether this instance [can test is minimum premium property int get set].
        /// </summary>
        [TestMethod]
        public void CanTestIsMinPremiumPropertyIntGetSet()
        {
            const int MinPremium = 1234;
            Assert.AreEqual(lastValue, StringExtension.IsMinPremium);
            this.SetIsMinPremium(MinPremium);
            StringExtension.IsMinPremium = MinPremium;
            Assert.AreEqual(MinPremium, StringExtension.IsMinPremium);
        }

        /// <summary>
        /// Determines whether this instance [can test is minimum premium property object get set].
        /// </summary>
        [TestMethod]
        public void CanTestIsMinPremiumPropertyObjectGetSet()
        {
            object minPremium = new object();
            Assert.AreEqual(lastValue, StringExtension.IsMinPremium);
            this.SetIsMinPremium(minPremium);
            StringExtension.IsMinPremium = minPremium;
            Assert.AreEqual(minPremium, StringExtension.IsMinPremium);
        }

        /// <summary>
        /// Determines whether this instance [can test equals ignore case success].
        /// </summary>
        [TestMethod]
        public void CanTestEqualsIgnoreCaseSuccess()
        {
            var string1 = "SomeString";
            var string2 = "SomeString";
            var string3 = string1.ToLower();
            var string4 = string1.ToUpper();

            var result = string1.EqualsIgnoreCase(string2);
            Assert.IsTrue(result);

            result = string1.EqualsIgnoreCase(string1);
            Assert.IsTrue(result);

            result = string1.EqualsIgnoreCase(string3);
            Assert.IsTrue(result);

            result = string1.EqualsIgnoreCase(string4);
            Assert.IsTrue(result);
        }

        /// <summary>
        /// Determines whether this instance [can test equals ignore case fails].
        /// </summary>
        [TestMethod]
        public void CanTestEqualsIgnoreCaseFails()
        {
            var string1 = "SomeString";
            var string2 = "SomeString1";
            var string3 = "1";
            var string4 = string.Empty;

            var result = string1.EqualsIgnoreCase(string2);
            Assert.IsFalse(result);

            result = string1.EqualsIgnoreCase(string3);
            Assert.IsFalse(result);

            result = string1.EqualsIgnoreCase(string4);
            Assert.IsFalse(result);

            result = string1.EqualsIgnoreCase(null);
            Assert.IsFalse(result);
        }

        /// <summary>
        /// Sets the is minimum premium.
        /// </summary>
        /// <param name="value">The value.</param>
        private void SetIsMinPremium(object value)
        {
            StringExtension.IsMinPremium = value;
            lastValue = value;
        }
    }
}
