// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 10-27-2016
//
// Last Modified By : rsteelea
// Last Modified On : 10-27-2016
// ***********************************************************************
// <copyright file="HelperTests.cs" company="Markel">
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
    /// Class HelperTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class HelperTests
    {
        /// <summary>
        /// Determines whether this instance [can is numeric validate int].
        /// </summary>
        [TestMethod]
        public void CanIsNumericValidateInt()
        {
            Assert.IsTrue(ValidationHelper.IsNumeric("1"));
        }

        /// <summary>
        /// Determines whether this instance [can is numeric validate negative int].
        /// </summary>
        [TestMethod]
        public void CanIsNumericValidateNegativeInt()
        {
            Assert.IsTrue(ValidationHelper.IsNumeric("-1"));
        }

        /// <summary>
        /// Determines whether this instance [can is numeric validate decimal].
        /// </summary>
        [TestMethod]
        public void CanIsNumericValidateDecimal()
        {
            Assert.IsTrue(ValidationHelper.IsNumeric("1.45"));
        }

        /// <summary>
        /// Determines whether this instance [can is numeric validate negative decimal].
        /// </summary>
        [TestMethod]
        public void CanIsNumericValidateNegativeDecimal()
        {
            Assert.IsTrue(ValidationHelper.IsNumeric("-8.34"));
        }

        /// <summary>
        /// Test cannot the is numeric validate string without number.
        /// </summary>
        [TestMethod]
        public void CannotIsNumericValidateStringWithoutNumber()
        {
            Assert.IsFalse(ValidationHelper.IsNumeric("Hello"));
        }

        /// <summary>
        /// Test cannot the is numeric validate string with number at end.
        /// </summary>
        [TestMethod]
        public void CannotIsNumericValidateStringWithNumberAtEnd()
        {
            Assert.IsFalse(ValidationHelper.IsNumeric("Hello1"));
        }

        /// <summary>
        /// Test cannot the is numeric validate string with number at beginning.
        /// </summary>
        [TestMethod]
        public void CannotIsNumericValidateStringWithNumberAtBeginning()
        {
            Assert.IsFalse(ValidationHelper.IsNumeric("1Hello"));
        }

        /// <summary>
        /// Test cannot the is numeric validate string with number in middle.
        /// </summary>
        [TestMethod]
        public void CannotIsNumericValidateStringWithNumberInMiddle()
        {
            Assert.IsFalse(ValidationHelper.IsNumeric("Hel1lo"));
        }

        /// <summary>
        /// Test cannot the is numeric validate empty string.
        /// </summary>
        [TestMethod]
        public void CannotIsNumericValidateEmptyString()
        {
            Assert.IsFalse(ValidationHelper.IsNumeric(string.Empty));
        }

        /// <summary>
        /// Nulls the is numeric throws exception.
        /// </summary>
        [TestMethod]
        public void CannotIsNumericValidateNull()
        {
            Assert.IsFalse(((string)null).IsNumeric());
        }
    }
}
