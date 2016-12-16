// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 11-01-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-01-2016
// ***********************************************************************
// <copyright file="DocumentSearchRepresentationTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CorrespondenceServices.Tests
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using CorrespondenceServices.Classes;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Class DocumentSearchRepresentationTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class DocumentSearchRepresentationTests
    {
        /// <summary>
        /// Determines whether this instance [can test document search representation initialization].
        /// </summary>
        [TestMethod]
        public void CanTestDocumentSearchRepresentationInitialization()
        {
            var rep = new DocumentSearchRepresentation();
            Assert.IsNotNull(rep);
            Assert.IsNull(rep.EditionDate);
            Assert.IsNull(rep.EffectiveOnlyDate);
            Assert.IsNull(rep.FormNumber);
        }

        /// <summary>
        /// Determines whether this instance [can test document search representation form number get set].
        /// </summary>
        [TestMethod]
        public void CanTestDocumentSearchRepresentationFormNumberGetSet()
        {
            const string FormNumber = "FormNumber";
            var rep = new DocumentSearchRepresentation();
            Assert.IsNotNull(rep);
            rep.FormNumber = FormNumber;
            Assert.AreEqual(FormNumber, rep.FormNumber);
        }

        /// <summary>
        /// Determines whether this instance [can test document search representation edition date get set].
        /// </summary>
        [TestMethod]
        public void CanTestDocumentSearchRepresentationEditionDateGetSet()
        {
            var editionDate = DateTime.Now;
            var rep = new DocumentSearchRepresentation();
            Assert.IsNotNull(rep);
            rep.EditionDate = editionDate;
            Assert.IsTrue(rep.EditionDate.HasValue);
            Assert.AreEqual(editionDate, rep.EditionDate.Value);

            rep.EditionDate = null;
            Assert.IsFalse(rep.EditionDate.HasValue);
            Assert.IsNull(rep.EditionDate);
        }

        /// <summary>
        /// Determines whether this instance [can test document search representation effective date get set].
        /// </summary>
        [TestMethod]
        public void CanTestDocumentSearchRepresentationEffectiveDateGetSet()
        {
            var effectiveDate = DateTime.Now;
            var rep = new DocumentSearchRepresentation();
            Assert.IsNotNull(rep);
            rep.EffectiveOnlyDate = effectiveDate;
            Assert.IsTrue(rep.EffectiveOnlyDate.HasValue);
            Assert.AreEqual(effectiveDate, rep.EffectiveOnlyDate.Value);

            rep.EffectiveOnlyDate = null;
            Assert.IsFalse(rep.EffectiveOnlyDate.HasValue);
            Assert.IsNull(rep.EffectiveOnlyDate);
        }
    }
}
