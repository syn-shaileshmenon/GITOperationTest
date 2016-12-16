// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-17-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-17-2016
// ***********************************************************************
// <copyright file="EntityTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Tests
{
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator.Entities;

    /// <summary>
    /// Class EntityTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class EntityTests
    {
        /// <summary>
        /// Class CarrierTests.
        /// </summary>
        [TestClass]
        public class CarrierTests
        {
            /// <summary>
            /// Determines whether this instance [can set and get properties].
            /// </summary>
            [TestMethod]
            public void CanSetAndGetProperties()
            {
                var carrier = new Carrier();
                Assert.IsNotNull(carrier);

                carrier.AddressLine1 = "Address Line 1";
                carrier.AddressLine2 = "Address Line 2";
                carrier.City = "City";
                carrier.State = "State";
                carrier.Code = "Code";
                carrier.IsMarkel = true;
                carrier.Name = "Name";
                carrier.ShortName = "Short Name";
                carrier.ZipCode = "Zip Code";

                Assert.AreEqual("Address Line 1", carrier.AddressLine1);
                Assert.AreEqual("Address Line 2", carrier.AddressLine2);
                Assert.AreEqual("City", carrier.City);
                Assert.AreEqual("State", carrier.State);
                Assert.AreEqual("Code", carrier.Code);
                Assert.IsTrue(carrier.IsMarkel);
                Assert.AreEqual("Name", carrier.Name);
                Assert.AreEqual("Short Name", carrier.ShortName);
                Assert.AreEqual("Zip Code", carrier.ZipCode);
            }
        }

        /// <summary>
        /// Class MappingDetailTests.
        /// </summary>
        [TestClass]
        public class MappingDetailTests
        {
            /// <summary>
            /// Determines whether this instance [can set and get properties test].
            /// </summary>
            [TestMethod]
            public void CanSetAndGetPropertiesTest()
            {
                var mappingDetails = new MappingDetail();
                Assert.IsNotNull(mappingDetails);

                mappingDetails.Function = "Function";
                mappingDetails.HasFunction = true;

                Assert.IsNull(mappingDetails.Params);
                mappingDetails.Params = new List<MappingDetail>();

                Assert.AreEqual("Function", mappingDetails.Function);
                Assert.IsTrue(mappingDetails.HasFunction);
                Assert.IsNotNull(mappingDetails.Params);
            }
        }

        /// <summary>
        /// Class MergeFieldTests.
        /// </summary>
        [TestClass]
        public class MergeFieldTests
        {
            /// <summary>
            /// Determines whether this instance [can set and get properties test].
            /// </summary>
            [TestMethod]
            public void CanSetAndGetPropertiesTest()
            {
                var mergeField = new MergedField();
                Assert.IsNotNull(mergeField);

                mergeField.Constant = "Constant";
                mergeField.Field = "Field";
                mergeField.Format = "Format";
                mergeField.Image = "Image";
                mergeField.JsonPath = "JsonPath";

                Assert.AreEqual("Constant", mergeField.Constant);
                Assert.AreEqual("Field", mergeField.Field);
                Assert.AreEqual("Format", mergeField.Format);
                Assert.AreEqual("Image", mergeField.Image);
                Assert.AreEqual("JsonPath", mergeField.JsonPath);
            }
        }
    }
}
