// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 10-31-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-01-2016
// ***********************************************************************
// <copyright file="CarrierCollectionTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Tests
{
    using System.Diagnostics.CodeAnalysis;
    using Castle.Core.Internal;
    using CorrespondenceServices.Classes;
    using CorrespondenceServices.Tests.Stubs;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.RestHelper.Classes;
    using Mkl.WebTeam.RestHelper.Enumerations;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Moq;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using RestSharp;

    /// <summary>
    /// Class CarrierCollectionTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class CarrierCollectionTests
    {
        /// <summary>
        /// Tests the method1.
        /// </summary>
        [TestMethod]
        public void TestMethod1()
        {
            var list = RestHelperStub.GetFakeCarrierList();
            var json = "{\"Collection\":" + JsonConvert.SerializeObject(list) + "}";
            var token = JToken.Parse(json);
            var ienum = token["Collection"].Children();

            var mockUser = new Mock<User>();
            var mockRestHelper = new Mock<IRestHelper>();
            mockRestHelper.Setup(r => r.GetServiceResult("Carrier", Method.GET, "Collection", It.IsAny<IUser>(), null, WebTeamServiceEndpoint.DeliveryServices, null))
                .Returns(ienum);

            var carriers = CarrierCollection.Carriers(mockRestHelper.Object, mockUser.Object);
            mockRestHelper.Verify(v => v.GetServiceResult(It.IsAny<string>(), It.IsAny<Method>(), It.IsAny<IUser>(), It.IsAny<object>(), It.IsAny<WebTeamServiceEndpoint>(), It.IsAny<string>()), Times.Never);
            mockRestHelper.Verify(v => v.GetServiceResult(It.IsAny<string>(), It.IsAny<Method>(), It.IsAny<string>(), It.IsAny<IUser>(), It.IsAny<object>(), It.IsAny<WebTeamServiceEndpoint>(), It.IsAny<string>()), Times.AtLeastOnce);
            Assert.IsFalse(carriers.IsNullOrEmpty());
            Assert.AreEqual(2, carriers.Count);
        }
    }
}
