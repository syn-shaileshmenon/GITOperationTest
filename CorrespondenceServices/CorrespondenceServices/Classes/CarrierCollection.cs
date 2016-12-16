// ***********************************************************************
// Assembly         : CorrespondenceServices
// Author           : rsteelea
// Created          : 08-05-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-05-2016
// ***********************************************************************
// <copyright file="CarrierCollection.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Classes
{
    using System.Collections.Generic;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.RestHelper.Enumerations;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using RestSharp;
    using WebGrease.Css.Extensions;

    /// <summary>
    /// Class CarrierCollection.
    /// </summary>
    public class CarrierCollection
    {
        /// <summary>
        /// The carrier lock
        /// </summary>
        private static object carrierLock = new object();

        /// <summary>
        /// The carriers
        /// </summary>
        private static List<Carrier> carriers = null;

        /// <summary>
        /// Retrieves the Carriers using the specified rest helper.
        /// </summary>
        /// <param name="restHelper">The rest helper.</param>
        /// <param name="user">The user.</param>
        /// <returns>List&lt;Carrier&gt;.</returns>
        public static List<Carrier> Carriers(IRestHelper restHelper, IUser user)
        {
            lock (carrierLock)
            {
                if (CarrierCollection.carriers == null)
                {
                    var responses = restHelper.GetServiceResult(
                        "Carrier",
                        Method.GET,
                        "Collection",
                        user);
                    CarrierCollection.carriers = new List<Carrier>();

                    responses.ForEach(item =>
                        carriers.Add(new Carrier
                        {
                            AddressLine1 = (item["AddressLine1"] ?? string.Empty).ToString(),
                            AddressLine2 = (item["AddressLine2"] ?? string.Empty).ToString(),
                            City = (item["City"] ?? string.Empty).ToString(),
                            Code = (item["Code"] ?? string.Empty).ToString(),
                            Name = (item["Name"] ?? string.Empty).ToString(),
                            ZipCode = (item["ZipCode"] ?? string.Empty).ToString(),
                            State = (item["State"] ?? string.Empty).ToString(),
                        }));
                }
            }

            return CarrierCollection.carriers;
        }
    }
}
