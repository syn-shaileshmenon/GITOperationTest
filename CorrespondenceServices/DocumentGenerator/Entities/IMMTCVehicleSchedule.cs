// <copyright file="IMMTCVehicleSchedule.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Entities
{
    /// <summary>
    /// Class IMMTCVehicleSchedule
    /// </summary>
    public class IMMTCVehicleSchedule
    {
        /// <summary>
        /// Gets or sets Year
        /// </summary>
        public string Year { get; set; }

        /// <summary>
        /// Gets or sets Manufacturer
        /// </summary>
        public string Manufacturer { get; set; }

        /// <summary>
        /// Gets or sets Identification Number
        /// </summary>
        public string IdentificationNumber { get; set; }

        /// <summary>
        /// Gets or sets Body type
        /// </summary>
        public string BodyType { get; set; }

        /// <summary>
        /// Gets or sets Limit of liability per vehicle
        /// </summary>
        public decimal LimitOfLiabilityPerVehicle { get; set; }

        /// <summary>
        /// Gets or sets Premium per vehicle
        /// </summary>
        public decimal PremiumPerVehicle { get; set; }

        /// <summary>
        /// Gets or sets multiple row grouping number
        /// </summary>
        public int MultipleRowGroupingNumber { get; set; }
    }
}
