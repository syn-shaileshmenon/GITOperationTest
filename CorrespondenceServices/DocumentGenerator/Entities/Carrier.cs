// <copyright file="Carrier.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Entities
{
    /// <summary>
    /// Class Carrier
    /// </summary>
    public class Carrier
    {
        /// <summary>
        /// Gets or sets AddressLine1
        /// </summary>
        public string AddressLine1 { get; set; }

        /// <summary>
        /// Gets or sets AddressLine2
        /// </summary>
        public string AddressLine2 { get; set; }

        /// <summary>
        /// Gets or sets city
        /// </summary>
        public string City { get; set; }

        /// <summary>
        /// Gets or sets Code
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether it is Markel
        /// </summary>
        public bool IsMarkel { get; set; }

        /// <summary>
        /// Gets or sets Name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets Short name
        /// </summary>
        public string ShortName { get; set; }

        /// <summary>
        /// Gets or sets state
        /// </summary>
        public string State { get; set; }

        /// <summary>
        /// Gets or sets ZipCode
        /// </summary>
        public string ZipCode { get; set; }
    }
}
