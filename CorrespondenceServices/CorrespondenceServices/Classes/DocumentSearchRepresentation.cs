// <copyright file="DocumentSearchRepresentation.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Classes
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web;

    /// <summary>
    /// Document search representation
    /// </summary>
    public class DocumentSearchRepresentation
    {
        /// <summary>
        /// Gets or sets the Effective date
        /// </summary>
        public DateTime? EffectiveOnlyDate { get; set; }

        /// <summary>
        /// Gets or sets the Edition date
        /// </summary>
        public DateTime? EditionDate { get; set; }

        /// <summary>
        /// Gets or sets the Form normalized number
        /// </summary>
        public string FormNumber { get; set; }
    }
}