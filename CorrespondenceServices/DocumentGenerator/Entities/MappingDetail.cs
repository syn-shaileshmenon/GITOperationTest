// <copyright file="MappingDetail.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Entities
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Mapping detail class that is to be added to dictionary for further processing of any field
    /// </summary>
    public class MappingDetail : MergedField
    {
        /// <summary>
        /// Gets or sets a value indicating whether mapping detail has a function
        /// </summary>
        public bool HasFunction { get; set; }

        /// <summary>
        /// Gets or sets the function value
        /// </summary>
        [JsonProperty("function")]
        public string Function { get; set; }

        /// <summary>
        /// Gets or sets the function parameters - list of mapping detail
        /// </summary>
        [JsonProperty("params")]
        public List<MappingDetail> Params { get; set; }
    }
}
