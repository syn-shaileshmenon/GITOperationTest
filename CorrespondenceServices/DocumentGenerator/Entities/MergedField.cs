// <copyright file="MergedField.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Entities
{
    using Newtonsoft.Json;

    /// <summary>
    /// Merged field detail
    /// </summary>
    public class MergedField
    {
        /// <summary>
        /// Gets or sets the merged field value
        /// </summary>
        [JsonProperty("field")]
        public string Field { get; set; }

        /// <summary>
        /// Gets or sets the constant of merged field
        /// </summary>
        [JsonProperty("constant")]
        public string Constant { get; set; }

        /// <summary>
        /// Gets or sets the image associated with field
        /// </summary>
        [JsonProperty("image")]
        public string Image { get; set; }

        /// <summary>
        /// Gets or sets the format used for the field
        /// </summary>
        [JsonProperty("format")]
        public string Format { get; set; }

        /// <summary>
        /// Gets or sets the jsonPath
        /// </summary>
        [JsonProperty("jsonPath")]
        public string JsonPath { get; set; }
    }
}
