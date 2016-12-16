// <copyright file="IJsonManager.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator
{
    using System.Collections.Generic;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Interface to json manager
    /// </summary>
    public interface IJsonManager
    {
        /// <summary>
        /// GenerateMappingDictionary method generates a dictionary based on passed in JToken object
        /// </summary>
        /// <param name="jsonMappingTemplate">JToken</param>
        /// <returns>Dictionary</returns>
        Dictionary<string, MappingDetail> GenerateMappingDictionary(JToken jsonMappingTemplate);
    }
}
