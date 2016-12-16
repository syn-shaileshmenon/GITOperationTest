// <copyright file="DownloadParameters.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Classes.Parameters
{
    using System;
    using Mkl.WebTeam.DocumentGenerator;

    /// <summary>
    /// Parameters needed to maintain from Edit/Save method to Download method
    /// </summary>
    [Serializable]
    public class DownloadParameters
    {
        /// <summary>
        /// Gets or sets the save format.
        /// </summary>
        /// <value>
        /// The save format.
        /// </value>
        public DocumentManager.FileFormat SaveFormat { get; set; }

        /// <summary>
        /// Gets or sets the name of the temporary file.
        /// </summary>
        /// <value>
        /// The name of the temporary file.
        /// </value>
        public string TempFileName { get; set; }
    }
}