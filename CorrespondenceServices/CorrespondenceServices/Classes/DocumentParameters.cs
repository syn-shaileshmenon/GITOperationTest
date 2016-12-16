// <copyright file="DocumentParameters.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Classes.Parameters
{
    using System.Collections.Generic;
    using Mkl.WebTeam.DocumentGenerator;

    /// <summary>
    /// Classes for Document Edit/Save/Download Parameters
    /// </summary>
    public class DocumentParameters
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentParameters"/> class.
        /// </summary>
        public DocumentParameters()
        {
            this.NamedParameters = new List<NamedValue>();
            this.FileName = string.Empty;
            this.FileFormat = DocumentManager.FileFormat.DOCX;
        }

        /// <summary>
        /// Gets or sets the file format.
        /// </summary>
        /// <value>
        /// The file format.
        /// </value>
        public DocumentManager.FileFormat FileFormat { get; set; }

        /// <summary>
        /// Gets or sets the name of the file.
        /// </summary>
        /// <value>
        /// The name of the file.
        /// </value>
        public string FileName { get; set; }

        /// <summary>
        /// Gets or sets the named parameters.
        /// </summary>
        /// <value>
        /// The named parameters.
        /// </value>
        public List<NamedValue> NamedParameters { get; set; }

        /// <summary>
        /// provides Named Value
        /// </summary>
        public class NamedValue
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="NamedValue"/> class.
            /// </summary>
            public NamedValue()
            {
            }

            /// <summary>
            /// Gets or sets the name.
            /// </summary>
            /// <value>
            /// The name.
            /// </value>
            public string Name { get; set; }

            /// <summary>
            /// Gets or sets the value.
            /// </summary>
            /// <value>
            /// The value.
            /// </value>
            public object Value { get; set; }
        }
    }
}