// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : rsteelea
// Created          : 11-22-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-22-2016
// ***********************************************************************
// <copyright file="CustomFunctions_CP1218.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System;
    using Aspose.Words;
    using DecisionModel.Models.Policy;

    /// <summary>
    /// Class CustomFunctions.
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Formats the c P1218.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="manager">The manager.</param>
        /// <param name="question">The question.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="policy">The policy.</param>
        /// <returns>System.String.</returns>
        public string Format_CP1218(string value, IPolicyDocumentManager manager, Question question, Bookmark bookmark, string bookmarkName, Policy policy)
        {
            if (string.Compare(bookmarkName, BlankMergeField + "2", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bookmarkName, BlankMergeField + "8", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bookmarkName, BlankMergeField + "14", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                return value.Substring(0, 3);
            }

            return value;
        }
    }
}
