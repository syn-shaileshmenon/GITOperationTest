// <copyright file="CustomFunctions_CP0450.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Aspose.Words;
    using Aspose.Words.Tables;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Helpers;

    /// <summary>
    /// Class CustomFunctions_CP0450
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Formats the CP 04 50.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="manager">The manager.</param>
        /// <param name="question">The question.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="policy">The policy.</param>
        /// <returns>System.String.</returns>
        public string Format_CP0450(string value, IPolicyDocumentManager manager, Question question, Bookmark bookmark, string bookmarkName, Policy policy)
        {
            if (string.Compare(bookmarkName, BlankMergeField + "2", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bookmarkName, BlankMergeField + "6", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bookmarkName, BlankMergeField + "10", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                return string.Format(FormatHelper.DateFormat, policy.EffectiveDate);
            }
            else if (string.Compare(bookmarkName, BlankMergeField + "3", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bookmarkName, BlankMergeField + "7", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bookmarkName, BlankMergeField + "11", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                return string.Format(FormatHelper.DateFormat, policy.ExpirationDate);
            }

            return value;
        }
    }
}
