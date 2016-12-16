// <copyright file="CustomFunctions_MEGL1522.cs" company="Markel">
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
    /// Class CustomFunctions_MEGL1522
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Formats the MEGL 1522.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="manager">The manager.</param>
        /// <param name="question">The question.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="policy">The policy.</param>
        /// <returns>System.String.</returns>
        public string Format_MEGL1522(string value, IPolicyDocumentManager manager, Question question, Bookmark bookmark, string bookmarkName, Policy policy)
        {
            if (string.Compare(bookmarkName, "BlankMergeField", StringComparison.InvariantCultureIgnoreCase) == 0 ||
                string.Compare(bookmarkName, BlankMergeField + "1", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                return string.Format(FormatHelper.CurrencyFormat, Convert.ToDecimal(value));
            }

            return value;
        }
    }
}
