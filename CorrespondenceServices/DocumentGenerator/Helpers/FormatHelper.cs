// <copyright file="FormatHelper.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Helpers
{
    /// <summary>
    /// Format Helper
    /// </summary>
    public static class FormatHelper
    {
        /// <summary>
        /// Date format
        /// </summary>
        public const string DateFormat = "{0:MM/dd/yyyy}";

        /// <summary>
        /// Currency format
        /// </summary>
        public const string CurrencyFormat = "{0:n0}";

        /// <summary>
        /// Currency format with decimal
        /// </summary>
        public const string CurrencyFormatWithDecimal = "{0:n2}";

        /// <summary>
        /// Format Value
        /// </summary>
        /// <param name="input">input</param>
        /// <param name="format">Format Data Type</param>
        /// <returns>string</returns>
        public static string FormatValue(object input, string format)
        {
            string formatedValue = string.Empty;
            if (input != null)
            {
                formatedValue = string.IsNullOrWhiteSpace(format) ? input.ToString() : string.Format(format, input);
            }

            return formatedValue;
        }
    }
}