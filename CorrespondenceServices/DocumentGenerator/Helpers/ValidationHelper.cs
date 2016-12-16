// <copyright file="ValidationHelper.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Helpers
{
    /// <summary>
    /// Class ValidationHelper.
    /// </summary>
    public static class ValidationHelper
    {
        /// <summary>
        /// Determines whether the specified s is numeric.
        /// </summary>
        /// <param name="s">The s.</param>
        /// <returns>boolean</returns>
        public static bool IsNumeric(this string s)
        {
            double output;
            return double.TryParse(s, out output);
        }
    }
}
