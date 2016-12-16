// <copyright file="KeyComparer.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Helpers
{
    using System.Collections.Generic;
    using Aspose.Words;

    /// <summary>
    /// Class KeyComparer
    /// </summary>
    public class KeyComparer : IEqualityComparer<KeyValuePair<string, Node>>
    {
        /// <summary>
        /// Determines whether the specified objects are equal.
        /// </summary>
        /// <param name="x">First object to compare</param>
        /// <param name="y">Second object to compare</param>
        /// <returns>true if the specified objects are equal; otherwise, false.</returns>
        public bool Equals(KeyValuePair<string, Node> x, KeyValuePair<string, Node> y)
        {
            return x.Key.Equals(y.Key);
        }

        /// <summary>
        /// Returns a hash code for the specified object.
        /// </summary>
        /// <param name="obj">Object for which hash code is to be returned</param>
        /// <returns>A hash code for the specified object.</returns>
        public int GetHashCode(KeyValuePair<string, Node> obj)
        {
            return obj.GetHashCode();
        }
    }
}
