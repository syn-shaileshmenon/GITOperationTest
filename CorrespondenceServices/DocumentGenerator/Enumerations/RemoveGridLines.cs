// ***********************************************************************
// Assembly         :
// Author           : rsteelea
// Created          : 11-14-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-14-2016
// ***********************************************************************
// <copyright file="RemoveGridLines.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace Mkl.WebTeam.DocumentGenerator.Enumerations
{
    using System;

    /// <summary>
    /// Enum RemoveGridLines
    /// </summary>
    [Flags]
    public enum RemoveGridLines
    {
        /// <summary>
        /// The none
        /// </summary>
        None = 0,

        /// <summary>
        /// The top
        /// </summary>
        Top = 1,

        /// <summary>
        /// The bottom
        /// </summary>
        Bottom = 2,

        /// <summary>
        /// The top bottom
        /// </summary>
        TopBottom = Top + Bottom,

        /// <summary>
        /// The left
        /// </summary>
        Left = 4,

        /// <summary>
        /// The right
        /// </summary>
        Right = 8,

        /// <summary>
        /// The sides
        /// </summary>
        Sides = Left + Right,

        /// <summary>
        /// All
        /// </summary>
        All = Top + Bottom + Left + Right
    }
}
