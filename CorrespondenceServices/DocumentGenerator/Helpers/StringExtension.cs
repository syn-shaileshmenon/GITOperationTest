// <copyright file="StringExtension.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Web;

    /// <summary>
    /// Reflection class to return objects base on string passed in
    /// </summary>
    public static class StringExtension
    {
        /// <summary>
        /// The is minimum premium
        /// </summary>
        private static object isMinPremium = string.Empty;

        /// <summary>
        /// Gets or sets the is minimum premium.
        /// </summary>
        /// <value>
        /// The is minimum premium.
        /// </value>
        public static object IsMinPremium
        {
            get { return isMinPremium; }
            set { isMinPremium = value; }
        }

        /// <summary>
        /// Gets the property object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="name">The name.</param>
        /// <returns>object</returns>
        public static object GetPropObject(this object obj, string name)
        {
            foreach (var part in name.Split('.'))
            {
                if (obj == null)
                {
                    return null;
                }

                var type = obj.GetType();
                var info = type.GetProperty(part);
                if (info == null)
                {
                    return null;
                }

                obj = info.GetValue(obj, null);
            }

            return obj;
        }

        /// <summary>
        /// Gets the property value.
        /// </summary>
        /// <typeparam name="T">type of object expected</typeparam>
        /// <param name="obj">The object.</param>
        /// <param name="name">The name.</param>
        /// <returns>object</returns>
        public static T GetPropValue<T>(this object obj, string name)
        {
            object retval = GetPropValue(obj, name);
            if (retval == null)
            {
                return default(T);
            }

            if (retval.GetType().ToString() == "System.DateTime")
            {
                retval = Convert.ToDateTime(retval).ToShortDateString();
                return (T)retval;
            }

            retval = retval.ToString();

            // throws InvalidCastException if types are incompatible
            return (T)retval;
        }

        /// <summary>
        /// Compares a string with another ignoring case
        /// </summary>
        /// <param name="str">Source string</param>
        /// <param name="valueToCompare">String to compare with</param>
        /// <returns>True if string are equal, else False</returns>
        public static bool EqualsIgnoreCase(this string str, string valueToCompare)
        {
            return str.Equals(valueToCompare, StringComparison.InvariantCultureIgnoreCase);
        }

        /// <summary>
        /// Gets the property value.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="name">The name.</param>
        /// <returns>object</returns>
        private static object GetPropValue(this object obj, string name)
        {
            foreach (string part in name.Split('.'))
            {
                if (obj == null)
                {
                    return null;
                }

                isMinPremium = string.Empty;
                if (part == "Premium" && obj.GetType().GetProperty("IsMinimumPremium") != null)
                {
                    object flag = obj.GetType().GetProperty("IsMinimumPremium").GetValue(obj, null);
                    if ((bool)flag)
                    {
                        isMinPremium = "\t MP";
                    }
                }

                Type type = obj.GetType();
                if (part.EndsWith("()"))
                {
                    string func = part.Substring(0, part.Length - 2);
                    try
                    {
                        obj = type.GetMethods()?.FirstOrDefault(x => x.Name == func)?.Invoke(obj, new object[] { });
                    }
                    catch (Exception)
                    {
                        return null;
                    }
                }
                else
                {
                    PropertyInfo info = type.GetProperty(part);
                    if (info == null)
                    {
                        return null;
                    }

                    obj = info.GetValue(obj, null);
                }
            }

            return obj;
        }
    }
}