// <copyright file="ReflectionHelper.cs" company="Markel">
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
    /// Reflection class to return objects base on property name passed in
    /// </summary>
    public static class ReflectionHelper
    {
        /// <summary>
        /// Custom functions class
        /// </summary>
        private const string CustomFunctionClass = "Mkl.WebTeam.DocumentGenerator.Functions.CustomFunctions";

        /// <summary>
        /// Gets the property object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="name">The name.</param>
        /// <returns>object</returns>
        public static object GetPropObject(this object obj, string name)
        {
            foreach (string part in name.Split('.'))
            {
                if (obj == null)
                {
                    return null;
                }

                Type type = obj.GetType();
                PropertyInfo info = type.GetProperty(part);
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
        /// <param name="format">Format string</param>
        /// <returns>object</returns>
        public static T GetPropValue<T>(this object obj, string name, string format)
        {
            object retval = GetPropValue(obj, name);
            if (retval == null)
            {
                return default(T);
            }

            if (!string.IsNullOrWhiteSpace(format))
            {
                retval = FormatHelper.FormatValue(retval, format);
            }
            else
            {
                retval = retval.ToString();
            }

            // throws InvalidCastException if types are incompatible
            return (T)retval;
        }

        /// <summary>
        /// Invokes a method on the function library instance using named parameters
        /// </summary>
        /// <param name="methodName">MethodName to invoke</param>
        /// <param name="methodParams">Method Parameters to the invoked method - must match method signature</param>
        /// <returns>object</returns>
        public static object InvokeFunction(string methodName, object[] methodParams)
        {
            Type customFunctions = Type.GetType(CustomFunctionClass);
            ConstructorInfo customFuncConstructor = customFunctions.GetConstructor(Type.EmptyTypes);
            object classObject = customFuncConstructor.Invoke(new object[] { });

            MethodInfo customFuncMethod = customFunctions.GetMethod(methodName);
            object customFuncValue = customFuncMethod.Invoke(classObject, methodParams);
            return customFuncValue;
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

                Type type = obj.GetType();
                PropertyInfo info = type.GetProperty(part);
                if (info != null)
                {
                    obj = info.GetValue(obj, null);
                }
            }

            return obj;
        }
    }
}