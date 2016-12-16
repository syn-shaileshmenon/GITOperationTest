// <copyright file="Startup.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(CorrespondenceServices.App_Start.Startup))]

namespace CorrespondenceServices.App_Start
{
    /// <summary>
    /// Startup class for Owin
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// Configurations the specified application.
        /// </summary>
        /// <param name="app">The application.</param>
        public void Configuration(IAppBuilder app)
        {
            // For more information on how to configure your application, visit http://go.microsoft.com/fwlink/?LinkID=316888
        }
    }
}
