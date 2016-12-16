// <copyright file="WebApiConfig.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices
{
    using System.Configuration;
    using System.Net.Http.Formatting;
    using System.Web.Http;
    using System.Web.Http.ExceptionHandling;
    using Mkl.WebTeam.WebCore2.Loggers;

    /// <summary>
    /// Configures Web API
    /// </summary>
    public static class WebApiConfig
    {
        /// <summary>
        /// Gets the URL prefix.
        /// </summary>
        /// <value>
        /// The URL prefix.
        /// </value>
        public static string UrlPrefix
        {
            get
            {
                return "api";
            }
        }

        /// <summary>
        /// Gets the URL prefix relative.
        /// </summary>
        /// <value>
        /// The URL prefix relative.
        /// </value>
        public static string UrlPrefixRelative
        {
            get
            {
                return "~/api";
            }
        }

        /// <summary>
        /// Registers the specified configuration.
        /// </summary>
        /// <param name="config">The configuration.</param>
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services
            config.Services.Add(typeof(IExceptionLogger), new Log4NetExceptionLogger(ConfigurationManager.AppSettings["CorrespondenceLogger"]));

            config.Formatters.Add(new BsonMediaTypeFormatter());

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional });
        }
    }
}
