// <copyright file="Global.asax.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices
{
    using System.Web.Http;
    using System.Web.Mvc;
    using System.Web.Optimization;
    using System.Web.Routing;
    using Common.Logging;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.RestHelper.Helpers;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Implementors;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.WebCore2;
    using SimpleInjector;
    using SimpleInjector.Integration.WebApi;

    /// <summary>
    /// Web Api Application
    /// </summary>
    public class WebApiApplication : RestHttpApplication
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WebApiApplication"/> class.
        /// </summary>
        public WebApiApplication()
            : base("Correspondence", "CorrespondenceTrace")
        {
        }

        /// <summary>
        /// Gets the default logger.
        /// </summary>
        /// <value>The default logger.</value>
        public ILog DefaultLogger
        {
            get { return this.Log; }
        }

        /// <summary>
        /// Application_s the start.
        /// </summary>
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            Aspose.Words.License license = new Aspose.Words.License();
            license.SetLicense("Aspose.Words.lic");
            ConfigureContainer();
        }

        /// <summary>
        /// Configures the container.
        /// </summary>
        private static void ConfigureContainer()
        {
            var container = new Container();

            container.Options.AllowOverridingRegistrations = true;
            container.RegisterWebApiRequest<IBootstrapper, Bootstrapper>();
            container.RegisterWebApiRequest<IJsonManager, JsonManager>();
            container.RegisterWebApiRequest<IRestHelper, RestHelper>();
            container.RegisterWebApiRequest<IStorageManager, StorageManager>();
            container.RegisterWebApiRequest<ITemplateProcessor, TemplateProcessor>();
            container.RegisterSingleton<ILog>(() => LogEvent.Log);

            // This is an extension method from the integration package.
            container.RegisterWebApiControllers(GlobalConfiguration.Configuration);

            container.Verify();

            GlobalConfiguration.Configuration.DependencyResolver = new SimpleInjectorWebApiDependencyResolver(container);
        }
    }
}
