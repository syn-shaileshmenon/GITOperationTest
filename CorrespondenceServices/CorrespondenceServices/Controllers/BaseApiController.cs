// <copyright file="BaseApiController.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace CorrespondenceServices.Controllers
{
    using System.Collections.Generic;
    using CorrespondenceServices.Classes;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.RestHelper.Classes;
    using Mkl.WebTeam.RestHelper.Interfaces;
    using Mkl.WebTeam.StorageProvider.Implementors;
    using Mkl.WebTeam.StorageProvider.Interfaces;
    using Mkl.WebTeam.WebCore2.Controllers;

    /// <summary>
    /// The Base Api Controller class
    /// </summary>
    public class BaseApiController : BaseRestController
    {
        /// <summary>
        /// sync lock
        /// </summary>
        private static readonly object SyncLock = new object();

        /// <summary>
        /// Boot strap
        /// </summary>
        private readonly IBootstrapper bootStrapper;

        /// <summary>
        /// Json manager
        /// </summary>
        private readonly IJsonManager jsonManager;

        /// <summary>
        /// Storage Manager
        /// </summary>
        private IStorageManager storageManager;

        /// <summary>
        /// template processor
        /// </summary>
        private ITemplateProcessor templateProcessor;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseApiController" /> class
        /// </summary>
        /// <param name="jsonManager">The json provider.</param>
        /// <param name="bootStrapper">The boot strap interface for storage provider .</param>
        /// <param name="restHelper">The rest helper.</param>
        /// <param name="storageManager">The storage manager.</param>
        /// <param name="templateProcessor">The template processor.</param>
        public BaseApiController(
            IJsonManager jsonManager,
            IBootstrapper bootStrapper,
            IRestHelper restHelper,
            IStorageManager storageManager,
            ITemplateProcessor templateProcessor)
        {
            this.bootStrapper = bootStrapper;
            this.jsonManager = jsonManager;
            this.RestHelper = restHelper;
            this.storageManager = storageManager;
            this.templateProcessor = templateProcessor;
        }

        /// <summary>
        /// Gets or sets the authenticated user.
        /// </summary>
        /// <value>The authenticated user.</value>
        public User AuthenticatedUser { get; set; }

        /// <summary>
        /// Gets Storage provider manager
        /// </summary>
        public IStorageManager StorageProviderManager
        {
            get
            {
                if (this.storageManager == null)
                {
                    lock (SyncLock)
                    {
                        if (this.storageManager == null)
                        {
                            this.storageManager = new StorageManager(this.bootStrapper);
                        }
                    }
                }

                return this.storageManager;
            }
        }

        /// <summary>
        /// Gets template processor
        /// </summary>
        public ITemplateProcessor TemplateProcessor
        {
            get
            {
                if (this.templateProcessor == null)
                {
                    lock (SyncLock)
                    {
                        if (this.templateProcessor == null)
                        {
                            this.templateProcessor = new TemplateProcessor(this.jsonManager);
                        }
                    }
                }

                return this.templateProcessor;
            }
        }

        /// <summary>
        /// Gets or sets the rest helper.
        /// </summary>
        /// <value>The rest helper.</value>
        protected IRestHelper RestHelper { get; set; }

        /// <summary>
        /// Gets the list of carriers
        /// </summary>
        protected virtual List<Carrier> Carriers => CarrierCollection.Carriers(this.RestHelper, this.AuthenticatedUser);
    }
}