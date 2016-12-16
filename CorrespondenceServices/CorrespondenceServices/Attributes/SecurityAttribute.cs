// ***********************************************************************
// Assembly         : CorrespondenceServices
// Author           : rsteelea
// Created          : 09-02-2014
//
// Last Modified By : rsteelea
// Last Modified On : 09-02-2014
// ***********************************************************************
// <copyright file="SecurityAttribute.cs" company="Markel">
//     Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

/// <summary>
/// The Attributes namespace.
/// </summary>
namespace CorrespondenceServices.Attributes
{
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Web.Http.Controllers;
    using System.Web.Http.Filters;
    using Mkl.WebTeam.RestHelper.Classes;
    using Newtonsoft.Json;

    /// <summary>
    /// Class SecurityAttribute.
    /// </summary>
    public class SecurityAttribute : ActionFilterAttribute
    {
        /// <summary>
        /// Occurs before the action method is invoked.
        /// </summary>
        /// <param name="actionContext">The action context.</param>
        public override void OnActionExecuting(HttpActionContext actionContext)
        {
            base.OnActionExecuting(actionContext);
            GetAuthenticatedUser(actionContext);
        }

        /// <summary>
        /// Called when [action executing asynchronous].
        /// </summary>
        /// <param name="actionContext">The action context.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>System.Threading.Tasks.Task.</returns>
        public override System.Threading.Tasks.Task OnActionExecutingAsync(HttpActionContext actionContext, CancellationToken cancellationToken)
        {
            GetAuthenticatedUser(actionContext);

            return base.OnActionExecutingAsync(actionContext, cancellationToken);
        }

        /// <summary>
        /// Gets the authenticated user.
        /// </summary>
        /// <param name="actionContext">The action context.</param>
        private static void GetAuthenticatedUser(HttpActionContext actionContext)
        {
            var user = new User();

            if (actionContext.Request.Headers.Contains(Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOUSERNAME))
            {
                var headerValue = actionContext.Request.Headers.FirstOrDefault(h => h.Key == Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOUSERNAME).Value.FirstOrDefault();
                user.Username = headerValue;

                if (actionContext.Request.Headers.Contains(Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOFIRSTNAME))
                {
                    headerValue = actionContext.Request.Headers.FirstOrDefault(h => h.Key == Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOFIRSTNAME).Value.FirstOrDefault();
                    user.FirstName = headerValue;
                }

                if (actionContext.Request.Headers.Contains(Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOLASTNAME))
                {
                    headerValue = actionContext.Request.Headers.FirstOrDefault(h => h.Key == Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOLASTNAME).Value.FirstOrDefault();
                    user.LastName = headerValue;
                }

                if (actionContext.Request.Headers.Contains(Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOAGENCYCODE))
                {
                    headerValue = actionContext.Request.Headers.FirstOrDefault(h => h.Key == Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOAGENCYCODE).Value.FirstOrDefault();
                    user.AgencyCode = headerValue;
                }

                if (actionContext.Request.Headers.Contains(Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOAGENCYNAME))
                {
                    headerValue = actionContext.Request.Headers.FirstOrDefault(h => h.Key == Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOAGENCYNAME).Value.FirstOrDefault();
                    user.AgencyName = headerValue;
                }

                if (actionContext.Request.Headers.Contains(Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOISINTERNAL))
                {
                    headerValue = actionContext.Request.Headers.FirstOrDefault(h => h.Key == Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOISINTERNAL).Value.FirstOrDefault();
                    bool internalUser = false;
                    bool.TryParse(headerValue, out internalUser);
                    user.IsInternalUser = internalUser;
                }

                if (actionContext.Request.Headers.Contains(Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOISBROKERAGE))
                {
                    headerValue = actionContext.Request.Headers.FirstOrDefault(h => h.Key == Mkl.WebTeam.RestHelper.Classes.HeaderConstants.SSOISBROKERAGE).Value.FirstOrDefault();
                    bool localBrokerage = false;
                    bool.TryParse(headerValue, out localBrokerage);
                    user.IsBrokerageUser = localBrokerage;
                }
            }
            else
            {
                if (File.Exists("./testuser.json"))
                {
                    var json = File.ReadAllText("./testuser.json");
                    user = JsonConvert.DeserializeObject<User>(json);
                }
                else
                {
                    user.Username = "bjoyce@210100_crcins.com";
                    user.IsInternalUser = true;
                    user.IsBrokerageUser = false;
                    user.AgencyName = null;
                    user.AgencyCode = null;
                }
            }

            ((CorrespondenceServices.Controllers.BaseApiController)actionContext.ControllerContext.Controller).AuthenticatedUser = user;
        }
    }
}