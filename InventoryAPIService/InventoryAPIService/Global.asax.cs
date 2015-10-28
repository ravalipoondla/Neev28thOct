// -----------------------------------------------------------------------
// <copyright file="Global.asax.cs" company="Neev">
// TODO: Update copyright text.
// </copyright>
// ----------------------------------------------------------------------

namespace PRDTool.WebAPI.Service
{
    using System;
    using System.Net;
    using System.ServiceModel.Activation;
    using System.Web;
    using System.Web.Routing;
    using System.ServiceModel;
    using Inventory.RestAPI.Service;

    /// <summary>
    /// Global class
    /// </summary>
    public class Global : HttpApplication
    {

        /// <summary>
        /// Application Start
        /// </summary>
        /// <param name="sender">sender parameter</param>
        /// <param name="e">is event argument</param>
        private void Application_Start(object sender, EventArgs e)
        {
            //Register Route
            this.RegisterRoutes();
        }

        /// <summary>
        /// Application End
        /// </summary>
        /// <param name="sender">sender parameter</param>
        /// <param name="e">event argument</param>
        private void Application_End(object sender, EventArgs e)
        {
            //  Code that runs on application shutdown
        }

        /// <summary>
        /// Application Error
        /// </summary>
        /// <param name="sender">sender parameter </param>
        /// <param name="e">event argument</param>
        private void Application_Error(object sender, EventArgs e)
        {
            
        }

        /// <summary>
        /// register routes
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "ignore"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "ignore")]
        private void RegisterRoutes()
        {
            // Edit the base address of Service1 by replacing the "Service1" string below
            RouteTable.Routes.Add(new ServiceRoute(string.Empty, new WebServiceHostFactory(), typeof(InventoryAPIService)));
        }
    }
}
