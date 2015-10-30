// -----------------------------------------------------------------------
// <copyright file="InventoryAPIService.cs" company="Neev">
// TODO: Update copyright text.
// </copyright>
// ----------------------------------------------------------------------

namespace Inventory.RestAPI.Service
{
    using Inventory.RestAPI.BL;
    using Inventory.RestAPI.Entities;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.ServiceModel;
    using System.ServiceModel.Activation;
    using System.ServiceModel.Web;

    /// <summary>
    /// Partial class representation of InventoryAPIService
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1724:TypeNamesShouldNotMatchNamespaces", Justification = "ignore"), ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]
    public partial class InventoryAPIService
    {
        /// <summary>
        /// field to hold APi service adaptor
        /// </summary>
        private IAPIService apiService = new APIService();

        /// <summary>
        /// Initializes a new instance of the  <see cref="InventoryAPIService"></see> class
        /// </summary>
        public InventoryAPIService()
        {
        }

        /// <summary>
        /// Gets user roles
        /// </summary>
        /// <returns></returns>
        [WebInvoke(UriTemplate = "/User/Roles",Method="GET",ResponseFormat = WebMessageFormat.Json,RequestFormat = WebMessageFormat.Json,BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public List<UserRole> UserRoles()
        {
            return this.apiService.GetUserRoles();
        }

        /// <summary>
        /// Validates User and returns Activities if validation successfull
        /// </summary>
        /// <param name="userRole">user role name</param>
        /// <param name="PassCode">pass code</param>
        /// <returns>returns null if validation fails else returns user activities</returns>
        [WebGet(UriTemplate = "/Validate/User/{UserRole}/{PassCode}", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public bool ValidateUser(string userRole,string passCode)
        {
            return this.apiService.ValidateUser(userRole,passCode);
        }
        
        /// <summary>
        /// Gets User Activities
        /// </summary>
        /// <returns>returns User activities</returns>
        [WebInvoke(UriTemplate = "/User/Activities/{roleId}?fromDate={fromDate}&toDate={toDate}", Method = "GET", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public List<UserActivity> GetUserActivities(string roleId,string fromDate,string toDate)
        {
            return this.apiService.GetUserActivities(Convert.ToInt32(roleId),fromDate,toDate);
        }

        /// <summary>
        /// Get all Proudct Invntories
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/Product", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public List<ProductInventory> GetAllProductInventories()
        {
            return this.apiService.GetAllProductInventories();
        }

        /// <summary>
        /// adds Proudct Inventory
        /// </summary>
        /// <returns>returns success or failure flag</returns>
        [WebInvoke(UriTemplate = "/Inventory/Product/Add", Method = "POST")]
        [OperationContract]
        public bool AddProductInventory(ProductInventory pi)
        {
            return this.apiService.AddProductInventory(pi);
        }

        /// <summary>
        /// Get all Proudct Invntories
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/Product/Delete/{productInventoryId}")]
        [OperationContract]
        public bool DeleteProductInventory(string productInventoryId)
        {
            return this.apiService.DeleteProductInventory(Convert.ToInt32(productInventoryId));
        }

        /// <summary>
        /// Get all Proudct Invntories
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/Products?fromDate={fromDate}&toDate={toDate}", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public List<ProductInventoryItem> GetAllInventoryData(string fromDate,string toDate)
        {
            return this.apiService.GetInventoryData(fromDate , toDate );
        }

        ///// <summary>
        ///// Get all Proudct Invntories
        ///// </summary>
        ///// <returns>returns all Product Inventories</returns>
        //[WebGet(UriTemplate = "/Inventory/Product/Export/{activitiesIDs}/{ExportFomratId}?fromDate={fromDate}&toDate={toDate}", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        //[OperationContract]
        //public Stream GetExportData(string activitiesIDs, string ExportFomratId, string fromDate, string toDate)
        //{
        //    Stream stream = null;
        //    if (!string.IsNullOrEmpty(activitiesIDs))
        //    {
        //        stream = this.apiService.GenerateInventoryDataExcelAsStream(activitiesIDs, ExportFomratId,  fromDate, toDate);
        //        stream.Position = 0;
        //        WebOperationContext.Current.OutgoingResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        WebOperationContext.Current.OutgoingResponse.Headers.Add("Content-disposition", "inline; filename=Export.xlsx");
        //        //WebOperationContext.Current.OutgoingResponse.ContentType = "application/pdf";
        //        //WebOperationContext.Current.OutgoingResponse.Headers.Add("Content-disposition", "inline; filename=Export.pdf");
        //    }
        //    return stream;
        //}

        /// <summary>
        /// Get all Proudct Invntories
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/Product/Export/Email/{toEmailAddress}/{activitiesIDs}/{ExportFomratId}?fromDate={fromDate}&toDate={toDate}", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public bool SendExportDataToEmail(string toEmailAddress,string activitiesIDs,string ExportFomratId, string fromDate, string toDate)
        {
            try
            {
                Stream stream = null;
                if (!string.IsNullOrEmpty(activitiesIDs))
                {
                    stream = this.apiService.GenerateInventoryDataExcelAsStream(activitiesIDs, ExportFomratId, fromDate, toDate);
                }
                if (stream != null)
                {
                    this.apiService.SendEmail(toEmailAddress, stream, ExportFomratId);
                }
                return true;
            }
            catch(Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// adds Raw Material Inventory
        /// </summary>
        /// <returns>returns success or failure flag</returns>
        [WebInvoke(UriTemplate = "/Inventory/RawMaterial/Add", Method = "POST")]
        [OperationContract]
        public bool AddRawMaterialInventory(RawMaterialInventory ri)
        {
            return this.apiService.AddRawMaterialInventory(ri);
        }

        /// <summary>
        /// Get all Proudct Invntories
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/RawMaterialInventory", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public List<RawMaterialInventory> GetAllRawMaterialInventories()
        {
            return this.apiService.GetAllRawMaterialInventories();
        }

        /// <summary>
        /// Get all Proudct Invntory Items
        /// </summary>
        /// <returns>returns all Product Inventory Items</returns>
        [WebGet(UriTemplate = "/Inventory/Product/Items", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public List<ProductInventoryItem> GetAllProductInventoryItems()
        {
            return this.apiService.GetAllProductInventoryItems();
        }

        /// <summary>
        /// Get all Proudct Invntories
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/Product/Item/Delete/{productInventoryTranId}")]
        [OperationContract]
        public bool DeleteProductInventoryItem(string productInventoryTranId)
        {
            return this.apiService.DeleteProductInventoryItem(Convert.ToInt32(productInventoryTranId));
        }


        /// <summary>
        /// adds Sales Details
        /// </summary>
        /// <returns>returns success or failure flag</returns>
        [WebInvoke(UriTemplate = "/Inventory/Product/Sales", Method = "POST")]
        [OperationContract]
        public bool AddProductSales(string productSoldItem)
        {
            List<ProductInventory> lstProduct = (List<ProductInventory>)JsonConvert.DeserializeObject(productSoldItem, typeof(List<ProductInventory>));

            foreach (ProductInventory pi in lstProduct)
            {
                this.apiService.AddProductInventory(pi);
            }

            return true;
        }

        /// <summary>
        /// Get all Raw Material Invntory Items
        /// </summary>
        /// <returns>returns all Raw Material Inventory Items</returns>
        [WebGet(UriTemplate = "/Inventory/RawMaterial/Items", ResponseFormat = WebMessageFormat.Json, RequestFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        [OperationContract]
        public List<RawMaterialInventoryItem> GetAllRawMaterialInventoryItems()
        {
            return this.apiService.GetAllRawMaterialInventoryItems();
        }

        /// <summary>
        /// Get all Proudct Invntories
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/RawMaterial/Item/Delete/{rawMaterialInventoryTranId}")]
        [OperationContract]
        public bool DeleteRawMaterialInventoryItem(string rawMaterialInventoryTranId)
        {
            return this.apiService.DeleteRawMaterialInventoryItem(Convert.ToInt32(rawMaterialInventoryTranId));
        }

        /// <summary>
        /// Delete 
        /// </summary>
        /// <returns>returns all Product Inventories</returns>
        [WebGet(UriTemplate = "/Inventory/RawMaterial/Delete/{rawMaterialInventoryId}")]
        [OperationContract]
        public bool DeleteRawMaterialInventory(string rawMaterialInventoryId)
        {
            return this.apiService.DeleteRawMaterialInventory(Convert.ToInt32(rawMaterialInventoryId));
        }
    }
}
