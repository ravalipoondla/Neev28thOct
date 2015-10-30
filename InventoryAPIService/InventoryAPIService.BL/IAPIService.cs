using Inventory.RestAPI.DAL;
using Inventory.RestAPI.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Inventory.RestAPI.BL
{
    public interface IAPIService
    {
        List<UserRole> GetUserRoles();
        bool ValidateUser(string userRoleName,string passCode);
        List<ProductInventory> GetAllProductInventories();
        bool AddProductInventory(ProductInventory productInventory);
        List<UserActivity> GetUserActivities(int roleId, string fromDate, string toDate);
        bool DeleteProductInventory(int productInventoryId);
        List<ProductInventoryItem> GetInventoryData(string fromDate, string toDate);
        MemoryStream GenerateInventoryDataExcelAsStream(string ActivitiesIDs, string ExportFomratId,string fromDate, string toDate);
        void SendEmail(string toEmailAddress, Stream fileStream , string ExportFomratId);
        bool AddRawMaterialInventory(RawMaterialInventory rawMaterialInventory);
        List<RawMaterialInventory> GetAllRawMaterialInventories();
        List<ProductInventoryItem> GetAllProductInventoryItems();
        bool DeleteProductInventoryItem(int productInventoryTranId);
        bool DeleteRawMaterialInventory(int rawMaterialInventoryId);
        bool DeleteRawMaterialInventoryItem(int rawMaterialInventoryTranId);
        List<RawMaterialInventoryItem> GetAllRawMaterialInventoryItems();
    }
}
