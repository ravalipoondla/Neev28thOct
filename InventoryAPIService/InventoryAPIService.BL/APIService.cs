using Inventory.RestAPI.DAL;
using Inventory.RestAPI.Entities;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Xml;

namespace Inventory.RestAPI.BL
{
    enum Activities
    {
            Inventory = 1,
            Sales,
            InTransit,
            Personnel,
            Returned,
            RawMaterialInventory
    }

    public class APIService : IAPIService
    {
        public List<UserRole> GetUserRoles()
        {
            var context = new NeevDatabaseContainer();
            var userRoles = (from item in context.GetUserRoles()
                            select new UserRole { Id = item.user_role_id, Name = item.user_role }).ToList<UserRole>();
           return userRoles;
        }

        public bool ValidateUser(string userRoleName, string passCode)
        {
            var context = new NeevDatabaseContainer();
            return context.ValidateUser(userRoleName,passCode).FirstOrDefault().Value.ToString() == "1" ? true : false;
        }

        public List<ProductInventory> GetAllProductInventories()
        {
            var context = new NeevDatabaseContainer();
            var productInventories = (from productInventory in context.GetALLInventories()
                                 select new ProductInventory { Id = productInventory.product_inventory_id, Name = productInventory.product_name,Quantity=productInventory.quantity.Value,CreatedDate= productInventory.creation_dt }).ToList<ProductInventory>();
            return productInventories;
        }

        public bool AddProductInventory(ProductInventory pi)
        {
            try
            {
                var context = new NeevDatabaseContainer();
                context.AddProductInventoryItem(pi.Id,pi.Name, pi.Quantity, pi.UnitPrice, pi.SoldFlag, pi.ReturnedFlag);
                return true;
            }
            catch(Exception)
            {
                return false;
            }
        }


        public bool DeleteProductInventory(int productInventoryId)
        {
            try
            {
                var context = new NeevDatabaseContainer();
                context.DeleteProductInventory(productInventoryId);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public List<UserActivity> GetUserActivities(int roleId,string fromDate,string toDate)
        {
            DateTime? dtFrom = !string.IsNullOrEmpty(fromDate) ? Convert.ToDateTime(fromDate) as DateTime? : null;
            DateTime? dtTo = !string.IsNullOrEmpty(toDate) ? Convert.ToDateTime(toDate) as DateTime? : null;

            var context = new NeevDatabaseContainer();
            var userActivities = (from userActivity in context.GetUserActivities(roleId,dtFrom,dtTo)
                                      select new UserActivity { Id = userActivity.activity_id, Name = userActivity.activity_name , quantity = userActivity.quantity,price = userActivity.price }).ToList<UserActivity>();
            return userActivities;
                
        }

        public List<ProductInventoryItem> GetInventoryData(string fromDate, string toDate)
        {
            DateTime? dtFrom = !string.IsNullOrEmpty(fromDate) ? Convert.ToDateTime(fromDate) as DateTime? : null;
            DateTime? dtTo = !string.IsNullOrEmpty(toDate) ? Convert.ToDateTime(toDate) as DateTime? : null;

            var context = new NeevDatabaseContainer();
            var inventoryItems = (from inventoryItem in context.GetInventoryData(dtFrom, dtTo)
                                  select new ProductInventoryItem { Id = inventoryItem.product_inventory_id, Name = inventoryItem.product_name, Quantity = inventoryItem.quantity, Price = inventoryItem.price,Percentage=inventoryItem.percentage }).ToList<ProductInventoryItem>();
            return inventoryItems;
        }


        /// <summary>
        /// preparing file path, delete existing file and create the folder.
        /// </summary>
        /// <param name="filepath">xml file path</param>
        public static void PrepareFilePath(string filepath)
        {
            FileInfo fi = new FileInfo(filepath);
            if (fi.Exists)
            {
                fi.Delete();
            }

            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
        }

        public MemoryStream GenerateInventoryDataExcelAsStream(string activitiesIDs, string ExportFomratId, string fromDate, string toDate)
        {
            DateTime? dtFrom = !string.IsNullOrEmpty(fromDate) ? Convert.ToDateTime(fromDate) as DateTime? : null;
            DateTime? dtTo = !string.IsNullOrEmpty(toDate) ? Convert.ToDateTime(toDate) as DateTime? : null;
            var fileStream = new MemoryStream();
            var context = new NeevDatabaseContainer();

            try
            {
                //first prepare the path to save the file
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["ExportDataTemplate"]);
                FileInfo existingFile = new FileInfo(excelPath);
                string folderLocation = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["TempExportDataLocation"]);

                try
                {
                    Array.ForEach(Directory.GetFiles(folderLocation), File.Delete);
                }
                catch (Exception ex)
                {
                    Log(ex.ToString());
                }
                string fileNameWithoutExtension = existingFile.Name.Split(new char[] { '.' })[0] + "_" + DateTime.Now.ToFileTime().ToString();
                string tempPath = Path.Combine(folderLocation, fileNameWithoutExtension + ".xlsx");
                PrepareFilePath(tempPath);
                existingFile.CopyTo(tempPath);
                FileInfo defaultTemplateFile = new FileInfo(tempPath);
                string pdfFilePath;
                string filePath = tempPath;


                using (OleDbConnection _objCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tempPath + @";Extended Properties=""Excel 12.0;HDR=Yes"""))
                {
                    _objCon.Open();

                    using (var excelCommand = new OleDbCommand())
                    {
                        excelCommand.Connection = _objCon;

                        foreach (string activityId in activitiesIDs.Split(new char[] { ',' }))
                        {
                            var inventoryItems = context.GetExportData(Convert.ToInt32(activityId), dtFrom, dtTo);

                            int actId = Convert.ToInt32(activityId);

                            excelCommand.CommandText = "CREATE TABLE [" + Enum.GetName(typeof(Activities), actId) + "] ([Name of product] VARCHAR, [Quantity] INT, [Price] DOUBLE,[Percentage] DOUBLE );";
                            excelCommand.ExecuteNonQuery();

                            //Name of product	Quantity	Price	Percentage
                            excelCommand.CommandText = "INSERT INTO [" + Enum.GetName(typeof(Activities), actId) + "$] ([Name of product],[Quantity],[Price],[Percentage]" +
                               ") VALUES(?,?,?,?)";

                            excelCommand.Parameters.Add(new OleDbParameter { DbType = System.Data.DbType.String, Size = 1000 });
                            excelCommand.Parameters.Add(new OleDbParameter { DbType = System.Data.DbType.Int32 });
                            excelCommand.Parameters.Add(new OleDbParameter { DbType = System.Data.DbType.Double });
                            excelCommand.Parameters.Add(new OleDbParameter { DbType = System.Data.DbType.Double });

                            var isPrepared = false;

                            foreach (var inventory in inventoryItems)
                            {
                                excelCommand.Parameters[0].Value = inventory.product_name;
                                excelCommand.Parameters[1].Value = inventory.quantity;
                                excelCommand.Parameters[2].Value = inventory.price;
                                excelCommand.Parameters[3].Value = inventory.percentage;

                                if (!isPrepared)
                                {
                                    excelCommand.Prepare();
                                    isPrepared = true;
                                }
                                excelCommand.ExecuteNonQuery();
                            }


                        }
                    }
                }

                Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                xlApp.DisplayAlerts = false;

                Workbook xlWorkBook = xlApp.Workbooks.Open(tempPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Sheets worksheets = xlWorkBook.Worksheets;
                worksheets[1].Delete();
                xlWorkBook.Save();
                //worksheets = xlWorkBook.Worksheets;

                if (ExportFomratId == "2")
                {
                    xlWorkBook = xlApp.Workbooks.Open(tempPath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    DataSet datasetInventory = new DataSet();
                    List<string> strSheetNames = new List<string>();
                    foreach (Worksheet worksheet in xlWorkBook.Worksheets)
                    {

                        using (OleDbConnection objCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tempPath + @";Extended Properties=""Excel 12.0;HDR=Yes"""))
                        {
                            //// First get the offers
                            OleDbDataAdapter objDataAdapter = new OleDbDataAdapter("SELECT * FROM [" + worksheet.Name + "$]", objCon);
                            objDataAdapter.Fill(datasetInventory, worksheet.Name);
                            objDataAdapter.Dispose();
                        }
                        strSheetNames.Add(worksheet.Name);
                    }

                    Document document = new Document();
                    pdfFilePath = folderLocation + fileNameWithoutExtension + ".pdf";
                    filePath = pdfFilePath;
                    //PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(folderLocation + fileNameWithoutExtension + ".pdf", FileMode.Create));
                    var fst = new FileStream(pdfFilePath, FileMode.Create);
                    PdfWriter writer = PdfWriter.GetInstance(document, fst);

                    document.Open();
                    iTextSharp.text.Font font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 5);
                    int i = 0;
                    foreach (System.Data.DataTable dt in datasetInventory.Tables)
                    {
                        document.Add(new Paragraph(strSheetNames[i++]));
                        document.Add(new Paragraph("                 "));
                        PdfPTable table = new PdfPTable(dt.Columns.Count);
                        //PdfPRow row = null;
                        float[] widths = new float[] { 4f, 4f, 4f, 4f };

                        table.SetWidths(widths);

                        table.WidthPercentage = 100;
                        //int iCol = 0;
                        //string colname = "";
                        PdfPCell cell = new PdfPCell(new Phrase("Neev"));

                        cell.Colspan = dt.Columns.Count;

                        foreach (DataColumn c in dt.Columns)
                        {

                            table.AddCell(new Phrase(c.ColumnName, font5));
                        }

                        foreach (DataRow r in dt.Rows)
                        {
                            if (dt.Rows.Count > 0)
                            {
                                table.AddCell(new Phrase(r[0].ToString(), font5));
                                table.AddCell(new Phrase(r[1].ToString(), font5));
                                table.AddCell(new Phrase(r[2].ToString(), font5));
                                table.AddCell(new Phrase(r[3].ToString(), font5));
                            }
                        }
                        //Paragraph para = new Paragraph(
                        document.Add(table);
                        document.Add(new Paragraph("                 "));
                    }
                    
                    //writer.Close();
                    document.Close();
                }
                
                xlWorkBook.Close();

                releaseObject(worksheets);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                var fStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                fileStream.SetLength(fStream.Length);
                fStream.Read(fileStream.GetBuffer(), 0, (int)fileStream.Length);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                return null;
            }
            return fileStream;
        }

        /// <summary>
        /// Read Excel For Offers
        /// </summary>
        /// <param name="prdPath">PRD Path</param>
        /// <param name="fileName">File Name</param>
        /// <returns>Price Plan Table</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope", Justification = "Reviewed")]
        private static System.Data.DataTable ReadExcelForOffers(string filePath)
        {
            DataSet datasetPricePlan = new DataSet();
            using (OleDbConnection objCon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + @";Extended Properties=""Excel 12.0;HDR=Yes"""))
            {
                //// First get the offers
                OleDbDataAdapter objDataAdapter = new OleDbDataAdapter("SELECT * FROM [Offers$]", objCon);
                objDataAdapter.Fill(datasetPricePlan, "PricePlan");
                objDataAdapter.Dispose();
            }

            return datasetPricePlan.Tables[0];
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        #region Log Message in case of Error And Sending mail part.
        /// <summary>
        /// Log messages
        /// </summary>
        /// <param name="logMessage">Log message</param>
        /// <param name="error">Error or informational</param>
        public void SendEmail(string toEmailAddress, Stream fileStream, string ExportFomratId)
        {
            MailMessage mail = new MailMessage();
            SmtpClient smtpServer = null;
            try
            {

                smtpServer = new SmtpClient(System.Configuration.ConfigurationManager.AppSettings["smtpClient"]);
                smtpServer.Port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["smtpPort"]);
                //smtpServer = new SmtpClient("smtp.gmail.com");
                //smtpServer.Port = 465;//587
                //smtpServer.EnableSsl = true;
                //smtpServer.UseDefaultCredentials = false;
                //string user = ""; //<--Enter your gmail id here
                //string pass = "";//<--Enter Your gmail password here

                //smtpServer.Credentials = new NetworkCredential(user, pass);


                XmlDocument configXML = new XmlDocument();
                configXML.Load(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ConfigurationManager.AppSettings["StatusEmailNotification"]));

                mail.From = new MailAddress(System.Configuration.ConfigurationManager.AppSettings["fromEmailAddress"].ToString());

                mail.To.Add(new MailAddress(toEmailAddress));



                //foreach (XmlNode childnode in configXML.SelectNodes("/Emails/To/Email_ID"))
                //{
                //    string email_ID = childnode.Attributes["ID"].Value.ToString();
                //    mail.To.Add(new MailAddress(email_ID));
                //}

                foreach (XmlNode childnode in configXML.SelectNodes("/Emails/CC/Email_ID"))
                {
                    string email_ID = childnode.Attributes["ID"].Value;
                    mail.CC.Add(new MailAddress(email_ID));
                }

                mail.Subject = configXML.SelectSingleNode("/Emails/Subject").InnerText;

                mail.Body = configXML.SelectSingleNode("/Emails/Message").InnerText;

                if(ExportFomratId == "1")
                    mail.Attachments.Add(new Attachment(fileStream, "ExportedData.xlsx"));
                else if(ExportFomratId == "2")
                    mail.Attachments.Add(new Attachment(fileStream, "ExportedData.pdf"));

                smtpServer.Send(mail);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            
                //Log(Constants.LineBreak + Environment.NewLine + "DateTime:\t" + DateTime.Now + Environment.NewLine +
                //                ex.ToString() + Environment.NewLine + Environment.NewLine + Constants.LineBreak + Environment.NewLine);
            }
            finally
            {
                mail.Dispose();
                smtpServer.Dispose();
            }


            //FileLog(logMessage);
        }
        #endregion

        #region LogMessage to track
        /// <summary>
        /// log error and/or information message
        /// </summary>
        /// <param name="message">message parameter</param>
        private static void Log(string message)
        {
            FileStream logfile = null;
            StreamWriter writer = null;
            try
            {
                string path = ConfigurationManager.AppSettings["InventoryAPILog"];
                string logFile = ConfigurationManager.AppSettings["ExceptionLogFileName"];

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                path = path + logFile;
                if (File.Exists(path))
                {
                    logfile = new FileStream(path, FileMode.Append, FileAccess.Write);
                }
                else
                {
                    logfile = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
                }

                writer = new StreamWriter(logfile);

                //DateTime time = DateTime.Now;
                writer.WriteLine(message); //+ " : " + time.ToString());
            }
            catch (Exception)
            {
                ////WriteMessageToDisk(Constants.LINE_BREAK + Environment.NewLine + "DateTime:\t" + DateTime.Now + Environment.NewLine +
                ////                    ex.ToString() + Environment.NewLine + Environment.NewLine + Constants.LINE_BREAK + Environment.NewLine);
            }
            finally
            {
                if (writer != null)
                {
                    writer.Dispose();
                }

                if (logfile != null)
                {
                    logfile.Dispose();
                }
            }
        }
        #endregion

        public bool AddRawMaterialInventory(RawMaterialInventory rawMaterialInventory)
        {
            try
            {
                var context = new NeevDatabaseContainer();
                context.AddRawMaterialInventoryItem(rawMaterialInventory.Id, rawMaterialInventory.Name, rawMaterialInventory.Quantity, rawMaterialInventory.UnitPrice,rawMaterialInventory.Threshold);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public List<RawMaterialInventory> GetAllRawMaterialInventories()
        {
            var context = new NeevDatabaseContainer();
            var rawMaterialInventories = (from rawMaterialInventory in context.GetALLRawMaterialInventories()
                                          select new RawMaterialInventory { Id = rawMaterialInventory.raw_material_inventory_id, Name = rawMaterialInventory.raw_material_name,Threshold= rawMaterialInventory.threshhold_value,Quantity=rawMaterialInventory.available_quantity.Value,UnitPrice = rawMaterialInventory.price.Value,CreatedDate= rawMaterialInventory.creation_dt}).ToList<RawMaterialInventory>();
            return rawMaterialInventories;
        }

        public List<ProductInventoryItem> GetAllProductInventoryItems()
        {
            var context = new NeevDatabaseContainer();
            var productInventoryItems = (from productInventoryItem in context.GetALLInventoryItems()
                                         select new ProductInventoryItem { Id = productInventoryItem.product_inventory_trans_id, Name = productInventoryItem.product_name, Quantity = productInventoryItem.quantity, Price = productInventoryItem.price.Value,CreatedDate = productInventoryItem.creation_dt }).ToList<ProductInventoryItem>();
            return productInventoryItems;
        }

        public bool DeleteProductInventoryItem(int productInventoryTranId)
        {
            try
            {
                var context = new NeevDatabaseContainer();
                context.DeleteProductInventoryItem(productInventoryTranId);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool DeleteRawMaterialInventory(int rawMaterialInventoryId)
        {
            try
            {
                var context = new NeevDatabaseContainer();
                context.DeleteRawMaterialInventory(rawMaterialInventoryId);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool DeleteRawMaterialInventoryItem(int rawMaterialInventoryTranId)
        {
            try
            {
                var context = new NeevDatabaseContainer();
                context.DeleteRawMaterialInventoryItem(rawMaterialInventoryTranId);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public List<RawMaterialInventoryItem> GetAllRawMaterialInventoryItems()
        {
            var context = new NeevDatabaseContainer();
            var productInventoryItems = (from rawMaterialInventoryItem in context.GetALLRawMaterialInventoryItems()
                                         select new RawMaterialInventoryItem { Id = rawMaterialInventoryItem.raw_material_inventory_trans_id, Name = rawMaterialInventoryItem.raw_material_name, Quantity = rawMaterialInventoryItem.quantity, Price = rawMaterialInventoryItem.price }).ToList<RawMaterialInventoryItem>();
            return productInventoryItems;
        }
    }
}
