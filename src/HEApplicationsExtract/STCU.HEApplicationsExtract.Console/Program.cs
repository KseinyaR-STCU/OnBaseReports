namespace STCU.HEApplicationsExtract.Console
{
    using System;
    using System.IO;
    using System.Reflection;
    using Serilog;
    using Core.Models;
    using Core.Services;
    using Hyland.Unity;
    using Hyland.Unity.UnityForm;
    using Excel = Microsoft.Office.Interop.Excel;

    class Program
    {
        static void Main(string[] args)
        {
            LogConfiguration.ConfigureSerilog();

            Log.Logger.Information("Extract Started! {0}", System.Configuration.ConfigurationManager.AppSettings["Application.Name"]);

            System.Data.DataTable table = QueryOnBase();

            Log.Logger.Information("Extract Ended! Unity Forms Extracted: {formsExtracted}", table.Rows.Count);

            if (table.Rows.Count > 0)
            {
                var tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());

                WriteExcelSpreadsheet(table, tempFile);
                StoreInOnBase(tempFile);

                File.Delete(tempFile);
            }

            Log.Logger.Information("Extract Stored in OnBase");

            LogConfiguration.FlushSerilog();
        }

        private static System.Data.DataTable QueryOnBase()
        {
            var extractCustomQueryName = System.Configuration.ConfigurationManager.AppSettings["extractCustomQueryName"];

            if (!Int32.TryParse(System.Configuration.ConfigurationManager.AppSettings["numberOfDaysExtracted"], out Int32 numberOfDaysExtracted))
            {
                Log.Logger.Warning("numberOfDaysExtracted in Config is invalid! Defaulting to 30 days!");
                numberOfDaysExtracted = 30;
            }

            if (!Int32.TryParse(System.Configuration.ConfigurationManager.AppSettings["maxQueryDocuments"], out Int32 maxQueryDocuments))
            {
                Log.Logger.Warning("maxQueryDocuments in Config is invalid! Defaulting to 10,000 documents!");
                maxQueryDocuments = 10000;
            }

            System.Data.DataTable table = SetupTable();

            using (Application obApp = OnBaseConnect())
            {
                DocumentQuery docQuery = obApp.Core.CreateDocumentQuery();

                CustomQuery extractCQ = obApp.Core.CustomQueries.Find(extractCustomQueryName);

                if (extractCQ != null)
                {
                    Log.Logger.Debug(String.Format("Custom Query Found: {0}", extractCustomQueryName));

                    docQuery.AddCustomQuery(extractCQ);
                    DateTime endTime = DateTime.Now;
                    DateTime startTime = endTime.AddDays(-1 * numberOfDaysExtracted);
                    docQuery.AddDateRange(startTime, endTime);

                    DocumentList docList = docQuery.Execute(maxQueryDocuments);

                    foreach (Document doc in docList)
                    {
                        Log.Logger.Debug("Doc: {docId} - {docName}", doc.ID, doc.Name);

                        Form form = obApp.Core.GetDocumentByID(doc.ID).UnityForm;

                        if (form != null)
                        {
                            HomeEquityApplication heApplication = new HomeEquityApplication(doc, form);
                            table.Rows.Add(heApplication.RowDataArray());
                        }

                    }
                }
                else
                {
                    Log.Logger.Error("Custom Query NOT Found: {extractCustomQueryName}", extractCustomQueryName);
                }
            }

            return table;

        }

        private static void WriteExcelSpreadsheet(System.Data.DataTable table, string tempFile)
        {
            Excel.Application appExcel = new Excel.Application();
            Excel.Workbook excelWorkbook = appExcel.Workbooks.Add(Type.Missing);
            Excel.Worksheet excelWorksheet;
            Excel.Range excelCellrange;

            try
            {
                appExcel.Visible = false;
                appExcel.DisplayAlerts = false;

                excelWorksheet = (Excel.Worksheet)excelWorkbook.ActiveSheet;
                excelWorksheet.Name = System.Configuration.ConfigurationManager.AppSettings["excelWorksheetName"];

                int rowcount = 1;

                foreach (System.Data.DataRow datarow in table.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= table.Columns.Count; i++)
                    {
                        // Write Column Names from table
                        if (rowcount == 2)
                        {
                            excelWorksheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                        }

                        // Write Row of Data from table
                        excelWorksheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                    }
                }

                excelCellrange = excelWorksheet.Range[excelWorksheet.Cells[1, 1], excelWorksheet.Cells[rowcount, table.Columns.Count]];


                excelWorksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, excelCellrange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "Table1";
                excelWorksheet.ListObjects["Table1"].TableStyle =
                                System.Configuration.ConfigurationManager.AppSettings["excelTableStyle"];

                excelCellrange.EntireColumn.AutoFit();

                excelWorkbook.SaveAs(tempFile, Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value, Missing.Value, false, false,
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlUserResolution,
                    false, Missing.Value, Missing.Value, Missing.Value);

            }
            catch (Exception ex)
            {
                Log.Logger.Error(ex.Message);
            }
            finally
            {
                excelWorkbook.Close();
                appExcel.Workbooks.Close();
                appExcel.Quit();

                excelWorksheet = null;
                excelCellrange = null;
                excelWorkbook = null;
            }
        }

        public static void StoreInOnBase(string tempFile)
        {
            using (Application obApp = OnBaseConnect())
            {
                // Get OnBase Storage Object
                Storage storage = obApp.Core.Storage;

                String docTypeName = System.Configuration.ConfigurationManager.AppSettings["outputDocumentType"];

                var documentType = obApp.Core.DocumentTypes.Find(docTypeName);

                if (documentType == null)
                {
                    //                    Console.WriteLine("Failed to find DocumentType: {0}", docTypeName);
                    Log.Logger.Error("Failed to find DocumentType: {@documentType}", docTypeName);
                }
                else
                {
                    // Set file type
                    var fileTypeName = "MS Excel Spreadsheet";
                    var fileExtension = @".xlsx";
                    FileType fileType = obApp.Core.FileTypes.Find(fileTypeName);

                    if (fileType == null)
                    {
                        //Console.WriteLine("Failed to find fileType: {0}", fileTypeName);
                        Log.Logger.Error("Failed to find FileType: {@fileTypeName}", fileTypeName);
                    }
                    else
                    {
                        // Read temporary file into MemoryStream
                        MemoryStream OBDocumentStream = new MemoryStream(File.ReadAllBytes(tempFile + fileExtension));

                        // Create Hyland PageData object for storage
                        PageData OBDocumentPageData = obApp.Core.Storage.CreatePageData(OBDocumentStream, fileExtension);

                        StoreNewDocumentProperties storeDocumentProperties = storage.CreateStoreNewDocumentProperties(documentType, fileType);

                        // Store Document in OnBase
                        Document OBDocumentNew = obApp.Core.Storage.StoreNewDocument(OBDocumentPageData, storeDocumentProperties);

                        if (OBDocumentNew == null)
                        {
                            //Console.WriteLine("Failed to store Document in OnBase!");
                            Log.Logger.Error("Failed to store Document in OnBase!");
                        }
                    }
                }
            }
        }

        private static System.Data.DataTable SetupTable()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            foreach (string headerName in HomeEquityApplication.rowHeaderArray())
            {
                table.Columns.Add(headerName, typeof(string));
            }
            return table;
        }

        private static Application OnBaseConnect()
        {
            Application app = null;

            try
            {
                app = Hyland.Unity.Application.Connect(STCU.OnBase.OnBaseProperties.Authentication);
            }
            catch (MaxLicensesException e)
            {
                //Console.WriteLine("Error: All available licenses have been consumed.");
                Log.Logger.Error("OnBaseConnect() error: {errorMessage}", e.Message);

                throw;
            }
            catch (SystemLockedOutException e)
            {
                //Console.WriteLine("Error: The system is currently in lockout mode.");
                Log.Logger.Error("OnBaseConnect() error: {errorMessage}", e.Message);
                throw;
            }
            catch (InvalidLoginException e)
            {
                //Console.WriteLine("Error: Invalid Login Credentials.");
                Log.Logger.Error("OnBaseConnect() error: {errorMessage}", e.Message);
                throw;
            }
            catch (AuthenticationFailedException e)
            {
                //Console.WriteLine("Error: NT Authentication Failed.");
                Log.Logger.Error("OnBaseConnect() error: {errorMessage}", e.Message);
                throw;
            }
            catch (MaxConcurrentLicensesException e)
            {
                //Console.WriteLine("Error: All concurrent licenses for this user group have been consumed.");
                Log.Logger.Error("OnBaseConnect() error: {errorMessage}", e.Message);
                throw;
            }
            catch (InvalidLicensingException e)
            {
                //Console.WriteLine("Error: Invalid Licensing.");
                Log.Logger.Error("OnBaseConnect() error: {errorMessage}", e.Message);
                throw;
            }

            if (app != null)
            {
                Log.Logger.Information("OnBaseConnect() success: {SessionID}", app.SessionID);
            }
            else
            {
                Log.Logger.Debug("OnBaseConnect() failed!");
            }

            return app;
        }
    }
}
