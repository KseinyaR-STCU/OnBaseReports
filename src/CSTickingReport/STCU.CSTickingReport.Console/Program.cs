namespace STCU.CSTickingReport.Console
{
    using System;
    using System.IO;
    using Core.Model;
    using Excel = Microsoft.Office.Interop.Excel;
    using Hyland.Unity;
    using Services;
    using Serilog;
    using System.Reflection;

    class Program
    {
        private static DateTime docDate = new DateTime();

        static void Main(string[] args)
        {

            LogConfiguration.ConfigureSerilog();
            Log.Logger.Information("Start Processing Ticking Report!");

            System.Data.DataTable table = QueryOnBase();

            if (table.Rows.Count > 0)
            {
                var tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());

                WriteExcelSpreadsheet(table, tempFile);
                StoreInOnBase(tempFile);

                File.Delete(tempFile);

            }

            Log.Logger.Information("Finished Processing Ticking Report! Transaction Count: {transactionCount}", table.Rows.Count);

            LogConfiguration.FlushSerilog();
        }

        private static System.Data.DataTable QueryOnBase()
        {
            System.Data.DataTable table = SetupTable();

            int transactionSub = -1;

            using (Application obApp = OnBaseConnect())
            {
                DocumentQuery docQuery = obApp.Core.CreateDocumentQuery();

                String docTypeName = System.Configuration.ConfigurationManager.AppSettings["inputDocumentType"];

                var documentType = obApp.Core.DocumentTypes.Find(docTypeName);

                DateTime endTime = DateTime.Now;
                DateTime startTime = new DateTime(endTime.Year, endTime.Month, endTime.Day - (endTime.DayOfWeek.Equals(DayOfWeek.Monday) ? 3 : 1), 0, 0, 0);

                docQuery.AddDocumentType(documentType);
                docQuery.AddDateRange(startTime, endTime);

                long maxDocuments = 10;
                using (QueryResult queryResults = docQuery.ExecuteQueryResults(maxDocuments))
                {
                    Document doc = null;
                    foreach (QueryResultItem queryResultItem in queryResults.QueryResultItems)
                    {
                        doc = queryResultItem.Document;
                        break;
                    }

                    if (doc == null)
                    {
                        Log.Logger.Information("Document Not Found! Date Range: {startTime} - {endTime}", startTime, endTime);
                    }
                    else
                    {
                        docDate = doc.DocumentDate;

                        using (PageData pd = obApp.Core.Retrieval.Text.GetDocument(doc.DefaultRenditionOfLatestRevision))
                        {
                            using (StreamReader streamReader = new StreamReader(pd.Stream))
                            {
                                String line = null;
                                Transaction trans = new Transaction();
                                while ((line = streamReader.ReadLine()) != null)
                                {
                                    if (transactionSub == -1 && line.StartsWith("153"))
                                    {
                                        transactionSub++;
                                        trans.TransLine1(line);
                                    }
                                    else if (transactionSub == 0)
                                    {
                                        transactionSub++;
                                        trans.TransLine2(line);
                                    }
                                    else if (transactionSub == 1)
                                    {
                                        transactionSub++;
                                        trans.TransLine3(line);
                                    }
                                    else if (transactionSub == 2)
                                    {
                                        transactionSub++;
                                        trans.TransLine4(line);

                                        // Reset Transaction Sub Counter
                                        transactionSub = -1;

                                        // Add Data to Table
                                        table.Rows.Add(trans.RowDataArray());
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return table;
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
                        Log.Logger.Error("Failed to find FileType: {@fileTypeName}", fileTypeName);
                    }
                    else
                    {
                        MemoryStream OBDocumentStream = new MemoryStream(File.ReadAllBytes(tempFile + fileExtension));

                        // Create Hyland PageData object for storage
                        PageData OBDocumentPageData = obApp.Core.Storage.CreatePageData(OBDocumentStream, fileExtension);

                        StoreNewDocumentProperties storeDocumentProperties = storage.CreateStoreNewDocumentProperties(documentType, fileType);

                        // Store Document in OnBase
                        Document OBDocumentNew = obApp.Core.Storage.StoreNewDocument(OBDocumentPageData, storeDocumentProperties);

                        if (OBDocumentNew == null)
                        {
                            Log.Logger.Error("Failed to store Document in OnBase!");
                        }
                        else
                        {
                            //Set Keywords
                            KeywordType ReportIdKWType = obApp.Core.KeywordTypes.Find("Report ID");
                            Keyword ReportIdKW = ReportIdKWType.CreateKeyword("00");
                            KeywordRecord ReportIdKWRecord = OBDocumentNew.KeywordRecords.Find(ReportIdKWType);
                            if (ReportIdKWRecord == null)
                            {
                                KeywordModifier ReportIdKeywordModifier = OBDocumentNew.CreateKeywordModifier();
                                ReportIdKeywordModifier.AddKeyword(ReportIdKW);
                                ReportIdKeywordModifier.ApplyChanges();
                            }
                            else
                            {
                                Log.Logger.Error("Keyword not null!");
                            }

                            // Stored Document will have today's DateTime; Update with DateTime scraped from Evidence Summary
                            DocumentPropertiesModifier OBDocModifier = OBDocumentNew.CreateDocumentPropertiesModifier();
                            OBDocModifier.DocumentDate = docDate;
                            OBDocModifier.ApplyChanges();

                            Log.Logger.Information("Storage to OnBase succesful. Doc Handle: {OBDocHandle}", OBDocumentNew.ID);
                        }
                    }
                }
            }

        }

        private static System.Data.DataTable SetupTable()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            foreach (string headerName in Transaction.rowHeaderArray())
            {
                table.Columns.Add(headerName, typeof(string));
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
                Log.Logger.Error("Excel Storage Error: {excelError}", ex.Message);
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

        private static Application OnBaseConnect()
        {
            Application app = null;

            try
            {
                app = Hyland.Unity.Application.Connect(STCU.OnBase.OnBaseProperties.Authentication);
            }
            catch (MaxLicensesException e)
            {
                Log.Logger.Error(e, "Error: All available licenses have been consumed.");
                throw;
            }
            catch (SystemLockedOutException e)
            {
                Log.Logger.Error(e, "Error: The system is currently in lockout mode.");
                throw;
            }
            catch (InvalidLoginException e)
            {
                Log.Logger.Error(e, "Error: Invalid Login Credentials.");
                throw;
            }
            catch (AuthenticationFailedException e)
            {
                Log.Logger.Error(e, "Error: NT Authentication Failed.");
                throw;
            }
            catch (MaxConcurrentLicensesException e)
            {
                Log.Logger.Error(e, "Error: All concurrent licenses for this user group have been consumed.");
                throw;
            }
            catch (InvalidLicensingException e)
            {
                Log.Logger.Error(e, "Error: Invalid Licensing.");
                throw;
            }

            if (app != null)
            {
                Log.Logger.Debug("Connection Successful. Connection ID: {@SessionID}", app.SessionID);
            }

            return app;
        }

    }
}
