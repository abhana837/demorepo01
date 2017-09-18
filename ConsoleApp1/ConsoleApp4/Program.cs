using Amazon;
using Amazon.DynamoDBv2;
using Amazon.DynamoDBv2.DocumentModel;
using Amazon.DynamoDBv2.Model;
using Amazon.Runtime;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;
using  Excel = Microsoft.Office.Interop.Excel;
using Amazon.DynamoDBv2.DataModel;
using System.Configuration;

namespace ConsoleApp4
{
    class Program
    {
        static void Main(string[] args)
        {
            //GetExcelData();
            ShowManu();
            string input = Console.ReadLine();

            do
            {
                OptionHub(input);
                input = Console.ReadLine();
                
            } while (input != "X");

            //InsertDocInDynamoDB();
            Console.ReadLine();

        }

        static AmazonDynamoDBClient client;

        private static void OptionHub(string opt)
        {
            switch (opt)
            {
                case "R":
                    ReadFromDynamoDB("GI");
                    break;

                case "W":
                    InsertDocInDynamoDB();
                    break;

                case "X":
                    break;


                default:
                    Console.WriteLine("Incorrect Option.");
                    break;
            }
        }

        private static void ShowManu()
        {
            Console.WriteLine("hiii.. Please enter below option");
            Console.WriteLine("Enter 'R' for Read");
            Console.WriteLine("Enter 'W' for Write");
            Console.WriteLine("Enter 'X' for Exit");
            Console.WriteLine("----------------------------------");
        }

        private static AmazonDynamoDBClient GetAWSClient()
        {
            LogEvent("Connecting AWSDynamoDB...");
            string AWSKey = ConfigurationSettings.AppSettings["AWSKey"];

            string AWSSecKey = ConfigurationSettings.AppSettings["AWSSecKey"];
            
            AmazonDynamoDBConfig clientConfig = new AmazonDynamoDBConfig();
            clientConfig.RegionEndpoint = RegionEndpoint.USEast1;
            var credentials = new BasicAWSCredentials(AWSKey, AWSSecKey);
            AmazonDynamoDBClient client = new AmazonDynamoDBClient(credentials, clientConfig);
            return client;
        }

        private static void InsertDocInDynamoDB()
        {
            
            DataTable dt = GetExcelData();
            client = GetAWSClient();
            try
            {                
                string tableName = "Zeus";

                Table tbl = Table.LoadTable(client, tableName);
                
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var rowDoc = new Document();
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        //if (dt.Columns[j].ToString() == "Id")
                        //    rowDoc[dt.Columns[j].ToString()] = Convert.ToInt32(dt.Rows[i][j]);
                        //else
                            rowDoc[dt.Columns[j].ToString()] = dt.Rows[i][j].ToString();
                    }
                    tbl.PutItem(rowDoc);
                    rowDoc = null;
                }

                LogEvent(dt.Rows.Count +" rows inserted in AWSDynamoDB");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                dt = null;
                client.Dispose();
                
            }
        }

        private static DataTable GetExcelData()
        {
            try
            {
                LogEvent("Reading Excel...");
                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\admin\Documents\abhbm\testData\ABOProcessFinal2.xlsx");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                DataTable dataTable = new DataTable();

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 1; i <= rowCount; i++)
                {
                    string[] rowDataArray = new string[xlRange.Columns.Count];
                    for (int j = 1; j <= colCount; j++)
                    {

                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            if (i == 1)
                            {
                                dataTable.Columns.Add(xlRange.Cells[i, j].Value2.ToString());
                            }
                            else
                            {
                                if (j==2)
                                    rowDataArray[j - 1] = xlRange.Cells[i, j].Value2.ToString().ToLower();
                                else
                                    rowDataArray[j - 1] = xlRange.Cells[i, j].Value2.ToString();

                                if (xlRange.Columns.Count == j)
                                {
                                    dataTable.Rows.Add(rowDataArray);
                                }
                            }

                        }

                    }
                    rowDataArray = null;
                }

                

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                LogEvent( dataTable.Rows.Count +" rows found.");
                return dataTable;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private static void LogEvent(string message)
        {
            Console.WriteLine(message);
        }

        private static void ReadFromDynamoDB(string Key) {

            //DynamoDBContext context = new DynamoDBContext(GetAWSClient());
            //QueryRequest reqQuery = new QueryRequest();
            //reqQuery.TableName = "Zeus";
            
            //Table t = Table.LoadTable(client, "Zeus");
            //Primitive primitive = new Primitive();
            client = GetAWSClient();
            // Define scan conditions
            Dictionary<string, Condition> conditions = new Dictionary<string, Condition>();

            // Title attribute should contain the string "Adventures"
            Condition titleCondition = new Condition();
            titleCondition.ComparisonOperator = ComparisonOperator.CONTAINS;
            titleCondition.AttributeValueList.Add(new AttributeValue { S = "GI" });
            conditions["Category"] = titleCondition;
            // Define marker variable
            Dictionary<string, AttributeValue> startKey = null;

            //do
            //{
                // Create Scan request
                ScanRequest request = new ScanRequest
                {
                    TableName = "Zeus",
                    ExclusiveStartKey = startKey,
                    ScanFilter = conditions
                };

                // Issue request
                ScanResult result = client.Scan(request).ScanResult;

                // View all returned items
                List<Dictionary<string, AttributeValue>> items = result.Items;
                foreach (Dictionary<string, AttributeValue> item in items)
                {
                    Console.WriteLine("Item:");
                    foreach (var keyValuePair in item)
                    {
                        Console.WriteLine("{0} : S={1}, N={2}, SS=[{3}], NS=[{4}]",
                            keyValuePair.Key,
                            keyValuePair.Value.S,
                            keyValuePair.Value.N,
                            string.Join(", ", keyValuePair.Value.SS ?? new List<string>()),
                            string.Join(", ", keyValuePair.Value.NS ?? new List<string>()));
                    }
                }

                // Set marker variable
            //    startKey = result.LastEvaluatedKey;
            //} while (startKey != null);

            // Pages attributes must be greater-than the numeric value "200"
            //Condition pagesCondition = new Condition();
            //pagesCondition.ComparisonOperator = ComparisonOperator.GT; ;
            //pagesCondition.AttributeValueList.Add(new AttributeValue { N = "200" });
            //conditions["Pages"] = pagesCondition;

            //var request = new ScanRequest
            //{
            //    TableName = "Zeus",

            //    KeyConditions = new Dictionary<string, Condition>
            //    {
            //        { "Category", new Condition()
            //            {
            //                ComparisonOperator = ComparisonOperator.EQ,
            //                AttributeValueList = new List<AttributeValue>
            //                {
            //                    new AttributeValue { S = "GI" }
            //                }
            //            }
            //        }
            //    }

            //};
            //var response = client.Query(request);

            //foreach (var item in response.Items)
            //{
            //    // Write out the first page of an item's attribute keys and values.
            //    // PrintItem() is a custom function.
            //    //PrintItem(item);
            //    Console.WriteLine("=====");
            //}
        }
    }
}
