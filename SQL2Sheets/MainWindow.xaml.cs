using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;
using System.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace SQL2Sheets
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static MainWindow AppWindow;
        private List<DataProject> dpList = new List<DataProject>();
        //TODO need to implement the backgroundworker thread to keep form responsive during database work
        //  With backgroundworker I need to implement worker_DoWork method and worker_RunWorkerCompleted method
        //BackgroundWorker worker;
        public MainWindow()
        {
            InitializeComponent();
            //set AppWindow so I can access it from other classes
            AppWindow = this;
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            //TODO need a method to check for valid data in the form
            SetDebugScreen(ProjectName.Text + SheetID.Text + ConnectionString.Text + SqlColumns.Text + SelectStatement.Text);

            // Need to figure out how to dynamically make more program objects based on user entries
            // Need to make a create object method and a destroy object method.
            // Now that I have multiple objects created that could run indefinetly now I need threads.

            ActivityTab.IsSelected = true;
            CreateDataProject();
        }

        private void CreateDataProject()
        {
            //Problem I dont know if the object is valid! If user puts in bad data program will crash.
            // Yes I know this is rediculous. I should convert the object into a JSON string and pass that instead.
            dpList.Add(new DataProject(ProjectName.Text, SheetID.Text, ConnectionString.Text, SqlColumns.Text, SelectStatement.Text));
        }

        private void DestroyDataProject()
        {
            //dpList.Remove();
        }

        public void SetDebugScreen(string data)
        {
            //How do I get the DataProject class to use this method?
            int tl = TextOutput.Text.Length + data.Length;
            TextOutput.Text = "{" + tl.ToString()+"} " + data + "\n" + TextOutput.Text;
        }

    }

    //Do I really need to do this? Or should all of the API and SQL stuff just be methods of the main Window?
    class DataProject
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        private string name;
        private string sID;
        private string connection;
        private string header;
        private string select;
        private string dataRange = "A1:ZZ";

        public DataProject(string projectName, string sheetID, string connectionString, string sqlColumns, string selectStatement )
        {
            //TODO add the initialization here then call mainRun
            // change the crazy string arguments into a single JSON object
            this.name = projectName;
            this.sID = sheetID;
            this.connection = connectionString;
            this.header = sqlColumns;
            this.select = selectStatement;

            this.MainRun();
        }


        //TODO start moving this crap to methods or other objects.
        public void MainRun()
        {
            //TODO Method to set credentials
            //TODO MEthod to connect and pull data from SQL
            //TODO Method to clear the sheet
            //TODO Method to write the header and data

            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                credPath = System.IO.Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Credential file saved to:\r\n" + credPath + "\r\n");                
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            // here is the actual data to be sent to sheet
            IList<object> headerList;
            headerList = header.Split(',');
            
            ValueRange valueRange = new ValueRange { MajorDimension = "ROWS" };
            
            SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Clear the Sheet");

            //API method to clear the sheet
            ClearValuesRequest clearValuesRequest = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest cr = service.Spreadsheets.Values.Clear(clearValuesRequest, sID, dataRange);
            // TODO add a try catch statement
            ClearValuesResponse clearResponse = cr.Execute();
            
            SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Delete all rows in Sheet");

            //API method to batch update
            DimensionRange dr = new DimensionRange
            {
                Dimension = "ROWS",
                StartIndex = 1,
                SheetId = 1809337217 //this is a problem
            };

            DeleteDimensionRequest ddr = new DeleteDimensionRequest() { Range = dr };

            Request r = new Request { DeleteDimension = ddr };

            //THIS IS FOR deleteDimension { "requests": [{ "deleteDimension": { "range": { "sheetId": 1809337217, "startIndex": 1}} }  ]};
            List<Request> batchRequests = new List<Request>() { r };

            BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest() { Requests = batchRequests };

            SpreadsheetsResource.BatchUpdateRequest bRequest = service.Spreadsheets.BatchUpdate(requestBody, sID);
            BatchUpdateSpreadsheetResponse busr = bRequest.Execute();
            
            SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Write the header to the Sheet");

            //API method to update the sheet 
            //THis performs a batch request to write the headerlist
            //TODO add the SQL data to this update request
            //THIS IS FOR DATA   { "requests": [{ "valueInputOption": "RAW", "data": [{ "majorDimension": "ROWS","range": "","values": [],[],[] }] }] }
            valueRange.Values = new List<IList<object>> { headerList };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, sID, dataRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result;

            //TODO when I get both header and data this execute needs to be in a throttled, thread to update at a specified rate
            result = update.Execute();
            
            SqlConnection sqlConnection = new SqlConnection(connection);

            SqlCommand cmd = new SqlCommand
            {
                Connection = sqlConnection,
                CommandTimeout = 60,
                CommandType = CommandType.Text,
                CommandText = select
            };

            try
            {
                SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Open the SQL connection");
                sqlConnection.Open();
            }
            catch (Exception e)
            {
                SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Whoops we cannot connect to SQL - exception " + e.HResult.ToString());
                //Environment.Exit(1);
            }
            
            SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Please wait while reading data from SQL server");
            SqlDataReader reader;
            reader = cmd.ExecuteReader();

            // Data is accessible through the DataReader object here.
            ValueRange valueDataRange = new ValueRange() { MajorDimension = "ROWS" };

            var dataList = new List<object>();
            valueDataRange.Values = new List<IList<object>> { dataList };

            //API to append data to sheet. This does not use the batch
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            SpreadsheetsResource.ValuesResource.AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueDataRange, sID, dataRange);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

            if (reader.HasRows)
            {
                Object[] colValues = new Object[reader.FieldCount];

                int throttleCount = 0;
                int cnt = 0;
                while (reader.Read())
                {
                    //This logic is flawed. If we get hit by the quota then the data row gets lost the next time this runs.
                    dataList.Clear();
                    for (int i = 0; i < reader.GetValues(colValues); i++)
                    {
                        dataList.Add(colValues[i]);
                    }

                    try
                    {
                        //This is the GOOGLE query Throttle they only allow 500 writes per 100 sec
                        System.Threading.Thread.Sleep(20);
                        AppendValuesResponse appendValueResponse = appendRequest.Execute();
                        SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Writing to Sheet: row{" + cnt.ToString() + "}");
                    }
                    catch (Exception e)
                    {
                        SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Whoa buddy slowdown! Exception " + e.HResult.ToString());
                        System.Threading.Thread.Sleep(3000);
                        throttleCount++;
                    }
                    cnt++;
                }
            }
            else
            {
                SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("No rows found");
            }
            
            SQL2Sheets.MainWindow.AppWindow.SetDebugScreen("Close reader and SQL");
            reader.Close();
            sqlConnection.Close();
        }
    }
}
