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
        private List<DataProject> dpList = new List<DataProject>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {
            // TODO this is place for logic????
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
            // ALSO. Yes I know this is rediculous. I should convert the object into a JSON string and pass that instead.
            dpList.Add(new DataProject(ProjectName.Text,SheetID.Text,ConnectionString.Text,SqlColumns.Text,SelectStatement.Text));
        }

        private void DestroyDataProject()
        {

            //dpList.Remove();
        }

        public void SetDebugScreen(string data)
        {
            int tl = TextOutput.Text.Length + data.Length;
            TextOutput.Text = tl.ToString() + data + "\n" + TextOutput.Text;
        }

    }

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
                Console.WriteLine("Credential file saved to:\r\n" + credPath + "\r\n");
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Prints the data from a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1c54Cy_B43h5-nmE7r6Slvj2w8Pl0XFxgaWpTxO9s9So/edit#gid=0

            // Define request parameters.
            String spreadsheetId = "1c54Cy_B43h5-nmE7r6Slvj2w8Pl0XFxgaWpTxO9s9So";

            // here is the actual data to be sent to sheet
            List<object> headerList = new List<object>() {
            "ID","DTStamp","DTShiftStart","ModelNbr","SerialNbr","PassFail","LineNbr","ShiftNbr","Computer","Word40","Word41","Word42"
            ,"Word43","Word44","Word45","Word46","Word47","Word48","Word49","Word50","Word51","Word52","Word53","Word54","Word55","Word56"
            ,"Word57","Word58","Word59","Word60","Word61","Word62","Word63","Word64","Word65","Word66","Word67","Word68","Word69","Word70"
            ,"Word71","Word72","Word73","Word74","Word75","Word76","Word77","Word78","Word79","Word80"};

            //var dataList = new List<object>();

            //Write some data in the very first active sheet
            String dataRange = "A1:ZZ";
            ValueRange valueRange = new ValueRange { MajorDimension = "ROWS" };

            Console.WriteLine("Clear the Sheet");

            //API method to clear the sheet
            ClearValuesRequest clearValuesRequest = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest cr = service.Spreadsheets.Values.Clear(clearValuesRequest, spreadsheetId, dataRange);
            // TODO add a try catch statement
            ClearValuesResponse clearResponse = cr.Execute();

            Console.WriteLine("Delete all rows in Sheet");

            //API method to batch update
            DimensionRange dr = new DimensionRange
            {
                Dimension = "ROWS",
                StartIndex = 1,
                SheetId = 1809337217 //this is a problem
            };

            DeleteDimensionRequest ddr = new DeleteDimensionRequest() { Range = dr };

            Request r = new Request { DeleteDimension = ddr };

            // { "requests": [{ "deleteDimension": { "range": { "sheetId": 1809337217, "startIndex": 1}} }  ]};
            List<Request> batchRequests = new List<Request>() { r };

            BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest() { Requests = batchRequests };

            SpreadsheetsResource.BatchUpdateRequest bRequest = service.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);
            BatchUpdateSpreadsheetResponse busr = bRequest.Execute();

            Console.WriteLine("Write the Headers to the Sheet");

            //API method to update the sheet
            valueRange.Values = new List<IList<object>> { headerList };
            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, dataRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result;
            result = update.Execute();
            
            //SqlConnection sqlConnection = new SqlConnection("Data Source=tul-mssql;Initial Catalog=Division;User ID=tqisadmin;Password=admin2k");
            SqlConnection sqlConnection = new SqlConnection("Data Source=TULLCND62158MM\\SQLEXPRESS;Initial Catalog=TulQual;User ID=tqisadmin;Password=admin2k");

            SqlCommand cmd = new SqlCommand
            {
                Connection = sqlConnection,
                CommandTimeout = 60,
                CommandType = CommandType.Text,
                CommandText = "SELECT TOP 10 [Date],[Shift],[Line],[Production],[Repair Bay Units],[Test Pulls],[Repair Bay Units w/o Test Pulls] " +
                    ",[CAL Units Inspected],[QAL Units Inspected],[CAL Units Nonconf],[QAL Units Nonconf],[Hold Order Units],[CAL/QAL Line/FDC Audit Units] " +
                    ",[CAL/QAL Line/FDC Audit Units Nonconf],[CAL/QAL Units Inspected],[CAL/QAL Units Nonconf] " +
                    "FROM [TulQual].[dbo].[RTYtblDataEntry]"
                /*
                CommandText = "SELECT TOP 100 [ID],[DTStamp],[DTShiftStart],[ModelNbr],[SerialNbr],[PassFail],[LineNbr],[ShiftNbr],[Computer],[Word40],[Word41],[Word42]" +
                    ",[Word43],[Word44],[Word45],[Word46],[Word47],[Word48],[Word49],[Word50],[Word51],[Word52],[Word53],[Word54],[Word55],[Word56]" +
                    ",[Word57],[Word58],[Word59],[Word60],[Word61],[Word62],[Word63],[Word64],[Word65],[Word66],[Word67],[Word68],[Word69],[Word70]" +
                    ",[Word71],[Word72],[Word73],[Word74],[Word75],[Word76],[Word77],[Word78],[Word79],[Word80] " +
                    "FROM[Division].[dbo].[asyTestRecords] where LineNbr = 2 and computer = 'LN' and dtstamp > '2/1/2018 5:00' order by dtstamp desc"
                 */
            };

            try
            {
                Console.WriteLine("Open the SQL connection");
                sqlConnection.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine("Whoops we cannot connect to SQL - exception {0}", e.HResult);
                //Environment.Exit(1);
            }

            Console.WriteLine("Please wait while reading data from SQL");
            SqlDataReader reader;
            reader = cmd.ExecuteReader();

            // Data is accessible through the DataReader object here.
            ValueRange valueDataRange = new ValueRange() { MajorDimension = "ROWS" };

            var dataList = new List<object>();
            valueDataRange.Values = new List<IList<object>> { dataList };

            //API to append data to sheet
            SpreadsheetsResource.ValuesResource.AppendRequest appendRequest = service.Spreadsheets.Values.Append(valueDataRange, spreadsheetId, dataRange);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            appendRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

            if (reader.HasRows)
            {
                //Console.WriteLine("{0}",reader.FieldCount);
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
                        Console.WriteLine("Writing to Sheet: row{0}", cnt);
                    }
                    catch (Exception e)
                    {

                        Console.WriteLine("Whoa buddy slowdown {0} exception {1}", throttleCount, e.HResult);
                        System.Threading.Thread.Sleep(3000);
                        throttleCount++;
                    }
                    cnt++;
                }
            }
            else
            {
                Console.WriteLine("No rows found.");
            }

            // sit here and wait for a while
            Console.WriteLine("Done waiting to close reader and SQL");
            System.Threading.Thread.Sleep(3000);

            reader.Close();
            sqlConnection.Close();
        }
    }
}
