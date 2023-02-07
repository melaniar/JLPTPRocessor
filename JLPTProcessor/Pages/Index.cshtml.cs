using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;
using System.IO;

namespace JLPTProcessor.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;

        [BindProperty]
        public string ReportType { get; set; }

        [BindProperty]
        public string YearType { get; set; }
        
        [BindProperty]
        public int TestSiteCode { get; set; }
        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {

        }
        public void OnPostProcessReport()
        {
            string folderRoot = "C:\\Projects\\JLPT";
            if (!(System.IO.Directory.Exists(folderRoot)))
                System.IO.Directory.CreateDirectory(folderRoot);
            string outFile = "";
            if(ReportType == "Master")
                outFile = Path.Combine(folderRoot, "2023_July_Edmonton_Master.csv");
            else
                outFile = Path.Combine(folderRoot, "2023_July_Edmonton.csv");

            if (!System.IO.File.Exists(outFile))
                System.IO.File.Create(outFile).Close();

            if (ReportType == "Master")
                System.IO.File.AppendAllText(outFile, $"{getMasterReportHeaders()}.{Environment.NewLine}");

            // string reportType= Request.Form["ReportType"];
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open("all-data-groupize.xlsx", FileMode.Open, FileAccess.Read))
            {
               

                IExcelDataReader excelDataReader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = a => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                DataSet dataSet = excelDataReader.AsDataSet(conf);
                //counter Level
                int CountLvl1 = 0;
                int CountLvl2 = 0;
                int CountLvl3 = 0;
                int CountLvl4 = 0;
                int CountLvl5 = 0;

                DataRowCollection rows = dataSet.Tables["Attendees"].Rows;
                //get the status of the registration and email address 
                //for test data only process those "confirmed" status
                for (int i = 0; i < rows.Count; i++)//foreach (DataRow row in rows)
                {
                    string result = "";
                  
                    if (i >= 5)//ignore header
                    {
                        var cols = rows[i].ItemArray.ToList();
                        if (cols[6]?.ToString() == "Confirmed") //RegistrationStatus
                        {
                            string usrEmail = cols[3].ToString();
                            if (ReportType == "Master")
                            {      
                                result = processMasterQuestionaire(dataSet, usrEmail);
                            }
                            else
                                processRegistrationQuestionaire(dataSet, usrEmail);
                            break;
                        }
                    }
                }
            }
        }
        private string processMasterQuestionaire(DataSet dataSet, string email) //master Report
        {
            DataRowCollection rows = dataSet.Tables["Questions"].Rows;

            string result = "";
            int startRow = 5; //NOT include the headers
            if (ReportType == "Master")
            {
                //Read the header too == include header
                startRow = 4;
            }

            for (int i = 0; i < rows.Count; i++)//foreach (DataRow row in rows)
            {
                if (i >= startRow)//start redaing from line 6 -- headers
                {
                    var cols = rows[i].ItemArray.ToList();
                   
                }
            }
        }

        private void processRegistrationQuestionaire(DataSet dataSet, string email)
        {
            DataRowCollection rows = dataSet.Tables["Questions"].Rows;


            int startRow = 5; //NOT include the headers
            //if (ReportType == "Master")
            //{
            //    //Read the header too == include header
            //    startRow = 4;
            //}

            for (int i = 0; i < rows.Count; i++)//foreach (DataRow row in rows)
            {
                if (i >= startRow)//start redaing from line 6 -- headers
                {
                    var cols = rows[i].ItemArray.ToList();
                    //rowDataList = item.ItemArray.ToList(); //list of each rows
                    //allRowsList.Add(rowDataList); //adding the above list of each row to another list
                }
            }
        }

        private string GetTestLevel(DataSet dataSet, string email)
        {
            DataRowCollection rows = dataSet.Tables["Orders"].Rows;
            string testL = "";
            for (int i = 0; i < rows.Count; i++)//foreach (DataRow row in rows)
            {
                if (i >= 5)//start redaing from line 6 -- headers
                {
                    var cols = rows[i].ItemArray.ToList();

                    if (cols[3] == email)
                    { //4th col == Email
                        string testLevel = cols[5]  != null ? (string)cols[5] : null; //ITem Name /ticket's name
                        if (testLevel.Contains("N1"))
                            testL = ReportType == "Master" ? "N1" : "1";
                        else if (testLevel.Contains("N2"))
                            testL = ReportType == "Master" ? "N2" : "2";
                        else if (testLevel.Contains("N3"))
                            testL = ReportType == "Master" ? "N3" : "3";
                        else if (testLevel.Contains("N4"))
                            testL = ReportType == "Master" ? "N4" : "4"; 
                        else if (testLevel.Contains("N5"))
                            testL = ReportType == "Master" ? "N5" : "5"; 

                        break;
                    }
                }
            }
            return testL;
        }

        private string getRegistrationReportHeaders()
        {
            return "Test Level,Year Code,Test Site,Test Level,Sequence No.,Full Name,Sex,DOB Year,DOB Month,DOB Day,Password,Native Language,Learning Place,Reason,	Occupation,	Occupation Details,	Media,	With Teachers,	With Friends,	With Family,With Supervisor,With Colleagues,With Customers,Times Taking N1,	Times Taking N2,Times Taking N3,Times Taking N4,Times Taking N5,Latest N1 Result,Latest N2 Result,Latest N3 Result,	Latest N4 Result,Latest N5 Result";
        }

        private string getMasterReportHeaders()
        {
            return "Sequence No.,Level, First Name, Last Name,DOB Year,DOB Month,DOB Day,Street,City,Province,Country,Postal,Email,	Affiliation, Registration Type,Times Taking N1,	Times Taking N2,Times Taking N3,Times Taking N4,Times Taking N5,Latest N1 Result,Latest N2 Result,Latest N3 Result,	Latest N4 Result,Latest N5 Result";
        }
    }
}