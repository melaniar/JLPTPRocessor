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
        //counter Level
        private int CountLvl1 = 0;
        private int CountLvl2 = 0;
        private int CountLvl3 = 0;
        private int CountLvl4 = 0;
        private int CountLvl5 = 0;

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
            //Generate CVS files similar to old subline
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
                System.IO.File.AppendAllText(outFile, $"{getMasterReportHeaders()}.{Environment.NewLine}"); //print the header
            else
                System.IO.File.AppendAllText(outFile, $"{getRegistrationReportHeaders()}.{Environment.NewLine}"); //print the header

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
                
                DataRowCollection rows = dataSet.Tables["Attendees"].Rows;
                //get the status of the registration and email address 
                //for test data only process those "confirmed" status
                for (int i = 0; i < rows.Count; i++)//foreach (DataRow row in rows)
                {
                    string result = "";
                  
                    if (i >= 5)//ignore header
                    {
                        var cols = rows[i].ItemArray.ToList();
                        if (cols[6]?.ToString() == "confirmed") //RegistrationStatus
                        {
                            string usrEmail = cols[3].ToString();
                            if (ReportType == "Master")
                            {
                                result = processMasterQuestionaire(dataSet, usrEmail);
                                System.IO.File.AppendAllText(outFile, $"{result}.{Environment.NewLine}"); //print the header
                            }
                            else
                            {
                                result = processRegistrationQuestionaire(dataSet, usrEmail);
                                System.IO.File.AppendAllText(outFile, $"{result}.{Environment.NewLine}"); //print the header
                            }
                            
                        }
                    }
                }
            }
        }
        private string processRegistrationQuestionaire(DataSet dataSet, string email) //Registration Report
        {
            string testLevel = GetTestLevel(dataSet, email);
            int sqNum = 0;
            if (testLevel.Contains("1"))
            {
                CountLvl1++;
                sqNum = CountLvl1;
            }
            else if (testLevel.Contains("2"))
            {
                CountLvl2++;
                sqNum = CountLvl2;
            }
            else if (testLevel.Contains("3"))
            {
                CountLvl3++;
                sqNum =CountLvl3;
            }
            else if (testLevel.Contains("4"))
            {
                CountLvl4++;
                sqNum =CountLvl4;
            }
            else if (testLevel.Contains("5"))
            {
                CountLvl5++;
                sqNum = CountLvl5;
            }



            DataRowCollection rows = dataSet.Tables["Questions"].Rows;

            string result = "";
            string seqFormat = "0000"; //seguence number to be padd with leading 0 with 4 digits length
            result = "\"" + testLevel + "\",\"" + YearType + "\",\"" + TestSiteCode + "\",\"" + testLevel + "\",\"" + sqNum.ToString(seqFormat) + "\",";

            for (int i = 5; i < rows.Count; i++)//foreach (DataRow row in rows)
            {
                var cols = rows[i].ItemArray.ToList();
                if (cols[3] == email)//start redaing from line 6 -- headers
                {
                    result += extractRegistrationData(rows[i], email);
                    break;
                }


            }

            return result;
        }

        private string processMasterQuestionaire(DataSet dataSet, string email) //master Report
        {
            string testLevel = GetTestLevel(dataSet, email);
            int sqNum = 0;
            if (testLevel.Contains("1"))
            {
                CountLvl1++;
                sqNum = (10000 + CountLvl1);
            }
            else if (testLevel.Contains("2"))
            {
                CountLvl2++;
                sqNum = (20000 + CountLvl2);
            }
            else if (testLevel.Contains("3"))
            {
                CountLvl3++;
                sqNum = (30000 + CountLvl3);
            }
            else if (testLevel.Contains("4"))
            {
                CountLvl4++;
                sqNum = (40000 + CountLvl4);
            }
            else if (testLevel.Contains("5"))
            {
                CountLvl5++;
                sqNum = (50000 + CountLvl5);
            }

           
            DataRowCollection rows = dataSet.Tables["Questions"].Rows;

            string result = "";
            result = sqNum.ToString() + "," + testLevel + ",";


            int startRow = 5; //NOT include the headers
            if (ReportType == "Master")
            {
                //Read the header too == include header
                startRow = 4;
            }

            for (int i = 0; i < rows.Count; i++)//foreach (DataRow row in rows)
            {
                var cols = rows[i].ItemArray.ToList();
                if (cols[3] == email)//start redaing from line 6 -- headers
                {
                    result += extractDataMaster(rows[i], email);
                    break;
                }
                   
                
            }

            return result;
        }

        private string extractDataMaster(DataRow row,  string email)
        {
            string result = "";
            string street="", city="", prov="", country="", postal="", phone="", registrationType="pay";
            string afiliation="";
            string[] timeTakingTest = {"0", "0" , "0" , "0" , "0" }; //new string[5]; //contain time of taking test l1-5 -- in order
            string[] resultTakingTest = { "0", "0", "0", "0", "0" }; //new string[5]; //contain lat result of taking test l1-5 -- in order

            var cols = row.ItemArray.ToList();
            for (int j = 0; j < cols.Count; j++)
            {
                if (j == 0) //col 1st = FName
                {
                    string fname = cols[j].ToString();
                    result += "\""+ fname.ToUpper() + "\",\"";
                }
                else if (j == 2) //col 3rd LName
                {
                    string lname = cols[j].ToString();
                    result += lname.ToUpper() + "\",";
                }
                /*else if (j == 3) //col 3rd Email
                {
                    tempEemail = cols[j].ToString(); //will position it later
                }*/
                /*else if (j == 4) //col Gender
               {
                   tempEemail = cols[j].ToString(); //will position it later
               }*/
                else if (j == 5) //col 6rd DOB
                {
                    string[] dob = cols[j].ToString().Split("-");

                    result += "\"" + dob[0] + "\"," + dob[1] + "," + dob[2] + ",";
                }
                else if (j == 14) //col 15 --Afiliation
                {
                    afiliation = cols[j].ToString();
                }
                else if (j == 15) //col 14 --time of taking N1
                {
                    timeTakingTest[0] = cols[j].ToString();
                }
                else if (j == 16) //col 14 --time of taking N2
                {
                    timeTakingTest[1] = cols[j].ToString();
                }
                else if (j == 17) //col 14 --time of taking N3
                {
                    timeTakingTest[2] = cols[j].ToString();
                }
                else if (j == 18) //col 14 --time of taking N4
                {
                    timeTakingTest[3] = cols[j].ToString();
                }
                else if (j == 19) //col 14 --time of taking N5
                {
                    timeTakingTest[4] = cols[j].ToString();
                }
                else if (j == 20) //col 14 --last result of taking N1
                {
                    resultTakingTest[0] = cols[j].ToString();
                }
                else if (j == 21) //col 14 --last result of taking N2
                {
                    resultTakingTest[1] = cols[j].ToString();
                }
                else if (j == 22) //col 14 --last result of taking N3
                {
                    resultTakingTest[2] = cols[j].ToString();
                }
                else if (j == 23) //col 14 --last result of taking N4
                {
                    resultTakingTest[3] = cols[j].ToString();
                }
                else if (j == 24) //col 14 --last result of taking N5
                {
                    resultTakingTest[4] = cols[j].ToString();
                }
                else if (j == 32) //col 33 -- street
                {
                    street = cols[j].ToString();
                }
                else if (j == 34) //col 33 -- city
                {
                    city = cols[j].ToString();
                }
                else if (j == 35) //col 33 -- prov
                {
                    prov = cols[j].ToString();
                }
                else if (j == 36) //col 33 -- country
                {
                    country = cols[j].ToString();
                }
                else if (j == 37) //col 33 -- postal
                {
                    postal = cols[j].ToString();
                }
                else if (j == 38) //col 33 -- phone
                {
                    phone = cols[j].ToString();
                }

            }
            result += "\"" + street + "\",\"" + city + "\",\"" + prov + "\",\"" + country + "\",\"" + postal + "\",\"" + email + "\",\"" + afiliation + "\",\"" + registrationType + "\",";

            for (int k = 0; k < timeTakingTest.Length; k++)
                result += "\"" + timeTakingTest[k] + "\",";

            for (int k = 0; k < resultTakingTest.Length; k++)
            {
                if (resultTakingTest[k] == "N/A")
                    result += "\"0\"";
                else if (resultTakingTest[k] == "Pass")
                    result += "\"1\"";
                else if (resultTakingTest[k] == "Fail")
                    result += "\"2\"";

                if(k < resultTakingTest.Length -1)
                    result += ",";

            }

            return result;
        }

        private string extractRegistrationData(DataRow row, string email)
        {
            string result = "";
            string street = "", city = "", prov = "", country = "", postal = "", phone = "", registrationType = "pay", fname="";
            string afiliation = "", learningPlace="", reason="", occupation="", occupationDet="";
            string media = "", withTeacher = "", withFam = "", withFriends = "", withSupervisor = "", withColl = "", withCustomer = "";
            string[] timeTakingTest = { "0", "0", "0", "0", "0" }; //new string[5]; //contain time of taking test l1-5 -- in order
            string[] resultTakingTest = { "0", "0", "0", "0", "0" }; //new string[5]; //contain lat result of taking test l1-5 -- in order

            var cols = row.ItemArray.ToList();
            for (int j = 0; j < cols.Count; j++)
            {
                if (j == 0) //col 1st = FName
                {
                    fname= cols[j].ToString();
                }
                else if (j == 2) //col 3rd LName
                {
                    string lname= cols[j].ToString();

                    string fullN = fname + " " + lname;

                    result += "\"" + fullN.ToUpper() +"\",";
                }
                /*else if (j == 3) //col 3rd Email
                {
                    tempEemail = cols[j].ToString(); //will position it later
                }*/
                else if (j == 4) //col Gender
               {
                   string gender = cols[j].ToString(); //will position it later
                    result += gender == "Male" ? "\"M\"," : "\"F\",";

               }
                else if (j == 5) //col 6rd DOB
                {
                    string[] dob = cols[j].ToString().Split("-");

                    result += "\"" + dob[0] + "\"," + dob[1] + "," + dob[2] + ",";
                }
                else if (j == 6) //col 6r -- password
                {
                    result += "\"" + cols[j].ToString() + "\",\"";
                }
                else if (j == 9) //col 6r -- Native lang
                {
                    result += getSingleSelectedCode(cols[j].ToString()) + "\",";
                }
                else if (j == 10) //col 6r -- reason
                {
                    reason = getSingleSelectedCode(cols[j].ToString()); // cols[j].ToString();
                }
                else if (j == 11) //col 6r -- occupation
                {
                    occupation = getSingleSelectedCode(cols[j].ToString());// cols[j].ToString();
                }
                else if (j == 12) //col 6r -- occupation Details
                {
                    occupationDet = getSingleSelectedCode(cols[j].ToString());// cols[j].ToString();
                }
                else if (j == 13) //col 6r -- learning place
                {
                    learningPlace = getSingleSelectedCode(cols[j].ToString());// cols[j].ToString();
                }
                else if (j == 14) //col 14 -- institution belong to
                {
                    afiliation = cols[j].ToString();
                }
                else if (j == 15) //col 14 --time of taking N1
                {
                    timeTakingTest[0] = cols[j].ToString();
                }
                else if (j == 16) //col 14 --time of taking N2
                {
                    timeTakingTest[1] = cols[j].ToString();
                }
                else if (j == 17) //col 14 --time of taking N3
                {
                    timeTakingTest[2] = cols[j].ToString();
                }
                else if (j == 18) //col 14 --time of taking N4
                {
                    timeTakingTest[3] = cols[j].ToString();
                }
                else if (j == 19) //col 14 --time of taking N5
                {
                    timeTakingTest[4] = cols[j].ToString();
                }
                else if (j == 20) //col 14 --last result of taking N1
                {
                    resultTakingTest[0] = cols[j].ToString();
                }
                else if (j == 21) //col 14 --last result of taking N2
                {
                    resultTakingTest[1] = cols[j].ToString();
                }
                else if (j == 22) //col 14 --last result of taking N3
                {
                    resultTakingTest[2] = cols[j].ToString();
                }
                else if (j == 23) //col 14 --last result of taking N4
                {
                    resultTakingTest[3] = cols[j].ToString();
                }
                else if (j == 24) //col 14 --last result of taking N5
                {
                    resultTakingTest[4] = cols[j].ToString();
                }
                else if (j == 25) //col 14 -- comm with Teacher
                {
                    withTeacher =   getMultiSelectedCodes(cols[j].ToString(), 5);
                }
                else if (j == 26) //col 14 -- comm with Teacher
                {
                    withFriends = getMultiSelectedCodes(cols[j].ToString(), 5);
                }
                else if (j == 27) //col 14 -- comm with Teacher
                {
                    withFam= getMultiSelectedCodes(cols[j].ToString(), 5);
                }
                else if (j == 28) //col 14 -- comm with Family
                {
                    withSupervisor = getMultiSelectedCodes(cols[j].ToString(), 5);
                }
                else if (j == 29) //col 14 -- comm with supervisor
                {
                    withColl = getMultiSelectedCodes(cols[j].ToString(), 5);
                }
                else if (j == 30) //col 14 -- comm with colleague
                {
                    withCustomer = getMultiSelectedCodes(cols[j].ToString(), 5);
                }
                else if (j == 31) //col 14 -- comm with Customer
                {
                    media = getMultiSelectedCodes(cols[j].ToString(), 9);
                }
               
                else if (j == 32) //col 33 -- street
                {
                    street = cols[j].ToString();
                }
                else if (j == 34) //col 33 -- city
                {
                    city = cols[j].ToString();
                }
                else if (j == 35) //col 33 -- prov
                {
                    prov = cols[j].ToString();
                }
                else if (j == 36) //col 33 -- country
                {
                    country = cols[j].ToString();
                }
                else if (j == 37) //col 33 -- postal
                {
                    postal = cols[j].ToString();
                }
                else if (j == 38) //col 33 -- phone
                {
                    phone = cols[j].ToString();
                }

            }
            result +="\"" + learningPlace + "\",\"" + reason + "\",\"" + occupation + "\",\"" + occupationDet + "\",\"" + media + "\",\"" + withTeacher + "\",\"" + withFriends + "\",\"";
            result += withFam + "\",\"" + withSupervisor + "\",\"" + withColl + "\",\"" + withCustomer + "\",";
            for (int k = 0; k < timeTakingTest.Length; k++)
                result += "\"" + timeTakingTest[k] + "\",";

            for (int k = 0; k < resultTakingTest.Length; k++)
            {
                if (resultTakingTest[k] == "N/A")
                    result += "\"0\"";
                else if (resultTakingTest[k] == "Pass")
                    result += "\"1\"";
                else if (resultTakingTest[k] == "Fail")
                    result += "\"2\"";

                if (k < resultTakingTest.Length - 1)
                    result += ",";

            }

            return result;
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
            return "\"Test Level\",\"Year Code\",\"Test Site\",\"Test Level\",\"Sequence No.\",\"Full Name\",\"Sex\",\"DOB Year\",\"DOB Month\",\"DOB Day\",\"Password\",\"Native Language\",\"Learning Place\",\"Reason\",\"Occupation\",\"Occupation Details\",\"Media\",	\"With Teachers\",	\"With Friends\",\"With Family\",\"With Supervisor\",\"With Colleagues\",\"With Customers\",\"Times Taking N1\",\"Times Taking N2\",\"Times Taking N3\",\"Times Taking N4\",\"Times Taking N5\",\"Latest N1 Result\",\"Latest N2 Result\",\"Latest N3 Result\",\"Latest N4 Result\",\"Latest N5 Result\"";
        }

        private string getMasterReportHeaders()
        {
            return "\"Sequence No.\",\"Level\",\" First Name\", \"Last Name\",\"DOB Year\",\"DOB Month\",\"DOB Day\",\"Street\",\"City\",\"Province\",\"Country\",\"Postal\",\"Email\",	\"Affiliation\",\" Registration Type\",\"Times Taking N1\",\"Times Taking N2\",\"Times Taking N3\",\"Times Taking N4\",\"Times Taking N5\",\"Latest N1 Result\",\"Latest N2 Result\",\"Latest N3 Result\",\"	Latest N4 Result\",\"Latest N5 Result\"";
        }
        private int getLangCode(string lang)
        {
            int code = 0;
            switch (lang)
            {
                case "Akan":
                    code = 602;
                        break;
                case "Amharic":
                    code = 602;
                    break;
                case "Arabic":
                    code = 701;
                    break;
                case "Ashanti":
                    code = 629;
                    break;
                case "Bambara":
                    code = 604;
                    break;
                case "Bemba":
                    code = 605;
                    break;
                case "Berber":
                    code = 606;
                    break;
                case "Chichewa":
                    code = 607;
                    break;
                case "Efik":
                    code = 608;
                    break;
                case "English":
                    code = 408;
                    break;
                case "Ewe":
                    code = 609;
                    break;
                case "French":
                    code = 411;
                    break;
                case "Fulani":
                    code = 610;
                    break;
                case "Ga":
                    code = 611;
                    break;
                case "Galla":
                    code = 612;
                    break;
                case "Hausa":
                    code = 613;
                    break;
                case "Ibo":
                    code = 614;
                    break;
                case "Kikongo":
                    code = 631;
                    break;
                case "Kikuyu":
                    code = 613;
                    break;
                case "Kinya Ruanda":
                    code = 632;
                    break;
                case "Kiswahili":
                    code = 616;
                    break;
                case "Lingala":
                    code = 617;
                    break;

                default:
                    code = 000;
                    break;

            }



             return code;
        }

        private string getSingleSelectedCode(string selected) // separator "-"
        {
            string code = "";
            if (selected.Contains("-")){
                string[] temp = selected.Split('-');
                code = temp[0].Trim();
            }

            return code;

        }
        private string getMultiSelectedCodes(string selected, int arrLength/*total options in this question*/) // separator "-"
        {
            
            string selectedCodes = "";
            string[] temp = selected.Split(',');
            string[] codes = new string[arrLength];
            //padd the ansswers
            for (int i = 0; i < arrLength; i++)
            {
                //initialize with empty
                codes[i] = " ";
            }
            for (int i=0; i< temp.Length; i++)
            {
                string strCode = getSingleSelectedCode(temp[i]);
                codes[Convert.ToInt16(strCode)-1] = strCode;
            }
            selectedCodes = string.Join("", codes);
            return selectedCodes;

        }
    }
}