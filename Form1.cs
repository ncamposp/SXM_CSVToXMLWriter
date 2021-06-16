using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Globalization;

namespace ExcelToXML
{
    public partial class Form1 : Form
    {
        /* Records
         * ***********************************************************************
         * The names of these strings correspond to the XML format ie. the SXM database uses the name Product Family instead of Make
         * ***********************************************************************
         * VIN - Vehicle Identification Number               -- SXM DB Column Name: VIN
         * FacReDate - Factory Release Date                  -- SXM DB Column Name: Factory Release Date OR Built Date
         * make - The vehicle's make                         -- SXM DB Column Nane: Product Family
         * model - The model number of the vehicle           -- SXM DB Column Name: Model No
         * modelYear - Model year of the vehicle             -- SXM DB Column Name: Model Year
         * dealerID - The 4-digit order dealer code          -- SXM DB Column Name: Order Dealer Code
         */
        public struct Records
        {
            public string VIN;
            public string FacReDate;
            public string make;
            public string model;
            public string modelYear;
            public string dealerID;

        }
        /* RecordsSold
         * **********************************************************************
         * VIN - Vehicle Identification Number              -- SXM DB Column Name: VIN
         * saleDate - The sold date of the vehicle          -- SXM DB Column Name: Retail Sold Rpt Date OR In-Service Date OR Warranty Registration Date OR Deilvered Date
         * make - The vehicle's make                        -- SXM DB Column Nane: Product Family
         * model - The model number of the vehicle          -- SXM DB Column Name: Model No
         * modelYear - Model year of the vehicle            -- SXM DB Column Name: Model Year
         * firstName - The country code ie. US or CA        -- SXM DB Column Name: Domicile Country
         * lastName - The Name of the buyer                 -- SXM DB Column Name: Customer Name
         * addressLn1 - The address of the buyer            -- SXM DB Column Name: Customer Street
         * city - The buyer's city                          -- SXM DB Column Name: Customer City
         * state - The buyer's state                        -- SXM DB Column Name: Customer State
         * zip - The buyer's zip code                       -- SXM DB Column Name: Customer Zip
         * country - The buyer's country code ie. USA/CAN   -- SXM DB Column Name: Domicile Country
         */
        public struct RecordsSold
        {
            public string VIN;
            public string saleDate;
            public string make;
            public string model;
            public string modelYear;
            public string firstName; //Currently unused -- Hard coded NA
            public string lastName; //Name of Buyer
            public string addressLn1;
            public string city;
            public string state;
            public string zip;
            public string country; //Currently unused -- Hard coded USA
            public string dealerID;
        }
        public Form1()
        {
            InitializeComponent();
        }

        /* button1_Click()
         * ***********************************************************************
         * 1. Opens a dialog for a formatted .csv file
         * 2. Makes an array of built records structs
         * 3. Reads from the csv file and inputs it into the structs
         * 4. Use the information in the structs to create an xml and fill it with data
         */
        private void button1_Click(object sender, EventArgs e)
        {
            //----------------------------------------------- Part 1 ----------------------------------------------- 
            OpenFileDialog fill = new OpenFileDialog();
            fill.ShowDialog();
            string csvPath = fill.FileName.ToString();


            //We're going to create two XMLs -- one for Built and one for Sold
            string dirPath = @"C:\Users\" + Environment.UserName + @"\Desktop\XMLs";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            string builtPath = dirPath + @"\RRDAIMLER_BUILT_" + DateTime.Now.ToString("yyyyMMdd")+ "_";
            //string soldPath = dirPath + @"\RRDAIMLER_SOLD_" + @".xml";

            //Get the number of records
            var lines = File.ReadAllLines(csvPath);
            var count = lines.Length - 1;

            //----------------------------------------------- Part 2 ----------------------------------------------- 
            //make an array of structs that will hold all the parsed csv info
            Records[] buildRecords = new Records[count];

            List<string> dates = new List<string>(); //dates will hold all dates but then be sorted to only contain unique dates.
            //Once we have unique dates, we know how many xml files we need to make. For each unique date, create an xml and run a loop.
            //Write it to an xml file if it matches the year. Skip if not. 

            //----------------------------------------------- Part 3 ----------------------------------------------- 
            //Parse through the csv and put things in the corresponding Records[]
            using (var reader = new StreamReader(csvPath))
            {
                reader.ReadLine(); //skip first
                int counter = 0; //index
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(','); //read line then split it
                    string toDisplay = string.Join(Environment.NewLine, values);
                    //MessageBox.Show(toDisplay);
                    //append split value into struct
                    buildRecords[counter].VIN = values[0];
                    if (values[1] == "")//this date will either be Factory Date or Built Date
                    {
                        buildRecords[counter].FacReDate = values[2];
                    }
                    else
                    {
                        buildRecords[counter].FacReDate = values[1];
                    }

                    buildRecords[counter].make = values[3];
                    buildRecords[counter].model = values[4];
                    buildRecords[counter].modelYear = values[5];
                    buildRecords[counter].dealerID = values[6];
                    //Add every year indiscriminately
                    dates.Add(buildRecords[counter].FacReDate);

                    counter++;
                }
            }
            //Remove all duplicates
            List<string> uniqueDates = dates.Distinct().ToList();
            //var message = string.Join(",",uniqueDates);

            //Now run a for loop for every unique year. 
            var numYears = uniqueDates.Count();

            //uniqueDates contains a "blank" years field. It is the last year. Let's not include it by using numYears-1
            //Another solutions would be to go through uniqueDates and remove "" (do this later)
            //uniqueDates.Remove("");
            //----------------------------------------------- Part 4 ----------------------------------------------- 
            for (int n = 0; n < numYears-1; n++)
            {
                var curYear = uniqueDates[n]; //curYear dictates whether we add a year during this loop
                                              //Next part, open the xmls and write the basic stuff to them
                XNamespace ns = "http://www.siriusxm.com/Schemas/SC/DataTransmission/RRS";
                XNamespace xsiXs = "http://www.w3.org/2001/XMLSchema-instance";
                XDocument objXDoc = new XDocument(
                    new XDeclaration("1.0", "UTF-8", "no"),
                    new XElement(ns + "RETAIL_SALE_XML",
                        new XAttribute(XNamespace.Xmlns + "xsi", xsiXs),
                        new XAttribute(xsiXs + "schemaLocation", "http://www.siriusxm.com/Schemas/SC/DataTransmission/RRS RetailTruckOEM_1_0.xsd")));


                objXDoc.Root.Add(new XElement(ns + "HEADER",
                    new XElement(ns + "SENDER_ID", "RRDAIMLER"),
                    new XElement(ns + "RETAILER_TRANSACTION_ID", 45531411 + n),
                    new XElement(ns + "FILE_SENT_DATE", "2021-06-14"))); //fix hard-coded date

                var xmlIndex = 1;
                for (int i = 0; i < count; i++)
                {
                    //for every record, check if the date matches our curYear

                    if (buildRecords[i].FacReDate == curYear)
                    {
                        //cast this date string to the right format YYYY-MM-DD instead of MM/DD/YYYY
                        string newDate = "";
                        if (buildRecords[i].FacReDate != "")
                        {
                            DateTime date = Convert.ToDateTime(buildRecords[i].FacReDate);

                            newDate = date.ToString("yyyy-MM-dd");
                        }
                        //MessageBox.Show(newDate);
                        objXDoc.Root.Add(new XElement(ns + "RETAIL_BUILT_RECORD",
                                        new XAttribute("TRANSACTION_ID", xmlIndex),
                                        new XElement(ns + "EVENT_TYPE_ID", "RETAIL_RADIO_BUILT"),
                                        new XElement(ns + "EVENT_DATE", "2021-06-14"),//hard coded
                                        new XElement(ns + "RADIO_ID", ""), //No radio IDs
                                        new XElement(ns + "VIN", buildRecords[i].VIN),
                                        new XElement(ns + "BUILT_DATE", newDate),
                                        new XElement(ns + "PROGRAM_CODE", "DATRK3MOAA"),
                                        new XElement(ns + "VEHICLE_MAKE", buildRecords[i].make),
                                        new XElement(ns + "VEHICLE_MODEL", buildRecords[i].model),
                                        new XElement(ns + "VEHICLE_MODEL_YEAR", buildRecords[i].modelYear),
                                        new XElement(ns + "DEALER_ID", buildRecords[i].dealerID)
                                        )
                            );
                        xmlIndex++;
                    }
                    else
                    {
                        //on to the next record
                    }

                }
                //Change the curYear to a normal format YYYY-MM-DD instead of MM/DD/YYYY
                string suffixDate = "";
                DateTime oldDate = Convert.ToDateTime(curYear);
                suffixDate = oldDate.ToString("yyyy-MM-dd");
                //make the path of the file corresponding to the date(may need to cast it back to appropriate windows format)
                objXDoc.Save(builtPath + suffixDate + @".xml");

            }

            MessageBox.Show("Complete");
        }
        /* button2_Click()
         * ***********************************************************************
         * 1. Opens a dialog for a formatted .csv file
         * 2. Makes an array of built records structs
         * 3. Reads from the csv file and inputs it into the structs
         * 4. Use the information in the structs to create an xml and fill it with data
         */
        private void button2_Click(object sender, EventArgs e)
        {
            //----------------------------------------------- Part 1 ----------------------------------------------- 
            //Specify the csv file
            OpenFileDialog fill = new OpenFileDialog();
            fill.ShowDialog();
            string csvPath = fill.FileName.ToString();

            //Create the SOLD XML path
            string dirPath = @"C:\Users\" + Environment.UserName + @"\Desktop\XMLs";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            //string builtPath = dirPath + @"\RRDAIMLER_BUILT_" + @".xml";
            string soldPath = dirPath + @"\RRDAIMLER_SOLD_" + DateTime.Now.ToString("yyyyMMdd")+ "_";

            //Get the number of records
            var lines = File.ReadAllLines(csvPath);
            var count = lines.Length - 1;
            //MessageBox.Show("Lines: "+count);

            //----------------------------------------------- Part 2 ----------------------------------------------- 
            //make an array of structs that will hold all the parsed csv info
            RecordsSold[] soldRecords = new RecordsSold[count];

            List<string> dates = new List<string>(); //dates will hold all dates but then be sorted to only contain unique dates.
            //Once we have unique dates, we know how many xml files we need to make. For each unique date, create an xml and run a loop.
            //Write it to an xml file if it matches the year. Skip if not. 

            //----------------------------------------------- Part 3 ----------------------------------------------- 
            //Parse through the csv and put things in the corresponding Records[]
            using (var reader = new StreamReader(csvPath))
            {
                reader.ReadLine(); //skip first
                int counter = 0; //index
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(','); //read line then split it
                    string toDisplay = string.Join(Environment.NewLine, values);
                    //MessageBox.Show(toDisplay);
                    //append split value into struct
                    soldRecords[counter].VIN = values[0];
                    //Logic for if a date is missing (swap over to switch case for readability?)
                    if (values[1] == "")
                    {
                        if (values[2] == "")
                        {
                            if (values[3] == "")
                            {
                                soldRecords[counter].saleDate = values[4];
                            }
                            else
                            {
                                soldRecords[counter].saleDate = values[3];
                            }
                        }
                        else
                        {
                            soldRecords[counter].saleDate = values[2];
                        }
                    }
                    else
                    {
                        soldRecords[counter].saleDate = values[1]; //saleDate will be dependent on what is available
                    }
                    soldRecords[counter].make = values[5];
                    soldRecords[counter].model = values[6];
                    soldRecords[counter].modelYear = values[7];
                    soldRecords[counter].firstName = values[8];//NA or CA
                    soldRecords[counter].lastName = values[9];
                    soldRecords[counter].addressLn1 = values[10];
                    soldRecords[counter].city = values[11];
                    soldRecords[counter].state = values[12];
                    soldRecords[counter].zip = values[13];
                    soldRecords[counter].dealerID = values[14];
                    soldRecords[counter].country = values[8];//USA or CAN
                    dates.Add(soldRecords[counter].saleDate);
                    counter++;
                }
            }
            //Remove all duplicates
            //MessageBox.Show("numyears: " + dates.Count);
            List<string> uniqueDates = dates.Distinct().ToList();
            //var message = string.Join(",",uniqueDates);
            //MessageBox.Show(message);
            //Now run a for loop for every unique year. 
            var numYears = uniqueDates.Count();
            //MessageBox.Show("numyears: "+ numYears);
            //uniqueDates.Remove("");
            //----------------------------------------------- Part 4 ----------------------------------------------- 
            for (int n = 0; n < numYears-1; n++)
            {
                var curYear = uniqueDates[n]; //curYear dictates whether we add a year during this loop
                                              //Next part, open the xmls and write the basic stuff to them
                XNamespace ns = "http://www.siriusxm.com/Schemas/SC/DataTransmission/RRS";
                XNamespace xsiXs = "http://www.w3.org/2001/XMLSchema-instance";
                XDocument objXDoc = new XDocument(
                    new XDeclaration("1.0", "UTF-8", "no"),
                    new XElement(ns + "RETAIL_SALE_XML",
                        new XAttribute(XNamespace.Xmlns + "xsi", xsiXs),
                        new XAttribute(xsiXs + "schemaLocation", "http://www.siriusxm.com/Schemas/SC/DataTransmission/RRS RetailTruckOEM_1_0.xsd")));


                objXDoc.Root.Add(new XElement(ns + "HEADER",
                    new XElement(ns + "SENDER_ID", "RRDAIMLER"),
                    new XElement(ns + "RETAILER_TRANSACTION_ID", 45531911 + n),
                    new XElement(ns + "FILE_SENT_DATE", "2021-06-14"))); //fix hard-coded date

                var xmlIndex = 1;
                //Next part, open the xmls and write the basic stuff to them
                //Open the excel file to grab data from
                for (int i = 0; i < count; i++)
                {
                    if (soldRecords[i].saleDate == curYear)
                    {
                        //cast this date string to the right format YYYY-MM-DD instead of MM/DD/YYYY
                        string newDate = "";
                        if (soldRecords[i].saleDate != "")
                        {
                            DateTime date = Convert.ToDateTime(soldRecords[i].saleDate);

                            newDate = date.ToString("yyyy-MM-dd");
                        }
                        //MessageBox.Show(newDate);
                        objXDoc.Root.Add(new XElement(ns + "RETAIL_BUILT_RECORD",
                                        new XAttribute("TRANSACTION_ID", xmlIndex),
                                        new XElement(ns + "EVENT_TYPE_ID", "RETAIL_RADIO_BUILT"),
                                        new XElement(ns + "EVENT_DATE", "2021-06-14"), //hard-coded
                                        new XElement(ns + "RADIO_ID", ""), //No radio IDs
                                        new XElement(ns + "VIN", soldRecords[i].VIN),
                                        new XElement(ns + "BUILT_DATE", newDate),
                                        new XElement(ns + "PROGRAM_CODE", "DATRK3MOAA"),
                                        new XElement(ns + "VEHICLE_MAKE", soldRecords[i].make),
                                        new XElement(ns + "VEHICLE_MODEL", soldRecords[i].model),
                                        new XElement(ns + "VEHICLE_MODEL_YEAR", soldRecords[i].modelYear),
                                        new XElement(ns + "FIRST_NAME", "NA"), //soldRecords[i].firstName
                                        new XElement(ns + "LAST_NAME", soldRecords[i].lastName),
                                        new XElement(ns + "ADDRESS_LINE1", soldRecords[i].addressLn1),
                                        new XElement(ns + "ADDRESS_LINE2", ""),
                                        new XElement(ns + "CITY", soldRecords[i].city),
                                        new XElement(ns + "STATE", soldRecords[i].state),
                                        new XElement(ns + "ZIP", soldRecords[i].zip),
                                        new XElement(ns + "COUNTRY", "USA"), //change this to the struct value if NOT only USA
                                        new XElement(ns + "PHONE", ""), //Not required
                                        new XElement(ns + "EMAIL", ""),
                                        new XElement(ns + "DEALER_ID", soldRecords[i].dealerID)
                                        )
                            );
                        xmlIndex++;
                    }

                    else
                    {
                        //skip record
                    }
                }
                //Change the curYear to a normal format YYYY-MM-DD instead of MM/DD/YYYY
                string suffixDate = "";
                DateTime oldDate = Convert.ToDateTime(curYear);
                suffixDate = oldDate.ToString("yyyy-MM-dd");
                //make the path of the file corresponding to the date(may need to cast it back to appropriate windows format)
                objXDoc.Save(soldPath + suffixDate + @".xml");

            }
            MessageBox.Show("Complete");
        }
    }
}
