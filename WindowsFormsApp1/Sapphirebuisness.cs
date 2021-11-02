using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OracleClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Dapper;
using System.IO;




namespace WindowsFormsApp1
{
    public class Sapphirebuisness
    {
        // runs the logic of building a csv file from sapphire.
        public void buildfileSaphCSV(string batchID, bool isGC,string fileDestination,string trayType, string trayNumber)
        {
            string trimBatchId = batchID.Trim();
            string regexPattern = @"^.[-]\d{8}[-]\d{2}$";//should match to any character, 8 digits, a dash/hyphen, 2 digits.
            Regex regex = new Regex(regexPattern);
            if (!regex.Match(trimBatchId).Success)
            {
                // if invalid ID, display message and stop trying to build.
                MessageBox.Show("Enter a valid Batch ID before pressing 'Send run list to instrument' button.");
                return;
            }

            string paramlistID = "";
            string batchDescription = "";
            getBatchInfoFromBatchID(trimBatchId, ref paramlistID, ref batchDescription);
            if (paramlistID == "ERROR")
            {
                //if we can't get the batch info due to something going wrong, the strings will be "ERROR. Don't continue"
                return;
            }
            if (paramlistID == "no items in batch")
            {
                MessageBox.Show("No items in batch.");
                return;
            }
            string method = "urine neurotransmitters.m";
            //they change the thing anyways.
            //may as well use the poput dialog in case they want it.
            method = ShowMyDialogBox(method);
            try
            {
                using (StreamWriter outputfile = new StreamWriter(fileDestination + trimBatchId + ".csv"))
                {
                    formatCSVHeader(outputfile);
                    //loop throught the 5 calibration lines
                    int csvrow = 1;

                    while (csvrow < 6)
                    {
                        //aren't the sequencecounter and row basically the same? They basically are.
                        formatCSVDetail(outputfile, ref csvrow, convertToAgilentTrayTypeA(csvrow.ToString(), trayType, trayNumber), "NT CAL " + csvrow.ToString(), "Calibration", method, csvrow.ToString(), "1");
                    }
                    //string querySQL = "SELECT * FROM workgroupitem where workgroupid='" + trimBatchId + "' order by to_number(workgroupitemid)";

                    using (var Con = new OracleConnection(ConfigurationManager.ConnectionStrings["sapphireP"].ConnectionString))
                    {
                        Con.Open();
                        try
                        {
                            var reader = Con.Query<workgroupitemClass>("SELECT * FROM workgroupitem where workgroupid='" +
                                trimBatchId + "' order by to_number(workgroupitemid)").ToList();
                            Regex saRegex = new Regex(@"^.*[-][S][a]$");
                            foreach (workgroupitemClass h in reader)
                            {
                                //need to deal with calibration again?
                                string subsampleid = h.keyid1;
                                string sampleType = "Sample";
                                string dilution = "1";
                                if (saRegex.Match(subsampleid).Success)
                                {
                                    dilution = getDilutionFactor(subsampleid, Con);

                                }
                                string truncatedsubsampleid = subsampleid.Replace("-Sa", "");
                                formatCSVDetail(outputfile, ref csvrow, convertToAgilentTrayTypeA(csvrow.ToString(), trayType, trayNumber), truncatedsubsampleid, sampleType, method, "", dilution);
                                if (isGC && sampleType == "Sample")
                                {
                                    //want to insert twice, put row in again with an extra S toindicate that its the second one.
                                    csvrow = csvrow - 1;//don't know if we need to preserve the row numbers through duplication.
                                    formatCSVDetail(outputfile, ref csvrow, convertToAgilentTrayTypeA(csvrow.ToString(), trayType, trayNumber), truncatedsubsampleid + "S", sampleType, method, "", dilution);

                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            //something went wrong. display a message and stop.
                            MessageBox.Show("Something has gone wrong. Terminating attempt to write to csv.\n" + ex.Message);
                            Con.Close();
                            return;
                        }
                        Con.Close();
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("terminating csv write. \n" + ex.Message);
                
            }
            MessageBox.Show("Done");
            return;
        }

        public string guessMethodXML(string batchID)
        {
            using(var Con = new OracleConnection(ConfigurationManager.ConnectionStrings["sapphireP"].ConnectionString))
            {
                var workgroupdesc = Con.QueryFirstOrDefault<string>("SELECT workgroupdesc FROM workgroup where workgroupid='" + batchID + "'");
                if( workgroupdesc != null)
                {
                    string checkdesc = workgroupdesc.ToLower();
                    if (checkdesc.Contains("dga"))
                    {
                        //double this one up.
                        return "DGA.m";
                    }
                    if (checkdesc.Contains("porphyrins"))
                    {
                        //double this one up.
                        return "Porphyrins.m";
                    }
                    if (checkdesc.Contains("coq10"))
                    {
                        return "ViteCoQ10.m";
                    }
                    if (checkdesc.Contains("permeabil") || checkdesc.Contains("lactulose"))
                    {
                        return "LactuloseMannitol.m";
                    }
                }
            }
            return "";
        }
        public void buildfileXML(string batchID, bool isGC, string fileDestination, string trayType, string trayNumber)
        {
            int blankPosition = 1;
            string trimBatchId = batchID.Trim();
            string regexPattern = @"^.[-]\d{8}[-]\d{2}$";//should match to any character, 8 digits, a dash/hyphen, 2 digits.
            Regex regex = new Regex(regexPattern);
            if (!regex.Match(trimBatchId).Success)
            {
                // if invalid ID, display message and stop trying to build.
                MessageBox.Show("Enter a valid Batch ID before pressing 'Send run list to instrument' button.");
                return;
            }
            int row = 0;
            string outputFile = @"C:\XML\IN\GC\DDI_BATCH.XML";//should this be named like this?

            try
            {
                using (StreamWriter writer = new StreamWriter(fileDestination+"DDI_BATCH.XML"))
                {
                    //write header
                    Format_XML_Header(writer);
                    //guess method
                    string method = ShowMyDialogBox(guessMethodXML(trimBatchId));
                    using (var Con = new OracleConnection(ConfigurationManager.ConnectionStrings["sapphireP"].ConnectionString))
                    {
                        try
                        {
                            if (method == "DGA.m" )
                            {
                                //add 4 cal lines
                                //add samples twice if using one of these two methods ALSO ONLY IF WE HAVE AN ACTUAL SAMPLE, NOT A CONTROL OR SOMETHING.
                                for (int i = 0; i < 4; i++)
                                {
                                    Format_XML_Line_Detail(writer, ref row, (row+1).ToString(), "Cal " + (row + 1), "CALIBRATION", blankPosition, method);
                                }
                            }
                            else if ( method == "Porphyrins.m")
                            {
                                //add 5 cal lines
                                for (int i = 0; i < 5; i++)
                                {
                                    Format_XML_Line_Detail(writer, ref row, (row + 1).ToString(), "Cal " + (row + 1), "CALIBRATION", blankPosition, method);
                                }
                            }
                            var something = Con.Query<workgroupitemClass>("SELECT * FROM workgroupitem where workgroupid = '" + batchID + "' order by to_number(workgroupitemid)").ToList<workgroupitemClass>();
                            int cntr = 0;
                            foreach (workgroupitemClass h in something)
                            {

                                if(method == "DGA.m"  && cntr == something.Count() - 2){
                                    row = 4;
                                }
                                if (method == "Porphyrins.m" && cntr == something.Count() - 2)
                                {
                                    row = 5;
                                }
                                string subSampleID = h.keyid1.Replace("-Sa", "");
                                //ASK IF THIS WILL HAVE CONTROLS AND SUCH IN FRONT/BACK OR NOT.
                                Format_XML_Line_Detail(writer, ref row, (row + 1).ToString(), subSampleID, "SAMPLE", blankPosition, method);
                                if (method == "DGA.m" && row > 6 && cntr<something.Count()-2)
                                {
                                    //add samples twice if using one of these two methods ALSO ONLY IF WE HAVE AN ACTUAL SAMPLE, NOT A CONTROL OR SOMETHING.
                                    Format_XML_Line_Detail(writer, ref row, (row + 1).ToString(), subSampleID + "S", "SAMPLE", blankPosition, method);
                                }
                                else if (method == "Porphyrins.m" && row > 7 && cntr<something.Count()-2)
                                {
                                    Format_XML_Line_Detail(writer, ref row, (row + 1).ToString(), subSampleID + "S", "SAMPLE", blankPosition, method);
                                }
                                cntr++;
                            }
                            Format_XML_Footer(writer);
                            

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Something has gone wrong. Terminating XML write. \n" + ex.Message);
                            return;
                        }
                    }
                    
                }
            }
            catch (DirectoryNotFoundException direx)
            {
                MessageBox.Show("Terminating xml write.\n directory not found: " + direx.Message);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Terminating xml write.\n" + ex.Message);
                return;
            }
            MessageBox.Show("Done");
            return;

        }

        private void Format_XML_Header(StreamWriter sw)
        {
            sw.WriteLine("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>");
            sw.WriteLine("<Samples xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:noNamespaceSchemaLocation=\"c:\\Chem32\\core\\worklist.xsd\">");
        }
        private void Format_XML_Footer(StreamWriter sw)
        {
            sw.WriteLine("</Samples>");
        }
        private void Format_XML_Line_Detail(StreamWriter sw, ref int row, string location, string subSampleID, string sampleType, int blankPosition, string Method)
        {
            row = row + 1;
            sw.WriteLine("<Sample>");
            sw.WriteLine("<Number>"+row+"</Number>");
            if (subSampleID == "BLK")
            {
                if(Method.ToLower()=="dbs viamin d.m")
                {
                    sw.WriteLine("<Location>Vial " + blankPosition + "</Location>");
                }
                else
                {
                    sw.WriteLine("<Location>" + blankPosition + "</Location>");
                }
            }
            else
            {
                sw.WriteLine("<Location>"+location+"</Location>");
            }
            sw.WriteLine("<Name>" + subSampleID + "</Name>");
            sw.WriteLine("<CDSMethod>" + Method + "</CDSMethod>");
            sw.WriteLine("<numberOfInj>" + 1 + "</numberOfInj>");
            sw.WriteLine("<sampleType>" + sampleType + "</sampleType>");
            if (sampleType == "CALIBRATION")
            {
                sw.WriteLine("<CalLevel>" + subSampleID.Replace("STD","") + "</CalLevel>");
                sw.WriteLine("<calibration>REPLACE</calibration>");
                sw.WriteLine("<UpdateRT> REPLACE</UpdateRT> ");
            }
            else//must be a sample then.
            {
                if (subSampleID != "BLK")
                {
                    sw.WriteLine("<description>LIMS sample</description>");
                    sw.WriteLine("<LimsID>" + subSampleID + "</LimsID>");
                }
            }
            sw.WriteLine("</Sample>");
        }



        private string convertToAgilentTrayTypeA(string inputstr,string trayType, string traynumber)
        {
            string output = "error";
            int rows = 0;
            int columns = 0;
            int row;
            int column;
            bool transpose = false;
            switch (trayType)
            {
                case "54 well":
                    rows = 6;
                    columns = 9;
                    break;
                case "96 well":
                    rows = 8;
                    columns = 12;
                    break;
                case "96 well transpose":
                    rows = 8;
                    columns = 12;
                    transpose = true;
                    break;
                case "150 well":
                    int j = int.Parse(inputstr)%150;
                    if(j == 0) { j = 150; }

                    return j.ToString();
            }
            int input = int.Parse(inputstr);
            if(input > 0 && input <= (rows * columns))
            {
                if (transpose)
                {
                    string tray = "p" + traynumber + "-";
                    column = (input - 1) % columns + 1;
                    row = ((input - 1) / columns);
                    output = tray + (char)(97 + row) + column;
                }
                else
                {
                    string tray = "p" + traynumber + "-";
                    row = (input - 1) % rows;
                    column = ((input - 1) / rows) + 1;
                    output = tray + (char)(97 + row) + column;
                }
            }
            return output;
        }


        private string getDilutionFactor(string keyid1,OracleConnection Con)
        {
            string paramlistID = "";
            string Output = "";
            string saphMat = getSapphireMatrix(keyid1);
            string sqlStr = "SELECT displayvalue,paramlistid FROM SDIDataItem where sdcid='" + saphMat + "Sample' and KeyID1 like '" + (keyid1 + "               ").Substring(0, 14) + "%'" +
                " and paramid='DilutionFactor' and paramtype='Adjusted'";
            var reader = Con.QueryFirstOrDefault<sdidataitem_displayvalue_paramlistid>(sqlStr);
            if (reader != null)
            {
                paramlistID = reader.paramlistid;
                Output = reader.displayvalue;
            }

            if(Output != null && Output != "")//the <2 is basically a null/emptystring check.
            {
                //now check for other stuff
                
                if (paramlistID.ToLower().Contains("small") && paramlistID.ToLower().Contains("urine animo")){
                    Output = "15.75";
                }
                else
                {
                    var queryOut = Con.QueryFirstOrDefault<string>("SELECT displayvalue FROM SDIDataItem where sdcid='" + saphMat + "Sample' and " +
                        "KeyID1 like '" + (keyid1 + "               ").Substring(0, 14) + "%' and paramid='Creatinine' and paramtype='Adjusted'");
                    string creatineVal = "0";
                    if (queryOut != null)
                    {
                        creatineVal = queryOut;
                    }
                    //what do if the creatine value doesn't exist? apparently set output to null and return.
                    if(creatineVal == null)
                    {
                        //assuning that the 2.whatever version is accurate.
                        return "";
                    }
                    //ask bill if we want rounding or decimal stripping.IMPORTANT.!!!!!!!!!!!. ONE OF THE DIFFERENCES BETWEEN ORIGINAL AND CURRENT/
                    if(paramlistID.ToLower().Contains("urine animo") || paramlistID.ToLower().Contains("urine adama"))
                    {
                        int d = (int) Math.Round(double.Parse(creatineVal));
                        if(d < 1)
                        {
                            //either too small or there is no value in the database, something has gone wrong.
                            Output = "";
                        }
                        else if (d <= 20)
                        {
                            Output = "1.1";
                        }
                        else if(d <= 40)
                        {
                            Output = "2";
                        }
                        else if(d <= 80)
                        {
                            Output = "4";
                        }
                        else if(d <= 120)
                        {
                            Output = "5";
                        }
                        else if(d <= 150)
                        {
                            Output = "8";
                        }
                        else
                        {
                            Output = "10";
                        }
                    }
                    else
                    {
                        int d = (int)Math.Round(double.Parse(creatineVal));
                        if (d < 1)
                        {
                            //either too small or there is no value in the database, something has gone wrong.
                            Output = "";
                        }
                        else if (d <= 20)
                        {
                            Output = "2";
                        }
                        else if (d <= 40)
                        {
                            Output = "4";
                        }
                        else if (d <= 80)
                        {
                            Output = "6";
                        }
                        else if (d <= 120)
                        {
                            Output = "8";
                        }
                        else if (d <= 150)
                        {
                            Output = "10";
                        }
                        else
                        {
                            Output = "12";
                        }

                    }

                }
            }
            
            return Output;
        }

        public string ShowMyDialogBox(string defaulttxt)
        {
            SapphireDialogBox testDialog = new SapphireDialogBox(defaulttxt);

            string output;
            // Show testDialog as a modal dialog and determine if DialogResult = OK.
            if (testDialog.ShowDialog() == DialogResult.OK)
            {
                // Read the contents of testDialog's TextBox.
                output = testDialog.methodstr;
            }
            else
            {
                output = "Cancelled";
            }
            testDialog.Dispose();
            return output;
        }
        private void getBatchInfoFromBatchID(string batchid, ref string paramlistID, ref string batchDescription )
        {
            string sqlQuery = "select keyid1 from workgroupitem where workgroupid = '" + batchid + "' order by usersequence ASC";
            using (var Con = new OracleConnection(ConfigurationManager.ConnectionStrings["sapphireP"].ConnectionString))
            {
                Con.Open();
                try
                {
                    var rdr = Con.QueryFirstOrDefault<string>(sqlQuery);
                    if (rdr == null)
                    {
                        //if there are no items read from the query, then the batch does not exist or is empty. close transaction and exit.
                        paramlistID = "no items in batch";
                        Con.Close();
                        return;
                    }
                    string subSampleID = rdr;
                    string anotherQuery = "select workgroupdesc from workgroup where workgroupid='" + batchid + "'";
                    var redr = Con.QueryFirstOrDefault<string>(anotherQuery);
                    if (redr != null)
                    {
                        batchDescription = redr;
                    }
                    string yetanotherQuery = "select paramlistid from sdidata where sdcid='" + getSapphireMatrix(batchid) + "Sample' and keyid1='" + subSampleID + "'";
                    var reader = Con.QueryFirstOrDefault<string>(yetanotherQuery);//do we need the paramlistversionid? don't think it gets used?
                    if (reader != null)
                    {
                        paramlistID = reader;
                    }
                    Con.Close();
                }
                catch(Exception exe)
                {
                    //something went wrong. display a message and stop.
                    MessageBox.Show("Something has gone wrong. Terminating attempt to retreive batch Info.\n" + exe.Message);
                    paramlistID = "ERROR";
                    batchDescription = "ERROR";
                    Con.Close();
                }
            }
            return;
        }

        private void formatCSVDetail(StreamWriter sw, ref int row, string location, string subSampleId, string sampleType, string method, string levelName, string dilution)
        {

            sw.WriteLine("" + row + "," + subSampleId + "," + location + "," + "D:\\MassHunter\\methods\\running method\\" + method +
                "," + sampleType + "," + levelName + "," + dilution);
            row = row + 1;
        }
        private void formatCSVHeader(StreamWriter sw)
        {
            sw.WriteLine("Number,Sample Name,Sample Position,Method,Sample Type,Level Name,Dilution");
        }
        private string getSapphireMatrix(string letter)
        {
            switch(letter.ToUpper()[0])
            {
                case 'U':
                    return "Urine";
                case 'B':
                    return "Blood";
                case 'F':
                    return "Fecal";
                case 'H':
                    return "Hair";
                case 'W':
                    return "Water";
                case 'O':
                    return "Other";
            }
            return "";
        }
    }
}
