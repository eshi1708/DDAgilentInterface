using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public class HarvestBuisness
    {
        //NOTE: ADDED 96 WELL TRANSPOSE ON REQUEST. Only difference is moving columnwise rather than rowwise. i.e. A1 A2 A3... instead of A1 B1 C1...

        private bool buildfilevalidate(bool isMasshunter, bool isGC, string selectedFile)
        {
            if (isMasshunter && isGC)
            {
                //we shouldn't ever end up here, but just to be sure.
                MessageBox.Show("Something has gone Very wrong. We can not have both a masshunter and gc run at the same time");
                return false;
            }
            if (!isGC && !isMasshunter)
            {
                MessageBox.Show("Please select an instrument before pressing 'Send run list to instrument'");
                return false;
            }

            if (selectedFile == null)
            {
                MessageBox.Show("Please select a csv batch file before pressing 'Send run list to instrument'");
                return false;
            }
            return true;
        }
        public void buildfileOrchard(bool isMasshunter, bool isGC, string selectedFile, string traytype, string fileformat, string outputLocation)
        {
            

            int LIMSBatchFileRow;
            string Method;
            int blankPosition;
            if(!buildfilevalidate(isMasshunter, isGC, selectedFile)) { return; }
            try
            {
                StreamReader inputread = new StreamReader(selectedFile);
                string line = inputread.ReadLine();
                if (line != null)
                {
                    //assume first line is column headers
                    if (!line.Contains("Order Choice"))
                    {
                        //not a valid csv input
                        MessageBox.Show("Cannot read 'Order Choice' column header in file.");
                        inputread.Close();
                        return;
                    }
                    string[] headers = line.Split(',');
                    int rowNumberCol = -1;
                    int patientNameCol = -1;
                    int orderChoiceCol = -1;
                    int sampleIDCol = -1;
                    for (int i = 0; i < 4; i++)
                    {
                        //note: file format seems to put quotes around every element.
                        if (headers[i] == "\"Order Choice\"" || headers[i] == "Order Choice")
                        {
                            orderChoiceCol = i;
                        }
                        if (headers[i] == "\"#\"" || headers[i] == "#")
                        {
                            rowNumberCol = i;
                        }
                        if (headers[i] == "\"Sample ID\"" || headers[i] == "Sample ID")
                        {
                            sampleIDCol = i;
                        }
                        if (headers[i] == "\"Patient\"" || headers[i] == "Patient")
                        {
                            patientNameCol = i;
                        }
                    }
                    line = inputread.ReadLine();
                    string[] delimiterarray = new string[] { "\"," };
                    string[] strArray = line.Split(delimiterarray, StringSplitOptions.None);
                    for (int i = 0; i < 4; i++)
                    {
                        strArray[i] = strArray[i].Replace("\"", "");
                    }
                    
                    string ORC = getOrderChoice(strArray[orderChoiceCol],traytype);
                    //check the order choices
                    
                    if (ORC == "VITA" || ORC == "VITD" && traytype== "")
                    {
                        inputread.Close();
                        MessageBox.Show("You must select a tray size first");
                        return;
                    }

                    switch ((ORC+"     ").Substring(0,4))//extra spaces to prevent outofrange exception
                    {
                        case "SCFA":
                            Method = "SCFA";
                            blankPosition = 150;
                            LIMSBatchFileRow = 10;
                            break;
                        case "VITA":
                        case "VITD":
                            if (traytype == "96 well" || traytype == "96 well transpose") { LIMSBatchFileRow = 6; }
                            else { LIMSBatchFileRow = 11; }
                            blankPosition = 1;
                            if (line.ToLower().Contains("epimer")) { Method = "DBS Vitamin D#epi.m"; }
                            else { Method = "DBS Vitamin D.m"; }
                            break;
                        default:
                            LIMSBatchFileRow = 10;
                            blankPosition = 1;
                            Method = ORC;
                            break;
                    }
                    int loopPosition150 = LIMSBatchFileRow;
                    string tempmethod = showMydialogBox();
                    if(tempmethod != "Cancelled")
                    {
                        Method = tempmethod;
                    }
                    string filelocation = outputLocation + "DDI_BATCH." + fileformat;
                    using (StreamWriter sw = new StreamWriter(filelocation))
                    {
                        int row = 0;
                        //add headers and first line/xmlblock
                        if (fileformat == "csv")
                        {
                            formatCSVHeaderOrchard(sw);
                            formatCSVDetailOrchard(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                        }
                        else
                        {
                            Format_XML_Header(sw);
                            Format_XML_Line_Detail(sw,ref row,blankPosition.ToString(),"BLK", "SAMPLE", blankPosition,Method);
                        }

                        //if 96 well tray, add another line/block
                        if (traytype == "96 well" || traytype == "96 well transpose")
                        {
                            if (fileformat == "csv"){ formatCSVDetailOrchard(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method); }
                            else{ Format_XML_Line_Detail(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method); }
                        }
                        switch( (ORC + "     ").Substring(0, 4))
                        {
                            case "SCFA":
                                if (fileformat == "csv")
                                {
                                    for(int i = 1; i <=4; i++)
                                    {
                                        formatCSVDetailOrchard(sw, ref row, i.ToString(), "STD "+i.ToString(), "CALIBRATION", blankPosition, Method);
                                    }
                                }
                                else
                                {
                                    for (int i = 1; i <= 4; i++)
                                    {
                                        Format_XML_Line_Detail(sw, ref row, i.ToString(), "STD " + i.ToString(), "CALIBRATION", blankPosition, Method);
                                    }
                                }

                                break;   
                            default:
                                if (fileformat == "csv")
                                {
                                    //note: to the best of my knowledge, the wells should be 96 and 54, not 94 and 56. Also, added in the 150 tray. it just uses 1-150 indexing.
                                    //ASK ABOUT DBS CAL VS CAL (since I added in the 150 well) also does capitalization matter for the agilent machines?
                                    if(traytype == "150 well")
                                    {
                                        for (int i = 1; i <= 5; i++)
                                        {
                                            
                                            formatCSVDetailOrchard(sw, ref row, i.ToString(), "Cal " + i.ToString(), "CALIBRATION", blankPosition, Method);
                                        }
                                    }
                                    else
                                    {
                                        for (int i = 1; i <= 5; i++)
                                        {
                                            formatCSVDetailOrchard(sw, ref row, convertToAgilentTrayTypeA(i.ToString(), traytype, "1"), "Cal " + i.ToString(), "CALIBRATION", blankPosition, Method);
                                        }
                                    }
                                    
                                }
                                else
                                {
                                    if (traytype == "150 well")
                                    {
                                        for (int i = 1; i <= 5; i++)
                                        {
                                            int j = i % 150;
                                            if (j == 0) { j = 150; }
                                            Format_XML_Line_Detail(sw, ref row, i.ToString(), "Cal " + i.ToString(), "CALIBRATION", blankPosition, Method);
                                            
                                        }
                                    }
                                    else
                                    {
                                        for (int i = 1; i <= 5; i++)
                                        {
                                            Format_XML_Line_Detail(sw, ref row, convertToAgilentTrayTypeA(i.ToString(), traytype, "1"), "Cal " + i.ToString(), "CALIBRATION", blankPosition, Method);
                                        }
                                    }
                                }
                                break;
                        }

                        if(traytype != "96 well" || traytype != "96 well transpose")
                        {
                            if(fileformat == "csv")
                            {
                                formatCSVDetailOrchard(sw,ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                formatCSVDetailOrchard(sw,ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                formatCSVDetailOrchard(sw,ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                            }
                            else
                            {
                                Format_XML_Line_Detail(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                Format_XML_Line_Detail(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                Format_XML_Line_Detail(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                            }
                        }

                        while(line != null)
                        {
                            //iterate trrough file and process it. This is where most things get handled.
                            string subSampleId = strArray[sampleIDCol];
                            if (subSampleId != "")
                            {
                                switch ((ORC + "     ").Substring(0, 4))
                                {
                                    case "SCFA":
                                        if (fileformat == "csv") { formatCSVDetailOrchard(sw, ref row, LIMSBatchFileRow.ToString(), subSampleId, "SAMPLE", blankPosition, Method); }
                                        else { Format_XML_Line_Detail(sw, ref row, LIMSBatchFileRow.ToString(), subSampleId, "SAMPLE", blankPosition, Method); }
                                        break;
                                    default:
                                        if (fileformat == "csv")
                                        {
                                            if (traytype == "150 well")
                                            {
                                                //loop back to wherever we started if we go over 150. by qing request.
                                                if(LIMSBatchFileRow == 150)
                                                {
                                                    LIMSBatchFileRow = loopPosition150;
                                                }
                                                formatCSVDetailOrchard(sw, ref row, LIMSBatchFileRow.ToString(), subSampleId, "SAMPLE", blankPosition, Method);
                                            }
                                            else
                                            {
                                                formatCSVDetailOrchard(sw, ref row, convertToAgilentTrayTypeA(LIMSBatchFileRow.ToString(), traytype, "1"), subSampleId, "SAMPLE", blankPosition, Method);
                                            }
                                        }
                                        else
                                        {
                                            if (traytype == "150 well")
                                            {
                                                if (LIMSBatchFileRow == 150)
                                                {
                                                    LIMSBatchFileRow = loopPosition150;
                                                }
                                                Format_XML_Line_Detail(sw, ref row, LIMSBatchFileRow.ToString(), subSampleId, "SAMPLE", blankPosition, Method);
                                            }
                                            else
                                            {
                                                Format_XML_Line_Detail(sw, ref row, convertToAgilentTrayTypeA(LIMSBatchFileRow.ToString(), traytype, "1"), subSampleId, "SAMPLE", blankPosition, Method);
                                            }
                                        }
                                        break;
                                }
                                LIMSBatchFileRow = LIMSBatchFileRow + 1;

                            }
                            //end process this line. get next line set up.

                            line = inputread.ReadLine();
                            if(line != null)
                            {
                                strArray = line.Split(delimiterarray,StringSplitOptions.None);
                            }
                            for (int i = 0; i < 4; i++)
                            {
                                strArray[i] = strArray[i].Replace("\"", "");
                            }
                        }

                        if (traytype == "96 well" || traytype == "96 well transpose")
                        {
                            if (fileformat == "csv")
                            {
                                formatCSVDetailOrchard(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                formatCSVDetailOrchard(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                formatCSVDetailOrchard(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                            }
                            else
                            {
                                Format_XML_Line_Detail(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                Format_XML_Line_Detail(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                                Format_XML_Line_Detail(sw, ref row, blankPosition.ToString(), "BLK", "SAMPLE", blankPosition, Method);
                            }
                        }
                        if (fileformat == "xml")
                        {
                            Format_XML_Footer(sw);
                        }
                    }
                }

                MessageBox.Show("Done");
            }
            catch(IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string getOrderChoice(string ORC, string traytype)
        {
            if (ORC.ToUpper().Contains("VITAMIN A"))
            {
                return "VITA";
            }
            if (ORC.ToUpper().Contains("VITAD"))
            {
                return "VITD";
            }
            if (ORC.ToUpper().Contains("VITAMIN D"))
            {
                return "VITD";
            }
            if (ORC.ToUpper().Contains("BLOODSPOT") || ORC.ToUpper().Contains("BLOOD SPOT"))
            {
                return "VITD";
            }
            return ORC;
        }

        private void formatCSVDetailOrchard(StreamWriter sw, ref int row, string location, string subSampleId, string sampleType, int blankPosition,string method)
        {
            
            
            row = row + 1;
            string locationString = "";
            if(subSampleId == "BLK")
            {
                if (method.ToLower() == "dbs vitamin d.m")
                {
                    locationString = locationString + "-";//stick in the minus sign
                }
                locationString = locationString + blankPosition.ToString();
            }
            else
            {
                locationString = location;
            }
            
            sw.WriteLine( row.ToString() + "," + locationString + "," + subSampleId + "," +  method +
                "," + sampleType);
            
        }
        private void formatCSVHeaderOrchard(StreamWriter sw)
        {
            sw.WriteLine("Number,Sample Position,Sample Name,Method,Sample Type");
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
            sw.WriteLine("<Number>" + row + "</Number>");
            if (subSampleID == "BLK")
            {
                if (Method.ToLower() == "dbs viamin d.m")
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
                sw.WriteLine("<Location>" + location + "</Location>");
            }
            sw.WriteLine("<Name>" + subSampleID + "</Name>");
            sw.WriteLine("<CDSMethod>" + Method + "</CDSMethod>");
            sw.WriteLine("<numberOfInj>" + 1 + "</numberOfInj>");
            sw.WriteLine("<sampleType>" + sampleType + "</sampleType>");
            if (sampleType == "CALIBRATION")
            {
                sw.WriteLine("<CalLevel>" + subSampleID.Replace("STD", "") + "</CalLevel>");
                sw.WriteLine("<calibration>REPLACE</calibration>");
                sw.WriteLine("<UpdateRT>REPLACE</UpdateRT> ");
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
        private string convertToAgilentTrayTypeA(string inputstr, string trayType, string traynumber)
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
                    int j = int.Parse(inputstr) % 150;
                    if (j == 0) { j = 150; }

                    return j.ToString();
            }
            int input = int.Parse(inputstr);
            if (input > 0 && input <= (rows * columns))
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
        private string showMydialogBox()
        {
            HarvestDialogBox testDialog = new HarvestDialogBox();

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
        
    }
}
