using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OracleClient;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Dapper;
using System.IO;
using System.Diagnostics;
using Dapper.Contrib.Extensions;

namespace WindowsFormsApp1
{
    public class SdiDataItem
    {
        public string EnteredText { get; set; }
        public string DisplayValue { get; set; }
        public string InstrumentId { get; set; }
        [ExplicitKey]
        public string KeyId1 { get; set; }
        [ExplicitKey]
        public string ParamId { get; set; }
        public bool ReleasedFlag { get; set; }
    }


    public class SenderToLIMS
    {
        public bool sendtoSaphhire(DataGridView dataview)
        {
            //set up regular expressions for later.
            Regex cellRegex1 = new Regex(@"^[A-W]\d\d\d\d\d\d-\d\d\d\d-\d-\d\d$");
            Regex cellRegex2 = new Regex(@"^[A-W]1");
            Regex cellRegex3 = new Regex(@"^[A-W]2");
            Regex cellRegex4 = new Regex(@"^[A-W]\d\d\d\d\d\d-\d\d\d\d-\d-\d\d-Sa$");
            string subsampleID;
            string paramID;
            string paramlistID = "";
            string Matrix = "";

            string cellValue;
            //make the oracle connection for now use the dev database. NOTE: configs set up to change to productiondb on release
            using (var Con = new OracleConnection(ConfigurationManager.ConnectionStrings["sapphireP"].ConnectionString))
            {
                Con.Open();
                //use transaction to revert changes in event of failure/error/weirdstuff. also revert for debug.
                using (var tx = Con.BeginTransaction())
                {
                    try {
                        //NOTE: packed cell stuff is old and not needed/ only used by a different lab. Deleted all lines depending on packed cells
                        //var t = new Stopwatch();
                        //t.Start();
                        for (int rw = 0; rw < dataview.RowCount; rw++)
                        {
                            if (dataview.Rows[rw].Cells[0].Value.ToString().ToLower().Contains("nt cal") || dataview.Rows[rw].Cells[0].Value.ToString().ToLower().Contains("blk")) { continue; }
                            subsampleID = "";
                            Matrix = "";
                            List<string> datalist = new List<string>();
                            List<string> paramIDlist = new List<string>();
                            for (int cl = 0; cl < dataview.ColumnCount; cl++)
                            {
                                paramID = dataview.Columns[cl].Name;
                                cellValue = dataview.Rows[rw].Cells[cl].Value.ToString();
                                //measuredunits is part of the packedcell stuff, and is not used. Or rather the case in which it would be used would never come through.
                                //most of the following cases are there to translate between formats
                                switch (paramID.ToLower())
                                {
                                    case "sample":
                                    case "sample name":
                                        bool n = cellRegex1.IsMatch(cellValue);
                                        if (cellRegex1.IsMatch(cellValue)) { subsampleID = cellValue + "-Sa"; }
                                        else if (cellRegex2.IsMatch(cellValue) || cellRegex3.IsMatch(cellValue) || cellRegex4.IsMatch(cellValue)) { subsampleID = cellValue; }
                                        //else must be blk or cal
                                        else { }
                                        //write the needed function later.
                                        getBatchInfoFromSubSampleID(subsampleID,ref paramlistID,ref Matrix,Con,tx);
                                        

                                        break;
                                    case "a-tocopherol":
                                        paramID = "Vitamin E alpha";
                                        break;
                                    case "g-tcopherol"://typo? copied from original code. Accurate. Keep as is.
                                        paramID = "Vitamin E gamma";
                                        break;
                                    case "5hiaa":
                                        paramID = "5-HIAA";
                                        break;
                                    case "methoxytyramine":
                                        paramID = "3-Methoxytyramine";
                                        break;
                                    case "gaba":
                                        paramID = "G-Aminobutyrate";
                                        break;
                                }
                                if (paramID != "" && paramID.ToLower() != "sample" && paramID.ToLower() != "sample name" && subsampleID != "")
                                {
                                    //removed units if statement due to not dealing with packed cells and therefore the if would never go through.                               

                                    

                                    //FIGURE OUT HOW TO MOVE THIS OUTSiDE THE FOR LOOPS
                                    //can maybe move out side of col loop. Current Idea, use this https://dapper-tutorial.net/execute#many-2 (note: curlybrace means anonymous type) maybe look up bulk insert?
                                    
                                    paramIDlist.Add(paramID);
                                    
                                    string multisqlstring = $"UPDATE sdidataitem SET enteredtext=:enteredtext, condition = null, displayvalue =:displayvalue,u_instrumentid='{Environment.MachineName}' WHERE keyid1=:keyid1 AND paramid=:paramid AND sdcid = :sdcid AND(releasedflag = 'N' OR releasedflag is null)";
                                    var x = new DynamicParameters();
                                    x.Add("enteredtext", cellValue, System.Data.DbType.AnsiString);
                                    x.Add("displayvalue", cellValue, System.Data.DbType.AnsiString);
                                    x.Add("keyid1", subsampleID, System.Data.DbType.AnsiString);
                                    x.Add("paramid", paramID, System.Data.DbType.AnsiString);
                                    x.Add("sdcid",Matrix+"Sample", System.Data.DbType.AnsiString);
                                    var a = Con.Execute(multisqlstring, x, transaction: tx);
                                }
                            }
                            //string multisqlstring = "UPDATE sdidataitem SET enteredtext = CAST(:enteredtext as varchar2 (255)), displayvalue = CAST(:displayvalue as varchar2 (255)) ,condition = null  ,u_instrumentid = '"
                            //    + Environment.MachineName + "' WHERE sdcid= '" + Matrix + "Sample' AND paramid= CAST(:paramid as varchar2 (20)) AND (releasedflag = 'N' OR releasedflag is null)";

                            

                           

                        }

                        //t.Stop();
                        //MessageBox.Show($"{t.ElapsedMilliseconds}ms");
                        //just always rollback for debug/dev sake. switch to commit when ready to deploy.DONT FORGET ME!
                        //tx.Rollback();
                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("sendtosapphire error: " + ex.Message);
                        //errors happened. revert any changes and display error message. Also, since we didn't finish, return false.
                        tx.Rollback();
                        return false;
                    }
                }
            }
            //successful if we get here.
            MessageBox.Show("Done");
            return true;
        }

        

        public bool sendtoHarvest(DataGridView dataview)
        {
            //determine where the file goes based on if it is a masshunter or not. if masshunter
            //we put the file path in the following string.?
            string filepath = ConfigurationManager.AppSettings.Get("HarvestOutPath");// maybe "c:\windows\temp"? or is this just a temp name?
            string tempfilepath = "";
            string subsampleID;
            string paramID = "";
            string cellValue;
            Random rnd = new Random();
            try
            {
                for (int rw = 0; rw < dataview.RowCount; rw++)
                {
                    subsampleID = "";
                    StreamWriter sw = null;
                    if (dataview.Rows[rw].Cells[0].Value.ToString().ToLower().Contains("nt cal")) { continue; }
                    for (int cl = 0; cl < dataview.ColumnCount; cl++)
                    {
                        //grab column name and cell value.
                        paramID = dataview.Columns[cl].Name.ToString().Replace("Final Conc.", "");
                        cellValue = dataview.Rows[rw].Cells[cl].Value.ToString();
                        if (paramID.ToLower() == "sample" || paramID.ToLower() == "sample name")
                        {
                            subsampleID = cellValue;
                            if ((subsampleID+"     ").Substring(0, 3) == "STD") { break; }
                            tempfilepath = Print_STH_Header(ref sw,subsampleID);
                        }
                        else if (paramID != "" && paramID != "#*")
                        {
                            //don't want to write the #* column. It's just there for reference anyways. not actual data.
                            sw.WriteLine("0.00\t0.00\t0.00\tBB\t0\t"+cellValue+"\t"+paramID);                            
                        }
                    }
                    if (sw == null)
                    {
                        continue;
                    }
                    if (sw != null)
                    {
                        sw.Close();
                    }

                    //attempt to rename the intended destination if it exists
                    if (File.Exists(filepath + subsampleID + ".txt"))
                    {
                        File.Move(filepath + subsampleID + ".txt", filepath + subsampleID + ((int)(rnd.NextDouble() * 100000000)).ToString() + ".txt");
                    }
                    File.Move(tempfilepath, filepath + subsampleID + ".txt");
                    
                    File.Delete(tempfilepath);
                }
                
            }
            catch(Exception ex)
            {
                MessageBox.Show("error in sending to Harvest: "+ ex.Message);
                return false;
            }
            MessageBox.Show("Done");
            //write to a temporary location, then move it to the drop file (prevent harvest from reading the file while you're writing a file into the drop file)
            return true; 
            
        }

        private string Print_STH_Header(ref StreamWriter sw,string subsampleID)
        {
            //build the streamwriter in here, since we need a new file for each row. streamwrite to a temporary folder.
            //also returns the temp file name, We'll need it later to move it into the original folder.
            string filename = Path.GetTempFileName();
            sw = new StreamWriter(filename);
            sw.WriteLine(@"File    \\" + Environment.MachineName + @"\varianws\data\");
            sw.WriteLine("--------------");
            sw.WriteLine("Injection Info");
            sw.WriteLine("--------------");
            sw.WriteLine("Sample:\t" + subsampleID);
            sw.WriteLine("--------------");
            sw.WriteLine("Peak Info for Channel Front");
            sw.WriteLine("--------------");
            sw.WriteLine("   Peaks (tR Timeoffset\tRRT\tSepcode\tWidth\tCounts\tResult\tName)");
            return filename;
        }

        private void getBatchInfoFromSubSampleID(string subsampleID, ref string paramlistID, ref string matrix,OracleConnection Con,OracleTransaction tx)
        {
            //just dumped the code for get sapphire matrix in here.
            switch ((subsampleID+"  ").Substring(0,1).ToUpper())
            {
                case "U":
                    matrix =  "Urine";
                    break;
                case "B":
                    matrix = "Blood";
                    break;
                case "F":
                    matrix = "Fecal";
                    break;
                case "H":
                    matrix = "Hair";
                    break;
                case "W":
                    matrix = "Water";
                    break;
                case "O":
                    matrix = "Other";
                    break;
            }

            var otherqueryout = Con.QueryFirstOrDefault<string>("select paramlistid from sdidata where sdcid='" + matrix + "Sample' and keyid1='" + subsampleID + "'", transaction: tx);
            if (otherqueryout != null)
            {
                paramlistID = otherqueryout;
            }
            else { paramlistID = ""; }
            return;
        }
    }
}
