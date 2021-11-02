using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using ExcelDataReader;
using System.Collections;
using System.Configuration;
using System.Data.OracleClient;
using System.Text.RegularExpressions;
using Dapper;
using Application = System.Windows.Forms.Application;
using System.Reflection;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private string batchid;
        private int trayNumber;
        private bool ismasshunter;
        
        private bool isgaschromo;
        private string worklistFolder;
        private string trayTypeSapphire;
        private string trayNumberSapphire;
        private string batchfileHarvest;
        private string trayKindHarvest;
        private string harvestOutfile;
        private Sapphirebuisness saph;
        HarvestBuisness hb ;
        SenderToLIMS LiMSsender;


        public Form1()
        {
            //initializes the form. also tries to set a few values based on what machine is running it. If it's one of the mass hunters (the QQQs) or the GCLC than it sets the initial button positions
            //and sets some of the destination folders.
            InitializeComponent();
            Text = String.Format("Agilent Interface ver" + Application.ProductVersion);
            button5.PerformClick();
            string machinename = Environment.MachineName;
            saph = new Sapphirebuisness();
            hb = new HarvestBuisness();
            LiMSsender = new SenderToLIMS();
            //NOTE: FIGURE OUT WHERE THE HARVEST STUFF IS SUPPOSED TO GO. AS IN CDRIVE OR DDRIVE OR SOMETHING ELSE.
            if (machinename == "QQQ-PC" || machinename == "QQQ4" || machinename == "QQQ5" || machinename == "QQQ-HP"|| machinename == "QQQ6490-HP")
            {
                //ONE OF THE MASSHUNTER MACHINES
                testbutton.Visible = false;
                radioButton3.PerformClick();
                radioButton4.PerformClick();
                textBox3.Text = "D:\\MassHunter\\worklist_import\\";
                worklistFolder = "D:\\MassHunter\\worklist_import\\";
                harvestOutfile= "D:\\MassHunter\\worklist_import\\";
                batchfileHarvest = @"C:\VarianOrchard_in\";
                textBox2.Text = batchfileHarvest;
            }
            else if (machinename == "GCLC")
            {
                //the GCLC Machine. aka the outside one.
                testbutton.Visible = false;
                radioButton1.PerformClick();
                radioButton6.PerformClick();
                textBox3.Text = "D:\\XML\\in\\GC\\";
                worklistFolder = "D:\\XML\\in\\GC\\";
                batchfileHarvest = @"C:\VarianOrchard_in\";
                textBox2.Text = batchfileHarvest;
                harvestOutfile = "D:\\XML\\in\\GC\\";
            }
            else
            {
                //any other machine. probably for testing/rewriting/debugging purposes.
                textBox3.Text = "C:\\MassHunter\\worklist_import\\";
                worklistFolder = "C:\\MassHunter\\worklist_import\\";
                batchfileHarvest = @"C:\VarianOrchard_in\";
                textBox2.Text = batchfileHarvest;
                harvestOutfile = @"C:\agilentinterfacetestfiledrop\";//stuff it in a seperate file regardless of masshunter or gclc for testing purposes
            }
            comboBox1.Items.Add("96 well");
            comboBox1.Items.Add("96 well transpose");
            comboBox1.Items.Add("54 well");
            comboBox1.Items.Add("150 well");
            comboBox1.SelectedIndex = 0;
            trayTypeSapphire = "96 well";
            comboBox2.Items.Add("150 well");
            comboBox2.Items.Add("96 well");
            comboBox2.Items.Add("96 well transpose");
            comboBox2.Items.Add("54 well");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // saphire tab quit button
            Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //sapphire tab textbox. Batchid goes here
            batchid = textBox1.Text;
            if (isgaschromo)
            {
                textBox4.Text = "" + batchid + ".xml";

            }
            else
            {
                textBox4.Text = "" + batchid + ".csv";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // sapphire tab send button. builds a csv or xml file
            
            if (ismasshunter)
            {
                saph.buildfileSaphCSV(batchid,isgaschromo,worklistFolder,trayTypeSapphire,trayNumberSapphire);
            }
            if (radioButton1.Checked)
            {
                saph.buildfileXML(batchid,true, worklistFolder,trayTypeSapphire,"1");
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //sapphire tab radiobutton top. corresponds to GC
            
            comboBox3.Items.Clear();
            comboBox3.Items.Add("1");
            comboBox3.SelectedIndex = 0;
            ismasshunter = false;
            isgaschromo = true;
            
            textBox3.Text = "D:\\XML\\in\\GC\\";
            worklistFolder = "D:\\XML\\in\\GC\\";
            isgaschromo = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //sapphire tab radiobutton middle. corresponds to LC
            
            comboBox3.Items.Clear();
            comboBox3.Items.Add("1");
            comboBox3.SelectedIndex = 0;
            ismasshunter = false;
            isgaschromo = false;
            
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            //sapphire tab radiobutton bottom. corresponds to masshunter
            
            comboBox3.Items.Clear();
            comboBox3.Items.Add("1");
            comboBox3.Items.Add("2");
            comboBox3.SelectedIndex = 0;
            ismasshunter = true;
            isgaschromo = false;
            
            textBox3.Text = "D:\\MassHunter\\worklist_import\\";
            worklistFolder = "D:\\MassHunter\\worklist_import\\";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //sapphire combobox neurotransmitter tray
            trayTypeSapphire = comboBox1.Text;
            
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            //harvest tab top radiobutton
            ismasshunter = false;
            isgaschromo = true;
            
            
            comboBox2.SelectedIndex = 0;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            //harvest tab middle radiobutton
            ismasshunter = false;
            isgaschromo = false;
            
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            //harvest tab bottom radiobutton
            ismasshunter = true;
            isgaschromo = false;
            
            
            comboBox2.SelectedIndex = 1;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Harvest tab close button
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Harvest send run list button.
            
            string filekind;
            if (ismasshunter)
            {
                filekind = "csv";
            }
            else
            {
                filekind = "xml";
            }
            hb.buildfileOrchard( ismasshunter, isgaschromo, batchfileHarvest, trayKindHarvest,filekind, harvestOutfile);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //refresh bnutton for harvest file selection.
            try
            {
                listBox1.Items.Clear();
                DirectoryInfo dinfo = new DirectoryInfo("C:\\VarianOrchard_In\\");
                FileInfo[] csvFiles = dinfo.GetFiles("*.csv");
                foreach (FileInfo f in csvFiles)
                {
                    listBox1.Items.Add(f.Name);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("C:\\VarianOrchard_In\\  Does not exist. Please create the file" );
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            worklistFolder = textBox3.Text;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //results tab quit button
            Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //results tab send to Limis button
            
            if (isSapphireBatch())
            {
                if (LiMSsender.sendtoSaphhire(dataGridView1))
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.Rows.Clear();
                }
            }
            else
            {
                if (LiMSsender.sendtoHarvest(dataGridView1))
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.Rows.Clear();
                }
            }

        }

        private bool isSapphireBatch()
        {
            int matchSapphire = 0;
            int matchHarvest = 0;
            Regex rx1 = new Regex(@"^[B-W]\d\d\d\d\d\d");
            Regex rx2 = new Regex(@"^\d\d\d\d\d\d");
            Regex rx3 = new Regex(@"^[!0-9][!0-9]");
            //check what kinf of data we have by looking through the sample names/IDs.
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string value = dataGridView1.Rows[i].Cells[0].Value.ToString();
                switch (value.ToLower().Trim())
                {
                    case "blk":
                    case "blank":
                    case "blnk":
                        //skip blanks.
                        break;
                    default:
                        if(value.Contains(" "))
                        {
                            //no spaces allowed.
                        }
                        else if (rx1.IsMatch(value))
                        {
                            matchSapphire += 1;
                        }
                        else if (rx2.IsMatch(value))
                        {
                            matchHarvest += 1;
                        }
                        else if (rx2.IsMatch(value))
                        {
                            matchHarvest += 1;
                        }
                        break;
                }
            }
            return (matchSapphire > matchHarvest);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //results tab Read Batch Output Button
            OpenFileDialog fileDialog1 = new OpenFileDialog();
            fileDialog1.InitialDirectory = @"F:\DDINEW\OracleImport";
            fileDialog1.DefaultExt = "xlsx";
            fileDialog1.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|Text files(*.txt)|*.txt";
            var a = fileDialog1.ShowDialog();
            if(a != DialogResult.OK)
            {
                return;
            }
            //end filedialogue section.
            //MessageBox.Show(fileDialog1.FileName);
            //next is figuring out how to display.
            if(fileDialog1.FileName == null)
            {
                return;
            }
            try
            {
                using (var filestream = File.Open(fileDialog1.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    string pathextention = Path.GetExtension(fileDialog1.FileName);
                    if (pathextention == ".xlsx" || pathextention == ".xls")
                    {
                        //read and display excel file. Use someone else's excel reading library, pulled from nuGet
                        string paramListID = "";//this var keeps track of what additional calculations we may need .
                        var dataset = ExcelReaderFactory.CreateReader(filestream).AsDataSet();
                        dataGridView1.AutoGenerateColumns = true;
                        dataGridView1.DataSource = dataset;
                        //We only want the sample name and the final Conc.s
                        //first find all the names.
                        List<string> names = new List<string>();
                        for(int i =0; i < dataset.Tables[0].Columns.Count; i++)
                        {
                            string s = dataset.Tables[0].Rows[0].ItemArray[i].ToString().Trim();
                            if (s != "" && s != null && !s.Contains("ISTD") && !s.Contains("Qualifier"))
                            {
                                // ignore if containts ISTD also ignore if qualifier.
                                //translate some stuff.
                                if(s.ToUpper() == "25-OH VITAMIN D2 RESULTS") { s = "VitaminD2"; }
                                if(s.ToUpper() == "25-OH VITAMIN D3 RESULTS") { s = "VitaminD3"; }
                                if(s.ToUpper() == "3-EPI-25-HYDROXY-VITD3 RESULTS") { s = "VitaminDC3"; }
                                s = s.Replace("Results", "").Replace("Sample Sample Name", "Sample").Trim();
                                names.Add(s);
                                //remove all instances of results and get rid of extraneous stuff if nescessary.
                                //also update the paramlistID if nescessary
                                if (s.ToLower().Contains("porphryn")) { paramListID = "Urine Porphryns"; }
                                else if (s.ToLower().Contains("s-adeno") || s.ToLower().Contains("s-adenos")) { paramListID = "SAM_SAH"; }
                                else
                                {
                                    switch (s.ToLower())
                                    {
                                        case "mannitol":
                                        case "lactulose":
                                            paramListID = "Intest. Permeability";
                                            break;
                                        case "norepinephrine":
                                        case "epinephrine":
                                            paramListID = "Neurotransbasic";
                                            break;
                                    }
                                }
                            }
                        }
                        //delete all columns without name and is not a conc. measurement.
                        
                        int n = 0;
                        
                        int c = dataset.Tables[0].Columns.Count;
                        //ensure we go through everything.
                        for(int i = 0; i < dataset.Tables[0].Rows.Count; i++)
                        {
                            if(dataset.Tables[0].Rows[i].ItemArray[4].ToString().ToLower() == "cal")
                            {
                                dataset.Tables[0].Rows.RemoveAt(i);
                                i--;
                            }
                            
                        }
                        for (int i = 0; i < c; i++)
                        {
                            string s = dataset.Tables[0].Rows[1].ItemArray[n].ToString().Trim();

                            switch (s)
                            {
                                case "Name":
                                case "Sample Name":
                                case "Final Conc.":
                                case "Conc. [ ppb ]":
                                case "Conc. [ ppm ]":
                                    //only keep the ones we want.
                                    n++;
                                    if (s.Contains("ppb"))
                                    {
                                        names[n] = names[n] + " Conc. [ ppb ]";
                                    }
                                    if (s.Contains("ppm"))
                                    {
                                        names[n] = names[n] + " Conc. [ ppm ]";
                                    }
                                    break;
                                default:
                                    dataset.Tables[0].Columns.RemoveAt(n);
                                    break;

                            }
                        }
                        //remove top two rows (theyr're the labels we've already gone through and no longer need.)
                        int tempholder = names.Count;
                        //commented out code useful for debugging
                        //MessageBox.Show("numberofNames: " + tempholder + "\n number of columns left: " + dataset.Tables[0].Columns.Count);
                        if (tempholder < dataset.Tables[0].Columns.Count)
                        {
                            // if we don't have enough names, there probably wasn't one for "Sample Name". Add it now to the front.
                            names.Insert(0, "Sample Name");
                        }
                        dataset.Tables[0].Rows.RemoveAt(0);
                        dataset.Tables[0].Rows.RemoveAt(0);
                        for (int i = 0; i < n; i++)
                        {
                            if (i == 0)
                            {
                                dataset.Tables[0].Columns[i].ColumnName = "Sample Name";
                            }
                            else
                            {
                                dataset.Tables[0].Columns[i].ColumnName = names[i].ToString();
                            }
                        }
                        

                        //call the calculated columns here (pass the dataset by reference.)
                        addCalculatedFieldColumns(paramListID, ref dataset);
                        for (int i = 0; i < dataset.Tables[0].Rows.Count; i++)
                        {
                            var row = dataset.Tables[0].Rows[i];
                            for (int j = 1; j < row.ItemArray.Count(); j++)
                            {
                                string s = Math.Round(double.Parse(dataset.Tables[0].Rows[i].ItemArray[j].ToString()), 3).ToString();
                                dataset.Tables[0].Rows[i][j] = s;
                                dataset.AcceptChanges();
                            }
                        }
                        //NOTE VERY MUCH NOT DONE YET.
                        //ADDENUM. MOSTLY DONE. NEED TO GET RID OF CALIBRATION BITS (by qing request).
                        dataGridView1.DataMember = dataset.Tables[0].ToString();
                    }
                    else
                    {
                        //parse the txt file
                        txtparser t = new txtparser();
                        var txtTable = t.parseTextFile(filestream);
                        if (txtTable == null)
                        {
                            //have error. already displayed message.
                            return;
                        }
                        DataSet d = new DataSet();
                        d.Tables.Add(txtTable);
                        for (int i = 0; i < d.Tables[0].Rows.Count; i++)
                        {
                            var row = d.Tables[0].Rows[i];
                            for (int j = 2; j < row.ItemArray.Count(); j++)
                            {
                                string s = Math.Round(double.Parse(d.Tables[0].Rows[i].ItemArray[j].ToString()), 3).ToString();
                                d.Tables[0].Rows[i][j] = s;
                                d.AcceptChanges();
                            }
                        }
                        dataGridView1.AutoGenerateColumns = true;
                        dataGridView1.DataSource = d;
                        //MessageBox.Show(d.Tables[0].Columns.Count.ToString());
                        
                        dataGridView1.DataMember = d.Tables[0].ToString();
                        dataGridView1.Columns[0].Width = 36;
                    }
                    
                }
            }
            catch(Exception ex)
            {
                //probably already in use. make a popup saying so.
                MessageBox.Show(ex.Message);

            }
            
            
        }

        
        private void button9_Click(object sender, EventArgs e)
        {
            //this button exists for test purposes.
            
            
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            trayNumberSapphire = comboBox3.Text
;        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Harvest tab list box. This is the one that updates with all the csv files in 'C:\VarianOrchard_in\'.
            var a = listBox1.SelectedItem;
            if (a != null) {
                batchfileHarvest = @"C:\VarianOrchard_in\" + a.ToString();
                textBox2.Text = batchfileHarvest;
            }
            else
            {
                batchfileHarvest = null;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //Harvest tab batch file location box. (we're gonna try to read from here.)
            batchfileHarvest = textBox2.Text;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Harvest tab Tray combobox/dropdown
            trayKindHarvest = comboBox2.Text;
        }
        
        /// <summary>
        /// This function does way too much
        /// </summary>
        /// <param name="paramlistId"></param>
        /// <param name="data"></param>
        /// 

        private void addCalculatedFieldIntestPerm( ref DataSet data)
        {
            var cols = data.Tables[0].Columns;
            var rows = data.Tables[0].Rows;
            long subsampleIDCol = -1;
            string colname;

            //PercRecovLactulose = ([Volume;Adjusted] / 1000) * [Lactulose;Adjusted]  * 100 / 5000
            //PercRecovMannitol = ([Volume;Adjusted] / 1000) * [Mannitol;Adjusted]   * 100 / 1000
            //Ratio = [PercRecovLactulose;Adjusted] / [PercRecovMannitol;Adjusted]
            long lactuloseCol = -1;
            long mannitolCol = -1;
            long percRecovLactuloseCol = -1;
            long percRecovMannitolCol = -1;
            long ratioCol = -1;
            long volumeCol = -1;
            for (int i = 0; i < cols.Count; i++)
            {
                colname = cols[i].ColumnName;
                //find columns if they exist
                switch (colname.ToLower())
                {
                    case "name":
                    case "sample name":
                        subsampleIDCol = i;
                        break;
                    case "lactulose":
                        lactuloseCol = i;
                        break;
                    case "mannitol":
                        mannitolCol = i;
                        break;
                    case "percRecovLactulose":
                        percRecovLactuloseCol = i;
                        break;
                    case "percRecovMannitol":
                        percRecovMannitolCol = i;
                        break;
                    case "ratio":
                        ratioCol = i;
                        break;
                    case "volume":
                        volumeCol = i;
                        break;
                }

            }
            //NOTE: if I interperet the vb 6 code literally, it should mean that unless the whateverCol is 1, the code runs, because that's how int to bool conversion works. Not sure if intended.
            if (percRecovLactuloseCol == -1)
            {
                data.Tables[0].Columns.Add("PercRecovLactulose", typeof(string));
                percRecovLactuloseCol = data.Tables[0].Columns["PercRecovLactulose"].Ordinal;
            }
            if (percRecovMannitolCol == -1)
            {
                data.Tables[0].Columns.Add("PercRecovMannitol", typeof(string));
                percRecovMannitolCol = data.Tables[0].Columns["PercRecovMannitol"].Ordinal;
            }
            if (ratioCol == -1)
            {
                data.Tables[0].Columns.Add("Ratio", typeof(string));
                ratioCol = data.Tables[0].Columns["Ratio"].Ordinal;
            }
            if (volumeCol == -1)
            {
                data.Tables[0].Columns.Add("Volume", typeof(string));
                volumeCol = data.Tables[0].Columns["Volume"].Ordinal;
            }
            string lactuloseVal;
            string mannitolVal;

            if (lactuloseCol > 0 && mannitolCol > 0)
            {
                for (int rw = 0; rw < data.Tables[0].Rows.Count; rw++)
                {
                    lactuloseVal = data.Tables[0].Rows[rw].ItemArray[lactuloseCol].ToString();
                    mannitolVal = data.Tables[0].Rows[rw].ItemArray[mannitolCol].ToString();
                    string volumeVal = "";
                    string subsampleID = "";
                    if (subsampleIDCol > -1)
                    {
                        subsampleID = data.Tables[0].Rows[rw].ItemArray[subsampleIDCol].ToString();
                        //query inside a for loop. See if fixable. database ops are expensive.
                        volumeVal = getDataSetValueForThisSubsample(subsampleID, "Volume", "Adjusted");
                        data.Tables[0].Rows[rw]["Volume"] = volumeVal;
                    }

                    string percRecovLactuloseVal = "";
                    string percRecovMannitolVal = "";
                    //CONTINUE FROM HERE.NOTE: use try catch blocks to skip possible errors in the math.https://stackoverflow.com/questions/15588249/ignore-error-and-continue-in-c-sharp
                    //NOTE: the above is a code smell. not sure if It should be written this way.
                    try { percRecovLactuloseVal = (double.Parse(volumeVal) / 1000 * double.Parse(lactuloseVal) * 100 / 5000).ToString(); }
                    catch { }
                    try { percRecovMannitolVal = (double.Parse(volumeVal) / 1000 * double.Parse(mannitolVal) * 100 / 1000).ToString(); }
                    catch { }
                    double tmpdoub;
                    if (double.TryParse(percRecovMannitolVal, out tmpdoub))
                    {
                        data.Tables[0].Rows[rw]["PercRecovMannitol"] = SignificantDigitRound(tmpdoub, 6).ToString();

                    }
                    if (double.TryParse(percRecovLactuloseVal, out tmpdoub))
                    {
                        data.Tables[0].Rows[rw]["PercRecovLactulose"] = SignificantDigitRound(tmpdoub, 6).ToString();
                    }
                    string ratioValue = "";
                    try { ratioValue = (double.Parse(percRecovLactuloseVal) / double.Parse(percRecovMannitolVal)).ToString(); }
                    catch { }
                    if (double.TryParse(ratioValue, out tmpdoub))
                    {
                        data.Tables[0].Rows[rw]["Ratio"] = SignificantDigitRound(tmpdoub, 6).ToString();
                    }
                }
            }
        }
        private void addCalculatedFieldNeurotransbasic( ref DataSet data)
        {
            var cols = data.Tables[0].Columns;
            var rows = data.Tables[0].Rows;
            
            string colname;
            long NorepinephrineCol = -1;
            long EpinephrineCol = -1;
            long Norep_Epinep_ratioCol = -1;

            for (int i = 0; i < cols.Count; i++)
            {
                colname = cols[i].ColumnName;
                //find columns if they exist
                switch (colname.ToLower())
                {
                    case "norepinephrine":
                        NorepinephrineCol = i;
                        break;
                    case "epinephrine":
                        EpinephrineCol = i;
                        break;
                    case "norep_epinep_ratio":
                        Norep_Epinep_ratioCol = i;
                        break;
                }
            }

            if (EpinephrineCol > 0 && NorepinephrineCol > 0 && Norep_Epinep_ratioCol < 0)
            {
                data.Tables[0].Columns.Add("Norep_Epinep_ratio", typeof(double));
                Norep_Epinep_ratioCol = data.Tables[0].Columns["Norep_Epinep_ratio"].Ordinal;
            }
            if (EpinephrineCol > 0 && NorepinephrineCol > 0 && Norep_Epinep_ratioCol > 0)
            {
                for (int rw = 0; rw < data.Tables[0].Rows.Count; rw++)
                {
                    string norephineprineVal = data.Tables[0].Rows[rw].ItemArray[NorepinephrineCol].ToString();
                    string ephineprineVal = data.Tables[0].Rows[rw].ItemArray[EpinephrineCol].ToString();
                    if (norephineprineVal != "" && ephineprineVal != "")
                    {
                        string Norep_Epinep_ratioVal = "";
                        try { Norep_Epinep_ratioVal = (double.Parse(norephineprineVal) / double.Parse(ephineprineVal)).ToString(); }
                        catch { }
                        double tempdoub;
                        if (double.TryParse(Norep_Epinep_ratioVal, out tempdoub))
                        {
                            data.Tables[0].Rows[rw]["Norep_Epinep_ratio"] = SignificantDigitRound(tempdoub, 6);
                            data.AcceptChanges();
                        }
                    }
                }

            }
        }
        private void addCalculatedFieldSAM_SAH(ref DataSet data)
        {
            var cols = data.Tables[0].Columns;
            var rows = data.Tables[0].Rows;
            
            string colname;

            long samCol = -1;
            long sahCol = -1;
            long sam_sah_ratioCol = -1;
            for (int i = 0; i < cols.Count; i++)
            {
                colname = cols[i].ColumnName;
                //find columns if they exist
                switch (colname.ToLower())
                {
                    case "s-adnosylmethionine":
                        samCol = i;
                        break;
                    case "s-adenosylhomocysten":
                        sahCol = i;
                        break;
                    case "SAM:SAH":
                        sam_sah_ratioCol = i;
                        break;
                }

            }
            if (samCol > 0 && sahCol > 0 && sam_sah_ratioCol < 0)
            {
                data.Tables[0].Columns.Add("SAM:SAH", typeof(string));
                sam_sah_ratioCol = data.Tables[0].Columns["SAM:SAH"].Ordinal;
            }
            if (samCol > 0 && sahCol > 0 && sam_sah_ratioCol > 0)
            {
                for (int rw = 0; rw < data.Tables[0].Rows.Count; rw++)
                {
                    string samVal = data.Tables[0].Rows[rw].ItemArray[samCol].ToString();
                    string sahVal = data.Tables[0].Rows[rw].ItemArray[sahCol].ToString();
                    if (samVal != "" && sahVal != "")
                    {
                        string sam_sah_ratioVal = "";
                        try { sam_sah_ratioVal = (double.Parse(samVal) / double.Parse(sahVal)).ToString(); }
                        catch { }
                        double tempdoub;
                        if (double.TryParse(sam_sah_ratioVal, out tempdoub))
                        {
                            data.Tables[0].Rows[rw]["SAM:SAH"] = SignificantDigitRound(tempdoub, 6).ToString();
                        }
                    }
                }

            }
        }
        private void addCalculatedFieldColumns(string paramlistId,ref DataSet data)
        {
            
            switch (paramlistId){
                case "Intest. Permeability":
                    addCalculatedFieldIntestPerm( ref data);
                    break;
                case "Neurotransbasic":
                    addCalculatedFieldNeurotransbasic( ref data);

                    break;
                case "SAM_SAH":

                    addCalculatedFieldSAM_SAH(ref data);
                    break;
            }

                    

        }

        private string getDataSetValueForThisSubsample(string subsampleID, string paramID, string paramType)
        {
            Regex rg1 = new Regex(@"^[A-Z]\d\d\d\d\d\d-\d\d\d\d-\d-\d\d-Sa$");

            Regex rg2 = new Regex(@"^[A-Z]\d\d\d\d\d\d-\d\d\d\d-\d-\d\d$");
            string letter = "";
            switch (subsampleID.Substring(0,1).ToUpper()[0])
            {
                case 'U':
                    letter = "Urine";
                    break;
                case 'B':
                    letter = "Blood";
                    break;
                case 'F':
                    letter = "Fecal";
                    break;
                case 'H':
                    letter = "Hair";
                    break;
                case 'W':
                    letter =  "Water";
                    break;
                case 'O':
                    letter = "Other";
                    break;
            }
            List<string> a;
            using (var Con = new OracleConnection(ConfigurationManager.ConnectionStrings["sapphireP"].ConnectionString))
            {
                string sSQL = "SELECT displayvalue FROM sdidataitem where sdcid = '" +letter + "Sample' and" +
                    " paramid = '" + paramID + "' and paramtype = '" + paramType + "' and";

                if (rg1.IsMatch(subsampleID)) { sSQL = sSQL + " keyid1 = '" + subsampleID + "'"; }
                else if (rg2.IsMatch(subsampleID)) { sSQL = sSQL + " keyid1 = '" + subsampleID + "-Sa'"; }
                else { sSQL = sSQL + " keyid1 like '" +(subsampleID + "                ").Substring(0,14) + "%'"; }
   

        
                a = Con.Query<string>(sSQL).ToList();

            }
            string output = "";
            if(a.Count > 0)
            {
                output = a[0];
            }
            return output;
        }
        /// <summary>
        /// rounds a value to the nth significant digit (rounds to nearest number). taking double since most common number types can be assigned to doubles.
        /// </summary>
        /// <param name="value"> The number to round</param>
        /// <param name="n"> The number of digits to round to.</param>
        /// <returns></returns>
        private double SignificantDigitRound(double value, int n)
        {
            //rounds a value to the nth significant digit (rounds to nearest number). taking double since most common number types can be assigned to doubles.
            var a = value.ToString("#." + new String('#', n - 1) + "E000");
            return double.Parse(a);

        }
    }
        
    
}
