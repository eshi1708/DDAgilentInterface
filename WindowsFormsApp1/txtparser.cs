using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public class txtparser
    {
        public DataTable parseTextFile(FileStream fileName)
        {
            //parses text file from the GC/LC.
            const string DELIMITER = "|";
            const int MAX_COL = 9;
            
            int[] endOfColumn = new int[MAX_COL + 1];
            string[] columnName = new string[MAX_COL + 1];
            DataColumn[] dataColumns = new DataColumn[MAX_COL + 1];
            string lineMinus2="";
            string lineMinus1="";
            string line="";
            bool startOfResultsDataSection = false;
            int col = 0;
            int delimiterPosition=-1;
            int i;
            string result;
            bool done = false;
            DataTable newTable = new DataTable("newTable");
            try
            {
                using (StreamReader sr = new StreamReader(fileName))
                {
                    while(!sr.EndOfStream && !done)
                    {
                        //keep two lines.
                        lineMinus2 = lineMinus1;
                        lineMinus1 = line;
                        line = sr.ReadLine();
                        if(line.Contains("*** End of Report ***"))
                        {
                            startOfResultsDataSection = false;
                            done = true;
                            //MessageBox.Show(newTable.Columns.Count.ToString());
                        }

                        if (startOfResultsDataSection)
                        {
                            
                            DataRow newrow = newTable.NewRow();
                            for(i = 0; i < col;i++)
                            {
                                if(i == col)
                                {
                                    var a = endOfColumn[i];
                                    var b = endOfColumn[i - 1];
                                    var c = line.Length;
                                    var tmplen = endOfColumn[i] - endOfColumn[i - 1];
                                    result = line.Substring(endOfColumn[i - 1] + 1, line.Length - endOfColumn[i - 1]-1).Trim();
                                }
                                else if (i != 0)
                                {
                                    result = line.Substring(endOfColumn[i - 1] + 1, endOfColumn[i] - endOfColumn[i - 1]).Trim();
                                }
                                else
                                {
                                    result = line.Substring(0, endOfColumn[i]).Trim();
                                }
                                newrow[columnName[i]] = result;
                            }
                            newTable.Rows.Add(newrow);


                        }
                        else if (!done)
                        {
                            
                            do
                            {
                                delimiterPosition = line.IndexOf(DELIMITER, delimiterPosition+1);
                                if (delimiterPosition + 1 > 0)//NOTE: VB6 GIVES 1 INDEXED OUTPUT FOR THE EQUIVALENT OF INDEX OF, SO WE ADD 1 TO THE VALUE.
                                {
                                    startOfResultsDataSection = true;
                                    endOfColumn[col] = delimiterPosition;
                                    if (col != 0)
                                    {
                                        columnName[col] = lineMinus2.Substring(endOfColumn[col - 1] + 1, endOfColumn[col] - endOfColumn[col - 1]).Trim();
                                    }
                                    else
                                    {
                                        columnName[col] = "#*";//maybe we actually need to keep the pound/hash and asterisk, so keep the first/0th column anyways.
                                    }


                                    DataColumn colmn = new DataColumn(columnName[col]);
                                    colmn.DataType = System.Type.GetType("System.String");
                                    newTable.Columns.Add(colmn);
                                    col = col + 1;
                                    
                                    
                                }
                                if (delimiterPosition == -1 && col != 0)
                                {
                                    //pick up the last column if it is not IS.
                                    string tmpstr = lineMinus2.Substring(endOfColumn[col - 1] + 1, line.Length - endOfColumn[col - 1] - 1).Trim();
                                    if (tmpstr != "IS")
                                    {
                                        columnName[col] = tmpstr;
                                        endOfColumn[col] = line.Length;
                                        DataColumn colmn = new DataColumn(columnName[col]);
                                        colmn.DataType = System.Type.GetType("System.String");
                                        newTable.Columns.Add(colmn);
                                    }
                                }
                            } while (delimiterPosition != -1);//delimiter goes to -1 when there is none left.
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("error parsing text file: " + ex.Message);
                return null;
            }
            
            return newTable;

        }
    }
}
