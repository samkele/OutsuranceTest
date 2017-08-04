using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutsuranceFilePrinter
{
    public partial class OutsuranceFilePrinter : Form
    {
        public OutsuranceFilePrinter()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            ImportExcelFile();
        }


        private void ImportExcelFile()
        {
            try
            {
                string inputFile;
                var fileDlg = new OpenFileDialog();
                fileDlg.RestoreDirectory = true;
                if (fileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    inputFile = fileDlg.FileName;
                else
                    return;
                var cn = new System.Data.OleDb.OleDbConnection(string.Format(
                @"provider=Microsoft.Jet.OLEDB.4.0;
                Data Source='{0}';
                Extended Properties='Excel 8.0;HDR=YES;IMEX=1';", EscapeConnectionString(inputFile)));
                try
                {
                    cn.Open();
                    //get sheet name
                    string sheet = string.Empty;
                    var schema = cn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                    if (schema.Rows.Count == 0)
                        throw new Exception("The selected workbook does not appear to contain any sheets.");
                  
                    else
                        sheet = schema.Rows[0]["TABLE_NAME"].ToString();
                    var da = new System.Data.OleDb.OleDbDataAdapter(string.Format("select * from [{0}]", sheet), cn);
                    var tbl = new DataTable();
                    da.Fill(tbl);
                    SortAndPrintNameData(tbl);
                    SortAndPrintAddressData(tbl);
                }
                finally
                {
                    cn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SortAndPrintNameData(DataTable tbl)
        {
        
            Dictionary<string, int> WordCountDict = new Dictionary<string, int>();

            foreach (DataRow row in tbl.Rows)
            {
                //if (counter > 0)
                //{
                string first = row[0].ToString();
                string last = row[1].ToString();


                if (!WordCountDict.ContainsKey(first))
                {
                    WordCountDict.Add(first, 1);
                }
                else
                {
                    WordCountDict[first]++;
                }
                if (!WordCountDict.ContainsKey(last))
                {
                    WordCountDict.Add(last, 1);
                }
                else
                {
                    WordCountDict[last]++;
                }
            }
         

            var SortedWordCount = from pair in WordCountDict
                                  orderby pair.Value descending,
                                  pair.Key
                                  select pair;

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\Names.txt"))
            {
                foreach (KeyValuePair<string, int> kvp in SortedWordCount)
                {
                    file.WriteLine(kvp.Key + ',' + kvp.Value);
                }
            }

        }


        public void SortAndPrintAddressData(DataTable tbl)
        {

            Dictionary<string, int> WordCountDict = new Dictionary<string, int>();

            List<string> StreetNameFirstList = new List<string>();

            foreach (DataRow row in tbl.Rows)
            {
                string[] splitAddress = row[2].ToString().Split(null, 2);

                string StreetNameFirst = splitAddress[1] + " " + splitAddress[0];
                StreetNameFirstList.Add(StreetNameFirst); 

              
            }


            StreetNameFirstList.Sort();

            List<string> SortedStreetNameList = new List<string>();

            foreach (string s in StreetNameFirstList)
            {

                int lastIndex = s.LastIndexOf(' ');
                if (lastIndex + 1 < s.Length)
                {
                    string firstPart = s.Substring(0, lastIndex);
                    string secondPart = s.Substring(lastIndex + 1);

                    string StreetNameFirst = secondPart + " " + firstPart;

                    SortedStreetNameList.Add(StreetNameFirst);

                }

          }

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.IO.Path.GetDirectoryName(Application.ExecutablePath) + "\\Addresses.txt"))
            {
                foreach (string address in SortedStreetNameList)
                {
                    file.WriteLine(address);
                }
            }

        }


        /// <summary>
        /// Currenlty only escapes the ' char by doubling up
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        internal static string EscapeConnectionString(string input)
        {
            return Regex.Replace(input, "'{1}", "''");
        }
    }
}

