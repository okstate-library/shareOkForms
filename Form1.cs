using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace shareOkForms
{
    public partial class Form1 : Form
    {
        public static readonly WebClient client = new WebClient();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
       
        private void LoadData_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result != DialogResult.OK)
            {
                return;
            }
            string htmlCode = "", htmlCode_1="", htmlCode_2="";
            int index = 0;
            string searchResult = "", searchResult_1 = "", searchResult_2 = "";
            int debug_index = 0;
            DataTable metadata_table = new DataTable();
            ReadExcel readExcel = new ReadExcel();
            // metadata_table = readExcel.ReadExcelFile(@"C:\Users\ponnaga\source\repos\shareOkForms\shareOkForms\rm_check_262.xlsx", "xlsx");
            metadata_table = readExcel.ReadExcelFile(openFileDialog1.FileName, "xlsx");
            //creating column names
            foreach (DataColumn column in metadata_table.Columns)
            {
                string cName = metadata_table.Rows[0][column.ColumnName].ToString();
                if (!metadata_table.Columns.Contains(cName) && cName != "")
                {
                    column.ColumnName = cName;
                }

            }
            metadata_table.Rows[0].Delete();
            metadata_table.Columns.Add("ErrorMessage", typeof(String));
            metadata_table.Columns.Add("isFileValid", typeof(String));
            //removing unwanted columns
            DataView view = new DataView(metadata_table);
            DataTable filtered_metadata_table = view.ToTable(false, "id", "collection", "dc.contributor.author", "dc.identifier.uri", "dc.title", "osu.filename", "isFileValid", "ErrorMessage");
            foreach (DataRow row in filtered_metadata_table.Rows)
            {
                debug_index++;
                try {
                    //case of duplicate handles
                    if (row["dc.identifier.uri"].ToString().IndexOf('|') != -1) {
                        String[] tokens = row["dc.identifier.uri"].ToString().Split(new String[] { "||"},StringSplitOptions.None);
                        try
                        {
                            htmlCode_1 = client.DownloadString(tokens[0]);
                        }
                        catch
                        {
                            searchResult_1 = "404 Page Not found";
                        }
                        try
                        {
                            htmlCode_2 = client.DownloadString(tokens[1]);
                        }
                        catch
                        {
                            searchResult_2 = "404 Page Not found";
                        }
                        if(searchResult_1!= "404 Page Not found"|| searchResult_2 != "404 Page Not found")
                        {
                            if(searchResult_1== "404 Page Not found")
                            {
                                htmlCode = htmlCode_2;
                            }
                            else
                            {
                                htmlCode = htmlCode_1;
                            }
                            if (htmlCode.IndexOf(row["osu.filename"].ToString()) == -1)
                            {
                                if (htmlCode_1.IndexOf(".pdf") == -1 || htmlCode_2.IndexOf(".pdf") == -1)
                                {
                                    searchResult = "No Attachment";
                                }
                                else
                                {
                                    searchResult = "Wrong Attachment";
                                }
                            }
                            else {
                                searchResult = "one of the link works";
                            }
                               
                        }
                        else
                        {
                            searchResult = searchResult_1;//or searchResult_2. doesnt matter. both are same here.
                        }
                    }
                    else
                    {
                        htmlCode = client.DownloadString(row["dc.identifier.uri"].ToString());
                        if (htmlCode.IndexOf(row["osu.filename"].ToString()) == -1)
                        {
                            if (htmlCode.IndexOf(".pdf") == -1)
                            {
                                searchResult = "No Attachment";
                            }
                            else
                            {
                                searchResult = "Wrong Attachment";
                            }
                        }
                        else
                        {
                            searchResult = "No Error";
                        }
                    }
                   
                }
                catch
                {
                    searchResult = "404 Page Not found";
                }
                finally
                {
                    row["ErrorMessage"] = searchResult;
                    index++;
                    label1.Text = "Procesing record " + index + " of " + filtered_metadata_table.Rows.Count;
                    label1.Refresh();
                }
                              
            }
            //binding it to the dataview
            dataGridView1.Visible = true;
            dataGridView1.DataSource = filtered_metadata_table;
            //some helper variables
            int LastIndex= openFileDialog1.FileName.LastIndexOf('\\');
            int PathLength= openFileDialog1.FileName.Length;
            int LastIndexFileExtension = openFileDialog1.FileName.LastIndexOf('.');
            String Path_without_fileName = openFileDialog1.FileName.Substring(0, LastIndex);
            //String fileName = openFileDialog1.FileName.Substring(LastIndex, PathLength - LastIndex-1);
            String fileName_without_extension= openFileDialog1.FileName.Substring(LastIndex, LastIndexFileExtension- LastIndex);
            //writing output to excel
            WriteToExcel writeToExcel = new WriteToExcel();
            writeToExcel.writeExcel(filtered_metadata_table, fileName_without_extension+"_output.xlsx", Path_without_fileName + "\\output");
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dialog_result= folderBrowserDialog1.ShowDialog();
            if (dialog_result != DialogResult.OK)
            {
                return;
            }
            String path = folderBrowserDialog1.SelectedPath;
            String[] fileNames= Directory.GetFiles(path);
            bool is_success=ProcessData.processHandleData(fileNames);
            MessageBox.Show(is_success.ToString());
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }
    }
}
