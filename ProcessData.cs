using System;

using System.Data;

namespace shareOkForms
{
    class ProcessData
    {
        public static bool processHandleData(String[] filePaths)
        {
            foreach(String filepath in filePaths)
            {
                string htmlCode = "";
                int index = 0;
                String searchResult = "";
                DataTable metadata_table = new DataTable();
                ReadExcel readExcel = new ReadExcel();
                // metadata_table = readExcel.ReadExcelFile(@"C:\Users\ponnaga\source\repos\shareOkForms\shareOkForms\rm_check_262.xlsx", "xlsx");
                metadata_table = readExcel.ReadExcelFile(filepath, "xlsx");
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
                metadata_table.Columns.Add("isFileValid", typeof(String));
                //removing unwanted columns
                DataView view = new DataView(metadata_table);
                DataTable filtered_metadata_table = view.ToTable(false, "id", "collection", "dc.contributor.author", "dc.identifier.uri", "dc.title", "osu.filename", "isFileValid");
                foreach (DataRow row in filtered_metadata_table.Rows)
                {
                    try
                    {
                        htmlCode = Form1.client.DownloadString(row["dc.identifier.uri"].ToString());
                        searchResult = htmlCode.IndexOf(row["osu.filename"].ToString()) == -1 ? "No" : "Yes";

                    }
                    catch
                    {
                        searchResult = "No";
                    }
                    finally
                    {
                        row["isFileValid"] = searchResult;
                        //row["Status"] = constructed_uri=constructFileURL.ConstructURL(row["dc.identifier.uri"].ToString(),row["osu.filename"].ToString());
                        index++;
                       // label1.Text = "Procesing record " + index + " of " + filtered_metadata_table.Rows.Count;
                       // label1.Refresh();
                    }

                }
                //binding it to the dataview
              //  dataGridView1.Visible = true;
               // dataGridView1.DataSource = filtered_metadata_table;
                //some helper variables
                int LastIndex = filepath.LastIndexOf('\\');
                int PathLength = filepath.Length;
                int LastIndexFileExtension = filepath.LastIndexOf('.');
                String Path_without_fileName = filepath.Substring(0, LastIndex);
                //String fileName = openFileDialog1.FileName.Substring(LastIndex, PathLength - LastIndex-1);
                String fileName_without_extension = filepath.Substring(LastIndex, LastIndexFileExtension - LastIndex);
                //writing output to excel
                WriteToExcel writeToExcel = new WriteToExcel();
                writeToExcel.writeExcel(filtered_metadata_table, fileName_without_extension + "_output.xlsx", Path_without_fileName + "\\output");
            }
            return true;
        }
    }
}
