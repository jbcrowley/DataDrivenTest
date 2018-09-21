using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace RetailChannels.ATG.Common
{
    public class ExcelOleDbReader
    {
        /// <summary>
        /// Converts specified Worksheet in a Microsoft Excel Workbook into a <see cref="DataTable"/> object.
        /// </summary>
        /// <param name="fileName">Full path and file name of the Microsoft Excel workbook to connect to.</param>
        /// <param name="sheetName">The name of the Excel Worksheet to retreive data from.</param>
        /// <returns>Returns a <see cref="DataTable"/> with the contents of the active worksheet.</returns>
        public DataTable ImportExcelFile(string fileName, string sheetName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException("fileName", "The filename must not be empty or null.");
            }
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException("sheetName", "The sheet name must not be empty or null.");
            }
            if (!File.Exists(fileName))
            {
                throw new ArgumentNullException("fileName", $"The file <{fileName}> does not exist.");
            }

            DataTable result = new DataTable();
            using (OleDbConnection connection = GetWorkbookConnection(fileName))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "$]", connection);
                adapter.Fill(result);
            }

            return result;
        }
        /// <summary>
        /// Creates a <see cref="OleDbConnection"/> based on the specified Microsoft Excel workbook.
        /// </summary>
        /// <param name="fileName">Full path and file name of the Microsoft Excel workbook to use as the data source.</param>
        /// <returns>Returns a <see cref="OleDbConnection"/>.</returns>
        private OleDbConnection GetWorkbookConnection(string fileName)
        {
            // http://stackoverflow.com/a/9274455/2386774
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException("fileName", "The filename must not be empty or null.");
            }

            OleDbConnectionStringBuilder connectionString = new OleDbConnectionStringBuilder();
            connectionString.DataSource = fileName;
            string extendedProperties = String.Empty;
            if (Path.GetExtension(fileName).Equals(".xls")) // for 97-03 Excel file
            {
                connectionString.Provider = "Microsoft.Jet.OLEDB.4.0";
                extendedProperties = "Excel 8.0";
            }
            else if (Path.GetExtension(fileName).Equals(".xlsx"))  // for 2007 Excel file
            {
                // Install 2007 Office System Driver: Data Connectivity Components
                // https://www.microsoft.com/en-us/download/details.aspx?id=23734
                connectionString.Provider = "Microsoft.ACE.OLEDB.12.0";
                extendedProperties = "Excel 12.0";
            }
            else
            {
                throw new ArgumentException("The file must be of type .xls or .xlsx.");
            }
            connectionString.Add("Extended Properties", extendedProperties + ";HDR=Yes;IMEX=1"); // HDR=ColumnHeader,IMEX=InterMixed

            return new OleDbConnection(connectionString.ToString());
        }
    }
}