using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace Excel2Sql
{
    class Program
    {
        static void Main(string[] args)
        {
            //Trying to figure out the relative path: 
            //string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            //Console.WriteLine(path);
            //Console.ReadLine();

            // Excel File Path Location
            //Relative
            //string excelfilepath = @".\College Football Statistics.xlsx";//For some reason relative path won't work, 
            //keep getting the message "Cannot update.  Database or object is read-only."  

            //Absolute
            string excelfilepath = @"C:\Users\Ben\source\repos\Excel2Sql\College Football Statistics.xlsx";
            //Excel file has to be open on my computer in order for it to work, maybe that is why the relative path 
            //above is not working.  
            
            //Trying a different version for xlsx excel files (This didn't fix my problem, just changing xls to xlsx above did): 
            //public static string path = @"C:\src\RedirectApplication\RedirectApplication\301s.xlsx";
            //public static string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelfilepath + ";Extended Properties=Excel 12.0;";

            // SQL Server TableName
            string TableName = "Power5Records";

            string exceldataquery = "select School,Nickname,Conference,Year,Wins,Losses,WinPerc from [Records$]";

            try
            {
                // Excel Connection String and SQL Server Connection String
                string excelconnectionstring = @"provider=microsoft.jet.oledb.4.0;
                      data source=" + excelfilepath +
                      ";extended properties=" + "\"excel 4.0;hdr=yes;\"";
                //“hdr=yes;” indicates that the first row contains column names, not data. “hdr=no;” indicates the opposite.
                //For some reason, it does skip row 2 of the excel sheet, so I've moved my data down a line in order to get it all to transfer.
                string sqlconnectionstring = @"Data Source=(LocalDb)\MSSQLLocalDB;Initial Catalog=CollegeFootball;Integrated Security=True; 
                    database = CollegeFootball; connection reset = false";

                //Execute A Query To Erase Any Previous Data From Power5Records Table
                string deletesqlquery = "delete from " + TableName;
                SqlConnection sqlconn = new SqlConnection(sqlconnectionstring);
                SqlCommand sqlcmd = new SqlCommand(deletesqlquery, sqlconn);

                sqlconn.Open();
                sqlcmd.ExecuteNonQuery();
                sqlconn.Close();

                // Build A Connection To Excel Data Source And Execute The Command
                OleDbConnection oledbconn = new OleDbConnection(excelconnectionstring);
                OleDbCommand oledbcmd = new OleDbCommand(exceldataquery, oledbconn);
                oledbconn.Open();
                OleDbDataReader dr = oledbcmd.ExecuteReader();

                // Connect To SQL Server DB And Perform a Bulk Copy Operation
                SqlBulkCopy bulkcopy = new SqlBulkCopy(sqlconnectionstring);

                // Provide Excel To Table Column Mapping If Any Difference In Name
                bulkcopy.ColumnMappings.Add("School", "School");
                bulkcopy.ColumnMappings.Add("Nickname", "Nickname");
                bulkcopy.ColumnMappings.Add("Conference", "Conference");
                bulkcopy.ColumnMappings.Add("Year", "Year");
                bulkcopy.ColumnMappings.Add("Wins", "Wins");
                bulkcopy.ColumnMappings.Add("Losses", "Losses");
                bulkcopy.ColumnMappings.Add("WinPerc", "WinPerc");

                // Provide The Table Name For Bulk Copy
                bulkcopy.DestinationTableName = TableName;

                while (dr.Read())
                {
                    bulkcopy.WriteToServer(dr);
                }

                oledbconn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        
        }
    }
}
