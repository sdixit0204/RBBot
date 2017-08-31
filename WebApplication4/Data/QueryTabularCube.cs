using System;
using System.IO;
using System.Collections.Generic;

using System.Linq;

using System.Text;

using System.Threading.Tasks;

using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.AdomdClient;
using System.Data.OleDb;
using System.Data;

namespace WebApplication4
{
    class QueryTabularCube

    {
        public static string final_result;
        public static string error ;

        public static String  call_DB()
        {
            DataTable table = new DataTable();

            using (OleDbConnection connection = new OleDbConnection(@"Provider=MSOLAP.8;Password=Egypt@123;User ID=shweta.dixit@hitachiconsulting.com;Persist Security Info=True;Initial Catalog=RBSample;Data Source=asazure://southcentralus.asazure.windows.net/rbonedev;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Update Isolation Level=2")) 
            {
                // The insertSQL string contains a SQL statement that 
                // inserts a new row in the source table.
                 //sqlquery = @"SELECT
                 //               {
                 //               [Measures].[Total Sellout Order Amount]
                 //                   }
                 //                   ON COLUMNS,
                 //               {
                 //                      [Store Order].[City].&[AMPAYON]
                 //               }
                 //               ON ROWS
                 //               FROM[Model]
                 //               WHERE[Order Date].[Date].&[2017-06-01T00:00:00]";
                string sqlquery = @"SELECT
                                    (
                                    [Product].[Price],
                                    [Product].[Product SKU Description].&[Aerosol Precious Silk 6x 2x3]
                                    )ON COLUMNS,
                                    {
                                           [Store Order].[City].&[AMPAYON]
                                    } ON ROWS
                                    FROM [Model]
                                    WHERE [Order Date].[Date].&[2017-06-01T00:00:00]

                                    ";


                OleDbCommand command = new OleDbCommand(sqlquery);


                // Set the Connection to the new OleDbConnection.
                command.Connection = connection;
                
                OleDbDataAdapter oledbAdapter;
                // Open the connection and execute the insert command. 
                try
                {
                    connection.Open();
                    oledbAdapter = new OleDbDataAdapter(sqlquery, connection);
                    oledbAdapter.Fill(table);
                    for (int i = 0; i <= table.Rows.Count - 1; i++)
                    {
                        final_result += "$ " + table.Rows[i][1].ToString();
                    }
                   
                    oledbAdapter.Dispose();
                    connection.Close();
                   
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    final_result = ex.Message;
                    

                }
                // The connection is automatically closed when the 
                // code exits the using block.
            }


            return final_result;
        }
        static void Main(string[] args)

        {
        }

    }

}