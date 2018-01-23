using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;

namespace LINQLESSON2.Model
{
    class Helper
    {
        public DataSet GetData()
        {
            string conString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            using (SqlConnection con = new SqlConnection(conString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "SELECT * FROM [CRCMS_new].[dbo].[Area];"+
                                   " SELECT * FROM [CRCMS_new].[dbo].[dic_Group];"+
                                    "SELECT * FROM [CRCMS_new].[dbo].[dic_Pavilion];";

                DataSet ds = new DataSet();

                SqlDataAdapter da = new SqlDataAdapter("", con);
                da.SelectCommand = cmd;

                da.TableMappings.Add("Table", "Area");
                da.TableMappings.Add("Table1", "dic_Group");
                da.TableMappings.Add("Table2", "dic_Pavilion");

                da.Fill(ds);

                return ds;

            }
        }
    }
}
