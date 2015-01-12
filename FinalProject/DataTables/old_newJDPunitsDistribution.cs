using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class old_newJDPunitsDistribution : reportHandling
    {
        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                table.Columns.Add("Sl. No.", typeof(string));
                table.Columns.Add("State", typeof(string));
                table.Columns.Add("Old", typeof(string));
                table.Columns.Add("New", typeof(string));
                table.Columns.Add("Total", typeof(string));
                
                OleDbDataReader dr;
                string strCommand = "",total = "";
                string[] age = new string[2] { "OD", "NW"};
                string[] no_of_units = new string[3];

                for (int i = 0; i < stateCode.Length; i++)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        strCommand = "Select count(*) from GeneralDetails where QreID LIKE '" + stateCode[i] + "/" + age[j] + "%'";
                        dr = getData(conn, strCommand);
                        dr.Read();
                        no_of_units[j] = dr[0].ToString();
                        dr.Close();
                        dr.Dispose();
                    }
                    total = (int.Parse(no_of_units[0]) + int.Parse(no_of_units[1])).ToString();
                    table.Rows.Add((i + 1).ToString(), states[i], no_of_units[0], no_of_units[1], total);
                }
                for(int i = 0 ; i < 2 ; i++)
                {
                   strCommand = "Select count(*) from GeneralDetails where QreID LIKE '%"+age[i]+"%'";
                   dr = getData(conn, strCommand);
                   dr.Read();
                   no_of_units[i] = dr[0].ToString();
                   dr.Close();
                   dr.Dispose();
                }
             
                total = (int.Parse(no_of_units[0]) + int.Parse(no_of_units[1])).ToString();
                table.Rows.Add("", "All India Total", no_of_units[0], no_of_units[1], total);

                wordHandling.openDocument();
                wordHandling.writetoWord(table, "State Wise Distribution of Old and New JDP Units", "Table 21", 0, 1);
                wordHandling.CloseandSave();
            }
        }
    }
}
