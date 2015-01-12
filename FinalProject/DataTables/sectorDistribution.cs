using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class sectorDistribution : reportHandling
    {
        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                table.Columns.Add("Sl.No.", typeof(string));
                table.Columns.Add("Sector", typeof(string));
                table.Columns.Add("No. of Units", typeof(string));

                OleDbDataReader dr;
                string strCommand ="";
                string[] query = new string[6] { "TypeofUnit = 'Jute Mill'", "TypeofUnit <> 'Jute Mill' and TypeofUnit <> 'Non-Mill Sector'",
                    "TypeofUnit = 'Non-Mill Sector'","where TypeofProducer LIKE 'Producer%'","where TypeofProducer LIKE 'Merchant%'",""};
                string[] cellValue = new string[6] { "Jute Mill Sector", "Other Mill Sector", "Non Mill Sector",
                    "JDP Manufacturing Units","Merchant Exporter","Total"};
              
                for(int i = 0 ; i < query.Length ; i++)
                {
                    if (i < 3)
                    {
                        strCommand = "Select count(TypeofUnit) from GeneralDetails where " + query[i];
                        dr = getData(conn, strCommand);
                        dr.Read();
                        table.Rows.Add((i + 1).ToString(), cellValue[i], dr[0].ToString());
                    }
                    else
                    {
                        strCommand = "Select count(TypeofProducer) from GeneralDetails " + query[i];
                        dr = getData(conn, strCommand);
                        dr.Read();
                        table.Rows.Add("", cellValue[i], dr[0].ToString());
                    }
                    dr.Close();
                    dr.Dispose();
                }

                wordHandling.openDocument();
                wordHandling.writetoWord(table, "SECTOR-WISE DISTRIBUTION OF SURVEYED UNITS (as on 31st March '13)", "Table 1", 0, 1);
                wordHandling.CloseandSave();
            }
        }
    }
}
