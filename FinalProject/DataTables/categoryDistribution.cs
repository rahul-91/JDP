using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class categoryDistribution : reportHandling
    {
        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                table.Columns.Add("Code.", typeof(string));
                table.Columns.Add("Category of Unit", typeof(string));
                table.Columns.Add("No.s", typeof(string));
                table.Columns.Add("Percentage (%)", typeof(double));

                OleDbDataReader dr;
                double total = 0, subtotal = 0;
                double percent;
                string[] typeofProducer = new string[5]{"Producer of only JDPs","Producer of JDPs and other products",
                    "Producer-cum-Exporter of only JDPs","Producer-cum-Exporter of JDPs and other products",
                    "Merchant Exporter of JDPs "};
               
                string strCommand = "Select count(*) from GeneralDetails";
                dr = getData(conn, strCommand);
                dr.Read();
                total = int.Parse(dr[0].ToString());
                for(int i = 0 ; i < typeofProducer.Length ; i++)
                {
                    strCommand = "Select count(TypeofProducer) from GeneralDetails where TypeofProducer = '" + typeofProducer[i] + "'";
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round(double.Parse(dr[0].ToString()) / total * 100, 2);
                    subtotal += int.Parse(dr[0].ToString());
                    table.Rows.Add(Convert.ToChar(i+65).ToString(),typeofProducer[i], dr[0].ToString(), percent);
                    if (i == 3)
                    {
                        percent = Math.Round(subtotal / total * 100, 2);
                        table.Rows.Add("", "SubTotal", subtotal.ToString(), percent);
                    }
                    dr.Close();
                    dr.Dispose();
                }

                dr.Close();
                dr.Dispose();
                percent = Math.Round(subtotal / total * 100, 2);
                table.Rows.Add("", "Total", subtotal.ToString(), percent);

                wordHandling.openDocument();
                wordHandling.writetoWord(table, "DISTRIBUTION OF SURVEYED UNITS BY CATEGORY ALL SECTOR (As on 31st March '13)", "Table 2A", 0, 1);
                wordHandling.CloseandSave();
            }
        }
    }
}
