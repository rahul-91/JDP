using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class ownershipStatus : reportHandling
    {
        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                table.Columns.Add("Ownership Type", typeof(string));
                table.Columns.Add("No.s", typeof(string));
                table.Columns.Add("Percentage (%)", typeof(double));

                OleDbDataReader dr;
                double total = 0, subtotal = 0;
                double percent;
                string[] companyStatus = new string[6] { "Proprietorship", "Partnership",
                    "Public/Private Limited", "Co-operative Society","Voluntary Organisation/NGO","Any other"};

                string strCommand = "Select count(*) from GeneralDetails";
                dr = getData(conn, strCommand);
                dr.Read();
                total = int.Parse(dr[0].ToString());

                foreach (string type in companyStatus)
                {
                    strCommand = "Select count(CompanyStatus) from GeneralDetails where CompanyStatus = '" + type + "'";
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round(double.Parse(dr[0].ToString()) / total * 100, 2);
                    subtotal += int.Parse(dr[0].ToString());
                    table.Rows.Add(type, dr[0].ToString(), percent);
                    dr.Close();
                    dr.Dispose();
                }

                dr.Close();
                dr.Dispose();

                percent = Math.Round(subtotal / total * 100, 2);
                table.Rows.Add("Total", subtotal.ToString(), percent); 

                wordHandling.openDocument();
                wordHandling.writetoWord(table, "OWNERSHIP STATUS OF SURVEYED UNITS (As on 31st March’13)", "Table 4", 0, 2);
                wordHandling.CloseandSave();
            }
        }
    }
}
