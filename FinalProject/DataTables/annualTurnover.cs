using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Configuration;

namespace FinalProject
{
    class annualTurnover : reportHandling
    {
        private static string query(int index, string year)
        {
            string[] query = new string[] { "<= 10", " > 10 AND " + year + " <= 25",
                    " > 25 AND " + year + " <= 50"," > 50 AND " + year + " <= 75",
                    " > 75 AND " + year + " <= 100"," > 100 AND " + year + " <= 200",
                    " > 200 AND " + year + " <= 500", " >500", " is null"};
            
            return query[index];
        }
        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                table.Columns.Add("Annual Turnover (in Rs. lakh)", typeof(string));
                table.Columns.Add("2010-11 (Apr.-Mar.)", typeof(string));
                table.Columns.Add("", typeof(string));
                table.Columns.Add("2011-12 (Apr.-Mar.)", typeof(string));
                table.Columns.Add("", typeof(string));
                table.Columns.Add("2012-13 (Apr.-Mar.)", typeof(string));
                table.Columns.Add("", typeof(string));
                table.Rows.Add("", "No.", "%", "No.", "%", "No.", "%");

                OleDbDataReader dr;

                string[] years = new string[3] { "AnnTurnover201011OnJDP", "AnnTurnover201112OnJDP", "AnnTurnover201213POnJDP"};
                string[] range = new string[] { "Upto 10", "10-25", "25-50", "50-75", "75-100", "100-200", 
                    "200-500", "Above 500", "Not Disclosed","Sector Total" };
                
                string[] values = new string[4];
                string[] percent = new string[4];

                string strCommand = "Select count(*) from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID";
                dr = getData(conn, strCommand);
                dr.Read();
                double total = int .Parse(dr[0].ToString());

                table.Rows.Add("Sector: Jute Mill", "", "", "", "", "", "");

                for (int i = 0; i < range.Length; i++)
                {
                    for (int j = 0; j < years.Length; j++)
                     {
                         if (range[i] == "Not Disclosed")
                         {
                             strCommand = "Select count(*) from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit = 'Non-Mill Sector' AND " + years[j] + query(i, years[j]);
                         }
                         else if (range[i] == "Sector Total")
                         {
                             strCommand = "Select count(*) from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit='Non-Mill Sector'";
                         }
                         else
                         {
                             strCommand = "Select count(" + years[j] + ") from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit = 'Non-Mill Sector' AND " + years[j] + query(i, years[j]);
                         }
                        dr = getData(conn, strCommand);
                        dr.Read();
                        values[j] = dr[0].ToString();
                        percent[j] = Math.Round((double.Parse(values[j]) / total) * 100, 2).ToString();
                        dr.Close();
                        dr.Dispose();
                    }
                    table.Rows.Add(range[i], values[0], percent[0], values[1], percent[1], values[2], percent[2]);
                }

                table.Rows.Add("Sector: Any Other", "", "", "", "", "", "");

                for (int i = 0; i < range.Length; i++)
                {
                    for (int j = 0; j < years.Length; j++)
                    {
                        if (range[i] == "Not Disclosed")
                        {
                            strCommand = "Select count(*) from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit = 'Non-Mill Sector' AND " + years[j] + query(i, years[j]);
                        }
                        else if (range[i] == "Sector Total")
                        {
                            strCommand = "Select count(*) from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit='Non-Mill Sector'";
                        }
                        else
                        {
                            strCommand = "Select count(" + years[j] + ") from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit = 'Non-Mill Sector' AND " + years[j] + query(i, years[j]);
                        }
                        dr = getData(conn, strCommand);
                        dr.Read();
                        values[j] = dr[0].ToString();
                        percent[j] = Math.Round((double.Parse(values[j]) / total) * 100, 2).ToString();
                        dr.Close();
                        dr.Dispose();
                    }
                    table.Rows.Add(range[i], values[0], percent[0], values[1], percent[1], values[2], percent[2]);                 
                }

                table.Rows.Add("Sector: Non Mill", "", "", "", "", "", "");

                for (int i = 0; i < range.Length; i++)
                {
                    for (int j = 0; j < years.Length; j++)
                    {
                        if (range[i] == "Not Disclosed")
                        {
                            strCommand = "Select count(*) from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit = 'Non-Mill Sector' AND " + years[j] + query(i, years[j]);
                        }
                        else if (range[i] == "Sector Total")
                        {
                            strCommand = "Select count(*) from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit='Non-Mill Sector'";
                        }
                        else
                        {
                            strCommand = "Select count(" + years[j] + ") from FinancialStatusoftheUnit,GeneralDetails where FinancialStatusoftheUnit.QreID = GeneralDetails.QreID AND TypeofUnit = 'Non-Mill Sector' AND " + years[j] + query(i, years[j]);
                        }
                        dr = getData(conn, strCommand);
                        dr.Read();
                        values[j] = dr[0].ToString();
                        percent[j] = Math.Round((double.Parse(values[j]) / total) * 100, 2).ToString();
                        dr.Close();
                        dr.Dispose();
                    }
                    table.Rows.Add(range[i], values[0], percent[0], values[1], percent[1], values[2], percent[2]);
                }
                dr.Close();
                dr.Dispose();
                
                wordHandling.openDocument();
                wordHandling.writetoWord(table, "DISTRIBUTION OF ANNUAL TURNOVER OF JDP UNITS DURING LAST 3 YEARS", "Table 11", 0, 1);
                wordHandling.CloseandSave();
            }
        }
    }
}   
