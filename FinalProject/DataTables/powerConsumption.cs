using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class powerConsumption : reportHandling
    {
        public static void modifyManpowerEmploymentTable()
        {
            string[] sqlCommand = new string[]{"Update ManpowerEmployment SET CNPWRPRDJDPSINKW = '0.745' where CNPWRPRDJDPSINKW = '1HP'",
                "Update ManpowerEmployment SET CNPWRPRDJDPSINKW = '0.5' where CNPWRPRDJDPSINKW = '1/2'",
                "Update ManpowerEmployment SET DOYOUCONSUMEPOWERPRODJDPS = 'Yes' where DOYOUCONSUMEPOWERPRODJDPS = 'YES'",
                "Update ManpowerEmployment SET CNPWRPRDJDPSINKW = '0' where CNPWRPRDJDPSINKW  = 'CANNOT SAY' OR CNPWRPRDJDPSINKW = 'CANN''T SAY' OR CNPWRPRDJDPSINKW is null OR DOYOUCONSUMEPOWERPRODJDPS ='No'",
                "Alter table ManpowerEmployment add (CNPWRPRDJDPSINKWTemp number(8,3))",
                "Update ManpowerEmployment set CNPWRPRDJDPSINKWTemp = to_number(CNPWRPRDJDPSINKW)" };
            
            foreach( string sqlComm in sqlCommand)
            {
                OleDbCommand cmd = new OleDbCommand(sqlComm, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                modifyManpowerEmploymentTable();
                table.Columns.Add("Particulars ", typeof(string));
                table.Columns.Add("No.s", typeof(string));
                table.Columns.Add("Percentage (%)", typeof(string));
                table.Rows.Add("Mill Sector","", "");
                //table.Rows.Add("Consumer Power", "", "");
                

                OleDbDataReader dr;
                string[] range = new string[] { "<= 5 AND CNPWRPRDJDPSINKWTemp <> 0","> 5 AND CNPWRPRDJDPSINKWTemp <= 10", " > 10 AND CNPWRPRDJDPSINKWTemp <= 25",      
                    " > 25 AND CNPWRPRDJDPSINKWTemp <= 50"," > 50 AND CNPWRPRDJDPSINKWTemp <= 100"," > 100 AND CNPWRPRDJDPSINKWTemp <= 250",
                    " > 250 AND CNPWRPRDJDPSINKWTemp <= 500"," > 500 AND CNPWRPRDJDPSINKWTemp <= 750"," > 750 AND CNPWRPRDJDPSINKWTemp <= 1000",
                    " > 1000 AND CNPWRPRDJDPSINKWTemp <= 2000",">= 2000"," = 0"};

                string[] cellname = new string[] { "Upto 5 KW", "5 KW to 10 KW", "10 KW to 25 KW", "25 KW to 50 KW",
                    "50 KW to 100 KW","100 KW to 250 KW","250 KW to 500 KW","500 KW to 750 KW","750 KW to 1000 KW",
                    "1000 KW to 2000 KW", "Above 2000 KW","Not Disclosed"};

                string[] consumePower = new string[] {"No","Yes"};
                string[] cellConsumePowerName = new string[]{"Do not Consume Power","Consume Power"};

                string strCommand = "Select count(*) from ManpowerEmployment";
                int total;
                string percent;

                dr = getData(conn, strCommand);
                dr.Read();
                total = int.Parse(dr[0].ToString());

                for (int i = 0; i < consumePower.Length; i++)
                {
                    strCommand = "Select count(DOYOUCONSUMEPOWERPRODJDPS) from ManpowerEmployment,GeneralDetails where GeneralDetails.QreID = ManpowerEmployment.QreID AND TypeofUnit = 'Jute Mill' AND DOYOUCONSUMEPOWERPRODJDPS ='" + consumePower[i] + "'";
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round((double.Parse(dr[0].ToString()) / total) * 100, 2).ToString();
                    table.Rows.Add(cellConsumePowerName[i], dr[0].ToString(), percent);
                    dr.Close();
                    dr.Dispose();
                }
                
                for (int i = 0; i < range.Length; i++)
                {
                    strCommand = "Select count(TypeofUnit) from ManpowerEmployment,GeneralDetails where GeneralDetails.QreID = ManpowerEmployment.QreID AND TypeofUnit = 'Jute Mill' AND DOYOUCONSUMEPOWERPRODJDPS = 'Yes' AND (CNPWRPRDJDPSINKWTemp " + range[i] + ")";
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round((double.Parse(dr[0].ToString()) / total) * 100, 2).ToString();
                    table.Rows.Add(cellname[i], dr[0].ToString(), percent);
                    dr.Close();
                    dr.Dispose();
                }

                table.Rows.Add("Non Mill Sector", "", "");

                string[] query = new string[] { "", "AND TypeofProducer LIKE 'Producer%'", "AND TypeofProducer LIKE 'Merchant%'" };
                string[] cellnameQuery = new string[] { "Total Non-mill units", "JDP Manufacturing Units","Merchant Exporters" };
                for (int i = 0; i < query.Length; i++)
                {
                    strCommand = "Select count(*) from ManpowerEmployment,GeneralDetails where GeneralDetails.QreID = ManpowerEmployment.QreID AND TypeofUnit = 'Non-Mill Sector'" + query[i];
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round((double.Parse(dr[0].ToString()) / total) * 100, 2).ToString();
                    table.Rows.Add(cellnameQuery[i], dr[0].ToString(), percent);
                    dr.Close();
                    dr.Dispose();
                }
              
                for (int i = 0; i < consumePower.Length; i++)
                {
                    strCommand = "Select count(DOYOUCONSUMEPOWERPRODJDPS) from ManpowerEmployment,GeneralDetails where GeneralDetails.QreID = ManpowerEmployment.QreID AND TypeofUnit = 'Non-Mill Sector' AND DOYOUCONSUMEPOWERPRODJDPS ='" + consumePower[i]+"'" ;
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round((double.Parse(dr[0].ToString()) / total) * 100, 2).ToString();
                    table.Rows.Add(cellConsumePowerName[i], dr[0].ToString(), percent);
                    dr.Close();
                    dr.Dispose();
                }
                //table.Rows.Add("Consumer Power", "", "");
                for (int i = 0; i < range.Length; i++)
                {
                    strCommand = "Select count(*) from ManpowerEmployment,GeneralDetails where GeneralDetails.QreID = ManpowerEmployment.QreID AND TypeofUnit = 'Non-Mill Sector' AND DOYOUCONSUMEPOWERPRODJDPS = 'Yes' AND (CNPWRPRDJDPSINKWTemp " + range[i] + ")";
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round((double.Parse(dr[0].ToString()) / total) * 100, 2).ToString();
                    table.Rows.Add(cellname[i], dr[0].ToString(), percent);
                    dr.Close();
                    dr.Dispose();
                }
                dr.Close();
                dr.Dispose();

                string sqlCommand = "Alter table ManpowerEmployment drop column CNPWRPRDJDPSINKWTemp";
                OleDbCommand cmd = new OleDbCommand(sqlCommand, conn);
                cmd.ExecuteNonQuery();

                wordHandling.openDocument();
                wordHandling.writetoWord(table, "POWER CONSUMPTION SCENARIO OF SURVEYED UNITS FOR PRODUCTION OF JDPs: 2012-13", "Table 16", 0, 1);
                wordHandling.CloseandSave();
            }
        }
    }
}
       