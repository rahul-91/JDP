using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class state_ownerJDPunitsDistribution : reportHandling
    {
        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                table.Columns.Add("State", typeof(string));
                table.Columns.Add("Proprietorship", typeof(string));
                table.Columns.Add("Partnership", typeof(string));
                table.Columns.Add("Public/ Pvt. Ltd.", typeof(string));
                table.Columns.Add("Co-op Society", typeof(string));
                table.Columns.Add("Society/ Voluntary Organisation", typeof(string));
                table.Columns.Add("Others", typeof(string));
                table.Columns.Add("Total", typeof(string));

                OleDbDataReader dr;
                string strCommand = "",total = ""; 
                string[] companyStatus = new string[6] { "Proprietorship", "Partnership",
                    "Public/Private Limited", "Co-operative Society","Voluntary Organisation/NGO","Any other"};
                string[] type_count = new string[6];
                
                for(int i = 0 ; i < stateCode.Length ; i++)
                {
                    for(int j = 0; j < 6 ; j++)
                    {
                        strCommand = "Select count(CompanyStatus) from GeneralDetails where CompanyStatus = '" + companyStatus[j] + "' AND QreID LIKE '" + stateCode[i] + "%'";
                        dr = getData(conn, strCommand);
                        dr.Read();
                        type_count[j] = dr[0].ToString();
                        dr.Close();
                        dr.Dispose();
                    }
                    strCommand = "Select count(*) from GeneralDetails where QreID LIKE '"+ stateCode[i] +"%'";
                    dr = getData(conn, strCommand);
                    dr.Read();
                    total = dr[0].ToString();
                    dr.Close();
                    dr.Dispose();

                    table.Rows.Add(states[i], type_count[0], type_count[1], type_count[2], type_count[3], type_count[4], type_count[5], total);
                }
                for (int i = 0; i < 6; i++)
                {
                    strCommand = "Select count(CompanyStatus) from GeneralDetails where CompanyStatus = '" + companyStatus[i] + "'";
                    dr = getData(conn, strCommand);
                    dr.Read();
                    type_count[i] = dr[0].ToString();
                    dr.Close();
                    dr.Dispose();
                }
                strCommand = "Select count(*) from GeneralDetails";
                dr = getData(conn, strCommand);
                dr.Read();
                total = dr[0].ToString();
                table.Rows.Add("All Total", type_count[0], type_count[1], type_count[2], type_count[3], type_count[4], type_count[5], total);
                dr.Close();
                dr.Dispose();
                wordHandling.openDocument();
                wordHandling.writetoWord(table, "STATE-WISE & OWNERSHIP-WISE DISTRIBUTION OF JDP UNITS (As on 31st March '13)", "Table 7", 0, 1);
                wordHandling.CloseandSave();
            }
        }
    }
}
