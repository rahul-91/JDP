using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class distributionChannels : reportHandling
    {
        public static void Report()
        {
            if (conn.State == ConnectionState.Open)
            {
                DataTable table = new DataTable();
                table.Columns.Add(" ", typeof(string));
                table.Columns.Add("Distribution Channel", typeof(string));
                table.Columns.Add("No.s", typeof(string));
                table.Columns.Add("Percentage (%)", typeof(double));

                OleDbDataReader dr;
                string[] channels = new string[] { "SDDOMDISTRIBUTORS", "SDDOMWHOLESALERS", "SDDOMOWNRETAILOUTLETS", 
                    "SDDOMDIRECTSELLING", "SDDOMNJBJMDCEXIBITIONFAIR", "SDDOMOTHERFAIRSSALESEMPSON", "","", 
                    "SDORSDIRECTEXP", "SDORSEXPTHUAGENFOREIGNIMP", "SDORSEXPTHRUOWNAGTFORCOUN", "SDORSINTERNATIONALTRADEMART","" };
                string[] channelName = new string[] { "Distributors", "Wholesalers", "Own Retail Outlets", 
                    "Direct Selling", "JMDC Exhibition", "Retail Sales through other e.g. Sales Emporia", 
                    "Other Exhibition/Fair", "Others", "Direct Export", "Export through Indian agent of Foreign Importer", 
                    "Export through own agent in Foreign Countries", "International Trade Mart","Others" };
                string sector = "";
                int total = 0;
                double percent;

                string strCommand = "Select count(*) from SalesDistribution";
                dr = getData(conn, strCommand);
                dr.Read();
                total = int.Parse(dr[0].ToString());

                for (int i = 0; i < channels.Length; i++)
                {
                    strCommand = "Select count(" + channels[i] + ") from SalesDistribution where " + channels[i] + " = 'Yes'";
                    switch (i)
                    {
                        case 0: sector = "Domestic";
                            break;
                        case 6: table.Rows.Add(sector, channelName[i],"0","0.00");
                            continue;
                        case 7: strCommand = "Select count(SDChannelsDomestic) from SalesDistribution where SDChannelsDomestic = 'Yes' AND SDDOMDISTRIBUTORS = 'No' AND SDDOMWHOLESALERS = 'No' AND SDDOMOWNRETAILOUTLETS = 'No' AND SDDOMDIRECTSELLING = 'No' AND SDDOMNJBJMDCEXIBITIONFAIR = 'No' AND SDDOMOTHERFAIRSSALESEMPSON = 'No'";
                            break;
                        case 8: sector = "Overseas";
                            break;
                        case 12: strCommand = "Select count(SDChannelsOverseas) from SalesDistribution where SDChannelsOverseas = 'Yes' AND SDORSDIRECTEXP = 'No' AND SDORSEXPTHUAGENFOREIGNIMP = 'No' AND SDORSEXPTHRUOWNAGTFORCOUN = 'No' AND SDORSEXPTHRUOWNAGTFORCOUN ='No' AND SDORSINTERNATIONALTRADEMART='No'";
                            break;
                    }
                    dr = getData(conn, strCommand);
                    dr.Read();
                    percent = Math.Round(double.Parse(dr[0].ToString()) / total * 100, 2);
                    table.Rows.Add(sector, channelName[i], dr[0].ToString(), percent);
                    sector = "";
                    dr.Close();
                    dr.Dispose();
                }
                dr.Close();
                dr.Dispose();
               
                wordHandling.openDocument();
                wordHandling.writetoWord(table, "DISTRIBUTION CHANNELS OF JDP UNITS : 2012-13", "Table 13", 0, 1);
                wordHandling.CloseandSave();
            }
        }
    }
}
