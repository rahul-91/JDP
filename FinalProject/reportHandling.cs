using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;


namespace FinalProject
{
    class reportHandling
    {
        protected static string[] states = new string[]{"A&N Islands","Andhra Pradesh","Assam","Bihar",
        "Chhattishgarh","Delhi","Goa","Gujarat","Haryana","Jharkhand","Karnataka","Kerala","Madhya Pradesh",
        "Maharashtra","Meghalaya","Nagaland","Orissa","Pondicherry","Rajasthan","Sikkim","Tamil Nadu",
        "Uttar Pradesh","West Bengal"};

        protected static string[] stateCode = new string[] {"AMN","ANP","ASM","BHR","CTG","DEL","GOA","GUJ",
        "HRY","JHR","KAR","KRL","MPR","MHR","MGh","NAG","ORI","PND","RAJ","SKM","TNU","UPR","WEB"};

        protected static OleDbConnection conn = new OleDbConnection(System.Configuration.ConfigurationManager.ConnectionStrings["oracleConnection"].ConnectionString);

        protected static OleDbDataReader getData(OleDbConnection conn, String strCommand)
        {
            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader dr;
            cmd.Connection = conn;
            cmd.CommandText = strCommand;
            cmd.CommandType = CommandType.Text;
            dr = cmd.ExecuteReader();
            return dr;
        }
        
        public static void generateReport(string index)
        {
            //try
            {
                conn.Open();
                for (int i = 0; i < index.Length; i++)
                {
                    switch (index.Substring(i,1))
                    {
                        case "0": sectorDistribution.Report();
                                  categoryDistribution.Report();
                                  ownershipStatus.Report();
                                  state_ownerJDPunitsDistribution.Report();
                                  annualTurnover.Report();
                                  distributionChannels.Report();
                                  powerConsumption.Report();
                                  old_newJDPunitsDistribution.Report();
                            break;
                        case "1": sectorDistribution.Report();
                            break;
                        case "2": categoryDistribution.Report();
                            break;
                        case "3": ownershipStatus.Report();
                            break;
                        case "4": state_ownerJDPunitsDistribution.Report();
                            break;
                        case "5": annualTurnover.Report();
                            break;
                        case "6": distributionChannels.Report();
                            break;
                        case "7": powerConsumption.Report();
                            break;
                        case "8": old_newJDPunitsDistribution.Report();
                            break;
                        default: System.Windows.Forms.MessageBox.Show("Index Error!", "ERROR", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            break;
                    }
                }
                conn.Close();
            }
            //catch (Exception ex)
            {
              // System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}
