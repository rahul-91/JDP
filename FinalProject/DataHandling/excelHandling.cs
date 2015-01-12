using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.OleDb;

namespace FinalProject
{
    class excelHandling
    {
        public static void readExcelfile(string filepath)
        {
            string[] worksheets = new string[4] { "GeneralDetails", "FinancialStatusoftheUnit", "SalesDistribution", "ManpowerEmployment"};
            string[] query = new string[4] {"CompanyStatus,TypeofUnit,TypeofProducer ",
                "AnnTurnover201011OnJDP,AnnTurnover201112OnJDP,AnnTurnover201213POnJDP ",
                "SalesDistChannelsDomestic,SalesDistDomesticDistributors,SalesDistDomesticWholesalers,SalesDistDomesticOwnRetailOutlets,SalesDistDomesticDirectSelling,SalesDistDomesticNJBJMDCExibitionFair,SalesDistDomesticOtherFairsSalesEmporiaSonali,SalesDistChannelsOverseas,SalesDistOverseasDirectExp,SalesDistOverseasExpthourghAgenForeignImp,SalesDistOverseasExpthroughownagentForeignCoun,SalesDistOverseasInternationalTradeMart",
                "DoyouconsumepowerprodJDPs,ConPowerProdJDPsIfYesConnloadProdJDPsinKW"};
            //System.DateTime exceltime = System.DateTime.Now;
            try
            {
                OleDbConnection con = new OleDbConnection(ConfigurationManager.ConnectionStrings["excelConnection"].ConnectionString + "Data Source=" + filepath);
               
                con.Open();
                DataTable[] excelTable = new DataTable[4];
               
                for (int i = 0; i < worksheets.Length; i++)
                {
                    OleDbCommand com = new OleDbCommand("Select QreID," + query[i] + " from [" + worksheets[i] + "$] where QreID is not null", con);
                    OleDbDataAdapter da = new OleDbDataAdapter(com);
                    
                    excelTable[i] = new DataTable();
                    da.Fill(excelTable[i]);
                    da.FillSchema(excelTable[i], SchemaType.Source);
                }
                //System.Windows.Forms.MessageBox.Show(System.DateTime.Now.Subtract(exceltime).ToString());
                dbHandling.writeDataBase(worksheets, excelTable);
                con.Close();               
            }
            catch (Exception ex)
            {
              System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}
