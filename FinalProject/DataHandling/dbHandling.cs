using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Configuration;
using System.Data.OleDb;

namespace FinalProject
{
    class dbHandling
    {
        public static void writeDataBase(string[] worksheets,DataTable[] excelTable)
        {
            try
            {
                OleDbConnection conn = new OleDbConnection(ConfigurationManager.ConnectionStrings["oracleConnection"].ConnectionString);
                conn.Open();
                //System.DateTime exceltime = System.DateTime.Now;
                for (int i = 0; i < worksheets.Length; i++)
                {
                    OleDbCommand trunc = new OleDbCommand("Truncate table " + worksheets[i], conn);
                    trunc.ExecuteNonQuery();
                    
                    //DataTable oraTable = new DataTable();
                    //oraTable.Merge(excelTable[i]);
                    //oraTable.AcceptChanges();

                    OleDbDataAdapter oda = new OleDbDataAdapter();
                    oda.SelectCommand = new OleDbCommand("Select * from " + worksheets[i], conn);
                    OleDbCommandBuilder ocb = new OleDbCommandBuilder(oda);

                    //DataSet ods = new DataSet();
                    DataTable oraTable = new DataTable();
                    oda.Fill(oraTable);

                    //oraTable = excelTable[i].Copy();
                    //oraTable.Merge(excelTable[i],true);
                    //oraTable.AcceptChanges();
                    //Form2 fm = new Form2();
                    //fm.dataGridView1.DataSource = oraTable ;
                    //fm.Show();
                    
                    //System.Windows.Forms.MessageBox.Show(oraTable.Rows.Count.ToString());

                    foreach (DataRow drow in excelTable[i].Rows)
                    {
                        oraTable.Rows.Add(drow.ItemArray);
                        //oraTable.ImportRow(drow);
                        //oraTable.AcceptChanges();
                    }

                    oda.Update(oraTable);
                }
                //System.Windows.Forms.MessageBox.Show(System.DateTime.Now.Subtract(exceltime).ToString());
                conn.Close();
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}
