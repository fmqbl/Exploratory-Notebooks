using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using MetroFramework;
using Excel = Microsoft.Office.Interop.Excel;
using IBM.Data.DB2;

namespace GetHouseData
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {

        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //metroLabel1.Hide();
        }
        
        private void metroButton2_Click(object sender, EventArgs e)
        {
            string ConStr = "Database = " + metroTextBox3.Text.Trim() + ";User ID=inetsoft;Password = etl5boxes; server = "+ metroTextBox3.Text.Trim().ToLower()+"."+ metroTextBox3.Text.Trim().ToLower()+".ei:50000; Max Pool Size=100;Persist security info =False;Pooling = True ";
            string Query = @"Select 
                sh.SHPMNT_REF,
                ship.NAME,
                ship.ADDR_LINE1,
                ship.ADDR_LINE2,
                CONCAT(CONCAT(ship.CITY, ','), ship.COUNTRY),
                cons.NAME,
                cons.ADDR_LINE1,
                cons.ADDR_LINE2,
                CONCAT(CONCAT(cons.CITY, ','), cons.COUNTRY),
                hdr.ORIGIN_NAME_AWB,
                hdr.DESTIN_NAME_AWB,
                hdr.FLIGHT_TO_AWB_1,
                hdr.FLIGHT_BY_AWB_1,
                hdr.CHRG_CODE_FRGHT_1,
                hdr.CHRG_CODE_OTHER_1,
                sh.PCS_GRS,
                CAST(ROUND(sh.WGT_GRS_2 / 10000, 0) AS DECIMAL(18, 1)),
                CAST(ROUND(sh.WGT_CHRG_2 / 10000, 0) AS DECIMAL(18, 1)),
                hdr.FLIGHT_NUM_AWB_1,
                hdr.HNDL_INFO_AWB_1,
                hdr.HNDL_INFO_AWB_2,
                SUBSTR(xmlserialize(xmlagg(xmltext(CONCAT( '!',concat(concat(itm.REMARKS_AWB_PFX,' '),itm.REMARKS_AWB)))) as VARCHAR(1024)), 3),
                sl.MAWB_BL_NO
                                    

                From
                EXPORT.SHPMNT_HDR sh
                inner join IASDB.SHIP_LOG sl on sl.INVOICE_NO = sh.SHPMNT_REF
                inner join HELPDB.CLIENT ship on ship.CLIENT_NO = sh.SHIPPER_ID
                inner join HELPDB.CLIENT cons on cons.CLIENT_NO = sh.CONSIGN_ID
                inner join EXPORT.AIR_FRGHT_HDR hdr on hdr.INVOICE_REF = sh.SHPMNT_REF
                inner join EXPORT.AIR_FRGHT_ITEM itm on itm.CARRIER_REF = hdr.CARRIER_REF
                where sh.SHPMNT_REF = '" + metroTextBox1.Text.Trim() + @"'
                Group by
                ship.NAME,
                ship.ADDR_LINE1,
                ship.ADDR_LINE2,
                ship.CITY,
                ship.COUNTRY,
                sh.SHPMNT_REF,
                cons.NAME,
                cons.ADDR_LINE1,
                cons.ADDR_LINE2,
                cons.city,
                cons.COUNTRY,
                CAST(ROUND(sh.WGT_GRS_2 / 10000, 0) AS DECIMAL(18, 1)),
                CAST(ROUND(sh.WGT_CHRG_2 / 10000, 0) AS DECIMAL(18, 1)),
                hdr.ORIGIN_NAME_AWB,
                hdr.DESTIN_NAME_AWB,
                hdr.FLIGHT_TO_AWB_1,
                hdr.FLIGHT_BY_AWB_1,
                hdr.CHRG_CODE_FRGHT_1,
                hdr.CHRG_CODE_OTHER_1,
                hdr.FLIGHT_NUM_AWB_1,
                sh.PCS_GRS,
                hdr.HNDL_INFO_AWB_1,
                hdr.HNDL_INFO_AWB_2,
                sl.MAWB_BL_NO
                ";

            using (DB2Connection myconnection = new DB2Connection(ConStr))
            {
                myconnection.Open();
                DB2Command cmd = myconnection.CreateCommand();
                cmd.CommandText = Query;
                DB2DataReader rd = cmd.ExecuteReader();

                if (rd.HasRows)
                {
                    while (rd.Read())
                    {
                        //richTextBox1.Text = richTextBox1.Text + rd.GetValue(0).ToString() + "   " + rd.GetValue(1).ToString();


                        Excel.Application app = new Excel.Application();
                        //string filepath = @"HAWBTemp.xlsx";
                        //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                        Excel.Workbook workbook = app.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + "/HAWBTemplate.xlsx", ReadOnly: false, Editable: true);
                        Excel.Worksheet worksheet = workbook.Worksheets.Item[1];
                        //worksheet.Cells[5, 7].Value = rd.GetValue(1).ToString();
                        //Shipper Update
                        worksheet.Range["D1"].Value = rd.GetValue(22).ToString();
                        worksheet.Range["N1"].Value = rd.GetValue(0).ToString();
                        worksheet.Range["M51"].Value = rd.GetValue(0).ToString();
                        worksheet.Range["B3"].Value = rd.GetValue(1).ToString();
                        worksheet.Range["B4"].Value = rd.GetValue(2).ToString();
                        worksheet.Range["B5"].Value = rd.GetValue(3).ToString();
                        worksheet.Range["B6"].Value = rd.GetValue(4).ToString();

                        //Consignee Update

                        worksheet.Range["B8"].Value = rd.GetValue(5).ToString();
                        worksheet.Range["B9"].Value = rd.GetValue(6).ToString();
                        worksheet.Range["B10"].Value = rd.GetValue(7).ToString();
                        worksheet.Range["B11"].Value = rd.GetValue(8).ToString();

                        //Origin-Destination
                        worksheet.Range["B18"].Value = rd.GetValue(9).ToString();
                        worksheet.Range["B21"].Value = rd.GetValue(10).ToString();

                        //Dest-First Flight
                        worksheet.Range["B19"].Value = rd.GetValue(11).ToString();
                        worksheet.Range["C19"].Value = rd.GetValue(12).ToString();

                        //Prepaid-Collect

                        if (rd.GetValue(13).ToString() == "P")
                        {
                            worksheet.Range["K19"].Value = " P";
                        }
                        else
                        {
                            worksheet.Range["K19"].Value = "     C";
                        }

                        if (rd.GetValue(14).ToString() == "P")
                        {
                            worksheet.Range["L19"].Value = "P";
                        }
                        else
                        {
                            worksheet.Range["L19"].Value = "    C";
                        }
                        
                        //Weight-Pieces
                        worksheet.Range["B27"].Value = rd.GetValue(15).ToString();
                        worksheet.Range["C27"].Value = rd.GetValue(16).ToString();
                        worksheet.Range["G27"].Value = rd.GetValue(17).ToString();
                        worksheet.Range["B37"].Value = rd.GetValue(15).ToString();
                        worksheet.Range["C37"].Value = rd.GetValue(16).ToString();
                        //Flight
                        worksheet.Range["F21"].Value = rd.GetValue(18).ToString();

                        //Handling Information
                        worksheet.Range["C23"].Value = rd.GetValue(19).ToString();
                        worksheet.Range["C24"].Value = rd.GetValue(20).ToString();
                        worksheet.Range["O46"].Value = metroTextBox2.Text;

                        //Data
                        DateTime dt = DateTime.Now;
                        worksheet.Range["I49"].Value = dt.ToShortDateString();

                        //Description

                        String description = rd.GetValue(21).ToString();

                        List<string> descResult = description.Split('!').ToList();

                        descResult = descResult.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

                        //richTextBox1.Text = richTextBox1.Text + descResult[0] +"///////" + descResult[1]+ "//////" + descResult[2];

                        for (int i = descResult.Count; i <= 7; i++)
                        {
                            descResult.Add("");
                        }

                        worksheet.Range["B28"].Value = descResult[0];
                        worksheet.Range["B29"].Value = descResult[1];
                        worksheet.Range["B30"].Value = descResult[2];
                        worksheet.Range["B31"].Value = descResult[3];
                        worksheet.Range["B32"].Value = descResult[4];
                        worksheet.Range["B33"].Value = descResult[5];
                        worksheet.Range["B34"].Value = descResult[6];
                        worksheet.Range["B35"].Value = descResult[7];

                        

                        /*String remarks = rd.GetValue(22).ToString();
                        List<string> remarksResult = remarks.Split('!').ToList();
                        remarksResult = remarksResult.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                        for (int i = remarksResult.Count; i <= 7; i++)
                        {
                            remarksResult.Add("");
                        }

                        MetroMessageBox.Show(this,remarksResult[1]);
                        MetroMessageBox.Show(this, remarksResult[2]);
                        MetroMessageBox.Show(this, remarksResult[3]);
                        MetroMessageBox.Show(this, remarksResult[4]);
                        MetroMessageBox.Show(this, remarksResult[5]);

                        worksheet.Range["M28"].Value = remarksResult[0];
                        worksheet.Range["M29"].Value = remarksResult[1];
                        worksheet.Range["M30"].Value = remarksResult[2];
                        worksheet.Range["M31"].Value = remarksResult[3];
                        worksheet.Range["M32"].Value = remarksResult[4];
                        worksheet.Range["M33"].Value = remarksResult[5];
                        worksheet.Range["M34"].Value = remarksResult[6];
                        worksheet.Range["M35"].Value = remarksResult[7];*/

                        workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,System.IO.Directory.GetCurrentDirectory() + "/"+ metroTextBox1.Text.Trim() +".pdf");
                        app.Application.ActiveWorkbook.Save();
                        workbook.Close(0);
                        app.Application.Quit();
                        app.Quit();

                        MetroFramework.MetroMessageBox.Show(this, "House Created!");
                    }

                }

            }
        }

    }   
}
