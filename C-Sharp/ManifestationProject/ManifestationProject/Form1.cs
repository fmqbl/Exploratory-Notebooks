using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework;
using IBM.Data.DB2;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;

namespace ManifestationProject
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        public Form1()
        {
            InitializeComponent();
            metroGrid1.Hide();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string ConStr = "Database = BRDB;User ID=inetsoft;Password = etl5boxes; server = khi.khi.ei:50012; Max Pool Size=100;Persist security info =False;Pooling = True ";

            string Query = @"select 
                    sh.MASTERBILL_NO as ""Master Number"",
                    ori.PORT_CITY as  ""Origin"",
                    des.PORT_CITY as ""Destination"",
                    sh.QTY as ""Packages"",
                    sh.QTY_UNITS as ""Unit"",
                    Max(CASE WHEN tr.TRACE_TYPE = 'CO' Then tr.ALL_TRACE_NO End) as ""Containers"",
                    Max(CASE WHEN tr.TRACE_TYPE = 'CO' Then tr.ALL_TRACE_QUANTITY END) as ""CO Size and QTY"",
                    sh.HOUSEBILL_NO as ""House Ref"",
                    MAX(CASE WHEN con.TRACE_TYPE = 'CO' ThEN con.MAX_TRACE_TRACE_NO END) AS ""Container Number"",
                    sh.CARRIER as ""Carrier"",
                    shipper.NAME as ""Shipper"",
                    Concat(shipper.ADDR_LINE1, shipper.ADDR_LINE2) as ""Shipper Address"",
                    shipper.NAME || ' ' || COALESCE(Concat(shipper.ADDR_LINE1, shipper.ADDR_LINE2), '') as ""Shipper Name and Address"",

                    ori.PORT_CITY as ""Shipper City"",
                    con.NAME as ""Consignee"",
                    Concat(con.ADDR_LINE1, con.ADDR_LINE2) as ""Consignee Address"",
                    con.NAME || ' ' || COALESCE(Concat(con.ADDR_LINE1, con.ADDR_LINE2), '') as ""Consignee Name and Address"",
                    des.PORT_CITY as ""Consignee City"",
                    sh.WGT as ""Weight"",
                    sh.DESC_OF_GOODS as ""Description"",
                    MAX(CASE WHEN con.TRACE_TYPE = 'SN' ThEN con.MAX_TRACE_TRACE_NO END) AS ""Seal Number""

                    from BRDB.IMPORT_SHIPMENT sh
                    inner
                    join BRDB.VW_IMPORT_EVENT_ETA ev on ev.FILE_NO = sh.FILE_NO
                    inner
                    join BRDB.CARRIERS cr on cr.CARRIER_CODE = sh.CARRIER
                    inner
                    join BRDB.CLIENTS shipper on shipper.CLIENT_NO = sh.SHIPPER_CLIENT_NO
                    inner
                    join BRDB.CLIENTS con on con.CLIENT_NO = sh.CONSIGNEE_CLIENT_NO
                    inner
                    join BRDB.IMPORT_TRACE_FLAT con on con.FILE_NO = sh.FILE_NO
                    inner
                    join BRDB.PORTS ori on ori.PORT_CODE = sh.ORIGIN
                    inner
                    join BRDB.PORTS des on des.PORT_CODE = sh.DESTINATION
                    inner
                    join brdb.IMPORT_TRACE_FLAT tr on tr.FILE_NO = sh.FILE_NO
                    where
                    days(ev.ETA_LA) > days(CURRENT TIMESTAMP) and days(ev.ETA_LA) <= days(CURRENT TIMESTAMP + 10 days)
                    and sh.FILE_TYPE = '7'
                    and sh.SHIPMENT_TYPE not in ('V', 'M')

                    GROUP by
                    sh.MASTERBILL_NO,
                    sh.ORIGIN,
                    sh.DESTINATION,
                    sh.QTY,
                    cr.CARRIER_NAME,
                    shipper.NAME,
                    sh.SHIPPER_CLIENT_NO,
                    sh.CARRIER,
                    sh.FILE_NO,
                    ev.ETA_LA,
                    sh.HOUSEBILL_NO,
                    shipper.CITY,
                    Concat(shipper.ADDR_LINE1, shipper.ADDR_LINE2),
                    shipper.NAME || ' ' || Concat(shipper.ADDR_LINE1, shipper.ADDR_LINE2),
                    con.NAME,
                    con.CITY,
                    sh.QTY_UNITS,
                    ori.PORT_CITY,
                    des.PORT_CITY,
                    Concat(con.ADDR_LINE1, con.ADDR_LINE2),
                    con.NAME || ' ' || con.ADDR_LINE1 || ' ' || con.ADDR_LINE2,
                    sh.WGT,
                    sh.DESC_OF_GOODS
                    ";


            using (DB2Connection myconnection = new DB2Connection(ConStr))
            {
                myconnection.Open();
                DB2Command cmd = myconnection.CreateCommand();
                cmd.CommandText = Query;
                //DB2DataReader rd = cmd.ExecuteReader();
                DB2DataAdapter dad = new DB2DataAdapter(cmd);
                DataTable dtRecord = new DataTable();
                dad.Fill(dtRecord);

                if (dtRecord != null)
                {
                    metroGrid1.DataSource = dtRecord;
                    metroGrid1.Show();
                }
              
            }
        }

        private void create_Email_Draft(String filename)
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                mailItem.Subject = "Test Subject";
                mailItem.To = "faisal.maqbool@expeditors.com";
                mailItem.Body = "Hi ! This is a test message";
                mailItem.Importance = Outlook.OlImportance.olImportanceLow;
                mailItem.Display(false);

                if (string.IsNullOrEmpty(filename) == false)
                {
                    // need to check to see if file exists before we attach !
                    if (!File.Exists(filename))
                        MessageBox.Show("Attached document " + filename + " does not exist", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        mailItem.Attachments.Add(filename, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                }

                mailItem.Close(Outlook.OlInspectorClose.olSave);

            }
            catch (Exception ex)
            {
                throw new Exception("Error occured trying to create email item --- " + Environment.NewLine + ex.Message);
            }
            


        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            //Iterating selected rows and creating manifest

            foreach (DataGridViewRow r in metroGrid1.SelectedRows)
            {
                if (Convert.ToString(r.Cells[9].Value) == "APLU")
                {
                    MetroFramework.MetroMessageBox.Show(this, Convert.ToString(r.Cells[10].Value));

                    Excel.Application app = new Excel.Application();
                    //string filepath = @"APLU.xls";
                    //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Excel.Workbook workbook = app.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + "/APLU.xls", ReadOnly: false, Editable: true);
                    Excel.Worksheet worksheet = workbook.Worksheets.Item[1];

                    //Update Values
                    worksheet.Range["B4"].Value = Convert.ToString(r.Cells[0].Value);
                    worksheet.Range["E4"].Value = Convert.ToString(r.Cells[7].Value);
                    worksheet.Range["B7"].Value = Convert.ToString(r.Cells[10].Value);

                    worksheet.Range["B9"].Value = Convert.ToString(r.Cells[1].Value);
                    worksheet.Range["B13"].Value = Convert.ToString(r.Cells[18].Value);
                    worksheet.Range["B20"].Value = Convert.ToString(r.Cells[15].Value);
                    worksheet.Range["B21"].Value = Convert.ToString(r.Cells[17].Value);

                    worksheet.Range["B24"].Value = Convert.ToString(r.Cells[3].Value);
                    worksheet.Range["B25"].Value = Convert.ToString(r.Cells[4].Value);

                    worksheet.Range["B28"].Value = Convert.ToString(r.Cells[19].Value);

                    //Containers Update

                    string containerStr = Convert.ToString(r.Cells[5].Value);

                    List<string> containers = containerStr.Split(',').ToList();

                    //containers = containers.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

                    //Container Pieces
                    string containerPieces = Convert.ToString(r.Cells[6].Value);

                    List<string> ctnPieces = containerPieces.Split(',').ToList();

                    //ctnPieces = ctnPieces.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

                    int counter = 31;

                    for (int i = 0; i < containers.Count; i++)
                    {
                        worksheet.Range["B" + counter.ToString()].Value = Convert.ToString(containers[i]);
                        worksheet.Range["D" + counter.ToString()].Value = Convert.ToString(ctnPieces[i].Split(' ').ToList()[0]);

                        counter = counter + 1;
                    }


                    //Save and Close
                    workbook.SaveAs(System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[0].Value), Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,Excel.XlSaveConflictResolution.xlUserResolution, true,Missing.Value, Missing.Value, Missing.Value);
                    app.Application.ActiveWorkbook.Save();
                    workbook.Close(0);
                    app.Application.Quit();
                    app.Quit();

                    create_Email_Draft(System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[0].Value) + ".xlsx");


                }
            }

        }
    }
}
