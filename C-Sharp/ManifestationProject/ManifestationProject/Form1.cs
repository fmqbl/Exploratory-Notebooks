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
using System.Data.OleDb;
using System;
using System.IO;
using System.Net;

namespace ManifestationProject
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        string ConStr = "Database = BRDB;User ID=inetsoft;Password = etl5boxes; server = khi.khi.ei:50012; Max Pool Size=100;Persist security info =False;Pooling = True ";

        string folder = @"/images/storage/";
        //string remoteFile = folder + FILE_NAME_IMAGE;
        string downLocation = System.IO.Directory.GetCurrentDirectory();


        public Form1()
        {
            InitializeComponent();
            metroGrid1.Hide();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }


        private void extract_document_locations(String fileNumber, String consolNumber, String hbNumber, String masterNumber)
        {
            String ConString = "Database = edoc; User ID = inetsoft; Password = etl5boxes; server = edoc.khi.ei:50002; Max Pool Size = 100; Persist security info = False; Pooling = True ";
            String masterFileType = "";
            String houseFileType = "";
            string docQuery = @"SELECT DISTINCT

                                EDOC.FOLDER.KEY_,
                                EDOC.FOLDER.DESCRIPTION,
                                EDOC.DOCUMENT.DESCRIPTION,
                                EDOC.DOCUMENT.CREATION_TIME,
                                EDOC.DOCUMENT.FILE_NAME_IMAGE,
                                EDOC.FOLDER.KEY_TYPE,
                                EDOC.DOCUMENT.DOC_TYPE,
                                EDOC.DOCUMENT.IMAGE_FILE_TYPE,
                                EDOC.DOCUMENT.CREATED_BY_NAME

                                from(EDOC.FOLDER INNER JOIN EDOC.FOLDER_DOCUMENT ON EDOC.FOLDER.G_U_I_D = EDOC.FOLDER_DOCUMENT.FOLDER__P_K)
                                LEFT OUTER JOIN EDOC.DOCUMENT ON EDOC.DOCUMENT.G_U_I_D = EDOC.FOLDER_DOCUMENT.DOCUMENT__P_K and EDOC.DOCUMENT.DOC_TYPE IN('HOU', 'HOR', 'MOR', 'MOB')
                                where EDOC.FOLDER.KEY_ IN('" + fileNumber + "','" + consolNumber + "','" + hbNumber + "')AND EDOC.DOCUMENT.FILE_NAME_IMAGE IS NOT NULL";

            //List<String> columnData = new List<String>();
            //Connecting to DB
            using (DB2Connection myconnection = new DB2Connection(ConString))
            {

                myconnection.Open();
                DB2Command cmd = myconnection.CreateCommand();
                cmd.CommandText = docQuery;
                //DB2DataReader rd = cmd.ExecuteReader();
                using (DB2DataReader reader = cmd.ExecuteReader())
                {
                    string host = "ftp://edoc.khi.ei//";
                    string user = "mpesup";
                    string pass = "mpesup";
                    bool checkHouse = false;
                    bool checkMaster = false;
                    

                    while (reader.Read())
                    {
                        string rf = folder + reader.GetString(4);
                        

                        //MetroFramework.MetroMessageBox.Show(this,reader.GetString(6));

                        string lf = "";
                        if ((reader.GetString(6) == "HOU") && ((reader.GetString(7) == "TIFF") && (!checkHouse)))
                        {
                            lf = downLocation + "/" + hbNumber + ".tiff";
                            download(rf, lf, host, user, pass);
                            checkHouse = true;
                            houseFileType = ".tiff";
                        }
                        if ((reader.GetString(6) == "HOR") && ((reader.GetString(7) == "TIFF") && (!checkHouse)))
                        {
                            lf = downLocation + "/" + hbNumber + ".tiff";
                            download(rf, lf, host, user, pass);
                            checkHouse = true;
                            houseFileType = ".tiff";
                        }
                        if ((reader.GetString(6) == "HOU") && ((reader.GetString(7) == "PDF") && (!checkHouse)))
                        {
                            lf = downLocation + "/" + hbNumber + ".pdf";
                            download(rf, lf, host, user, pass);
                            checkHouse = true;
                            houseFileType = ".pdf";
                        }
                        if ((reader.GetString(6) == "HOR") && ((reader.GetString(7) == "PDF") && (!checkHouse)))
                        {
                            lf = downLocation + "/" + hbNumber + ".pdf";
                            download(rf, lf, host, user, pass);
                            checkHouse = true;
                            houseFileType = ".pdf";
                        }

                        if ((reader.GetString(6) == "MOR") && ((reader.GetString(7) == "PDF") && (!checkMaster)))
                        {
                            lf = downLocation + "/" + masterNumber + ".pdf";
                            download(rf, lf, host, user, pass);
                            checkMaster = true;
                            masterFileType = ".pdf";
                        }
                        if ((reader.GetString(6) == "MOR") && ((reader.GetString(7) == "TIFF") && (!checkMaster)))
                        {
                            lf = downLocation + "/" + masterNumber + ".tiff";
                            download(rf, lf, host, user, pass);
                            checkMaster = true;
                            masterFileType = ".tiff";
                        }
                        if ((reader.GetString(6) == "MOB") && ((reader.GetString(7) == "PDF") && (!checkMaster)))
                        {
                            lf = downLocation + "/" + masterNumber + ".pdf";
                            download(rf, lf, host, user, pass);
                            checkMaster = true;
                            masterFileType = ".pdf";
                        }
                        if ((reader.GetString(6) == "MOB") && ((reader.GetString(7) == "TIFF") && (!checkMaster)))
                        {
                            lf = downLocation + "/" + masterNumber + ".tiff";
                            download(rf, lf, host, user, pass);
                            checkMaster = true;
                            masterFileType = ".tiff";
                        }
                    }
                }

            }
            create_Email_Draft(downLocation + "/" + masterNumber + ".xlsx", downLocation + "/" + masterNumber + masterFileType,  downLocation + "/" + hbNumber + houseFileType);
            //Show


        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            

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

                    --Shipper
                    shipper.NAME as ""Shipper"",
                    shipper.ADDR_LINE1 as ""Shipper Address"",
                    --Concat(shipper.NAME, shipper.CITY) as ""Shipper Name and Address"",

                    shipper.NAME || ' ' || shipper.ADDR_LINE1 as ""Shipper Name and Address"",

                    ori.PORT_CITY as ""Shipper City"",

                    con.NAME as ""Consignee"",
                    con.ADDR_LINE1 as ""Consignee Address"",
                    con.NAME || ' ' || con.ADDR_LINE1 as ""Consignee Name and Address"",

                    des.PORT_CITY as ""Consignee City"",
                    sh.WGT as ""Weight"",
                    sh.DESC_OF_GOODS as ""Description"",
                    MAX(CASE WHEN con.TRACE_TYPE = 'SN' ThEN con.MAX_TRACE_TRACE_NO END) AS ""Seal Number"",
                    sh.FILE_NO,
                    sh.CONSOL_NO

                    from BRDB.IMPORT_SHIPMENT sh
                    inner join BRDB.VW_IMPORT_EVENT_ETA ev on ev.FILE_NO = sh.FILE_NO
                    inner join BRDB.CARRIERS cr on cr.CARRIER_CODE = sh.CARRIER
                    inner join BRDB.CLIENTS shipper on shipper.CLIENT_NO = sh.SHIPPER_CLIENT_NO
                    inner join BRDB.CLIENTS con on con.CLIENT_NO = sh.CONSIGNEE_CLIENT_NO
                    inner join BRDB.IMPORT_TRACE_FLAT con on con.FILE_NO = sh.FILE_NO
                    inner join BRDB.PORTS ori on ori.PORT_CODE = sh.ORIGIN
                    inner join BRDB.PORTS des on des.PORT_CODE = sh.DESTINATION
                    inner join brdb.IMPORT_TRACE_FLAT tr on tr.FILE_NO = sh.FILE_NO
                    inner join BRDB.EDOC_FOLDER EDOC on EDOC.FOLDER = SH.FILE_NO
                    where
                    days(ev.ETA_LA) > days(CURRENT TIMESTAMP) and days(ev.ETA_LA) <= days(CURRENT TIMESTAMP + 10 days)
                    and sh.FILE_TYPE = '7'
                    and sh.SHIPMENT_TYPE not in ('V', 'M')


                    GROUP by
                    sh.FILE_NO,
                    sh.CONSOL_NO,
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
                    shipper.NAME || ' ' || shipper.ADDR_LINE1,
                    shipper.ADDR_LINE1,
                    con.NAME,
                    con.CITY,
                    sh.QTY_UNITS,
                    ori.PORT_CITY,
                    des.PORT_CITY,
                    Concat(con.ADDR_LINE1, con.ADDR_LINE2),
                    con.NAME || ' ' || con.ADDR_LINE1 || ' ' || con.ADDR_LINE2,
                    con.ADDR_LINE1,
                    con.ADDR_LINE2,
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

        private void create_Email_Draft(String filename, String master, String house)
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
                        mailItem.Attachments.Add(master, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                        mailItem.Attachments.Add(house, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                }

                mailItem.Close(Outlook.OlInspectorClose.olSave);

            }
            catch (Exception ex)
            {
                throw new Exception("Error occured trying to create email item --- " + Environment.NewLine + ex.Message);
            }
        }

        public static void download(string remoteFile, string localFile, string hostIP, string userName, string password)
        {
            try
            {
                /* Create an FTP Request */
                FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create(hostIP + remoteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(userName, password);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                /* Establish Return Communication with the FTP Server */
                FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Get the FTP Server's Response Stream */
                Stream ftpStream = ftpResponse.GetResponseStream();
                /* Open a File Stream to Write the Downloaded File */
                FileStream localFileStream = new FileStream(localFile, FileMode.Create);
                /* Buffer for the Downloaded Data */
                //int bufferSize = 10240;
                //byte[] byteBuffer = new byte[bufferSize];
                //int bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                /* Download the File by Writing the Buffered Data Until the Transfer is Complete */
                try
                {
                    //while (bytesRead > 0)
                    //{
                    //    localFileStream.Write(byteBuffer, 0, bytesRead);
                    //    bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                    //}

                    byte[] buffer = new byte[10240];
                    int read;
                    while ((read = ftpStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        localFileStream.Write(buffer, 0, read);
                        Console.WriteLine("Downloaded {0} bytes", localFileStream.Position);
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.InnerException.ToString());
                }
                /* Resource Cleanup */

                //MessageBox.Show("Download done for " + localFile);
                localFileStream.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.InnerException.ToString());
            }
            return;
        }

        

        private void metroButton2_Click(object sender, EventArgs e)
        {
            //Iterating selected rows and creating manifest

            foreach (DataGridViewRow r in metroGrid1.SelectedRows)
            {
                if (Convert.ToString(r.Cells[9].Value) == "APLU")
                {
                    //MetroFramework.MetroMessageBox.Show(this, Convert.ToString(r.Cells[10].Value));

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

                    extract_document_locations(Convert.ToString(r.Cells[21].Value), Convert.ToString(r.Cells[22].Value), Convert.ToString(r.Cells[7].Value), Convert.ToString(r.Cells[0].Value));
                    //create_Email_Draft(System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[0].Value) + ".xlsx", System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[0].Value) + ".pdf", System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[7].Value) + ".tiff");
                }
                if (Convert.ToString(r.Cells[9].Value) == "APLU")
                {
                    //MetroFramework.MetroMessageBox.Show(this, Convert.ToString(r.Cells[10].Value));

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
                    workbook.SaveAs(System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[0].Value), Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
                    app.Application.ActiveWorkbook.Save();
                    workbook.Close(0);
                    app.Application.Quit();
                    app.Quit();

                    extract_document_locations(Convert.ToString(r.Cells[21].Value), Convert.ToString(r.Cells[22].Value), Convert.ToString(r.Cells[7].Value), Convert.ToString(r.Cells[0].Value));
                    //create_Email_Draft(System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[0].Value) + ".xlsx", System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[0].Value) + ".pdf", System.IO.Directory.GetCurrentDirectory() + "/" + Convert.ToString(r.Cells[7].Value) + ".tiff");
                }
            }
            MetroMessageBox.Show(this, "Manifest is Done");

        }
    }
}
