using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using GemBox.Spreadsheet;
using System.Globalization;

namespace Criteria_Validator
{
    public partial class CriteriaValidator : Form
    {

        List<string> ElementsIds = new List<string>();
        List<string> Areas = new List<string>();
        List<string> Volumes = new List<string>();
        List<double> TotalDurations = new List<double>();
        List<DateTime> DateOfOrderings = new List<DateTime>();

        static string smtpAddress = "smtp.gmail.com";
        static int portNumber = 587;
        static bool enableSSL = true;
        static string emailFromAddress = "navisvalidator@gmail.com"; //Sender Email Address  
        static string password = "NavisValidator100%"; //Sender Password  
        static string subject = "Ping";
        static string body = "Hello, This Email was sent as a ping.";

        string EmailList_FileName = "EmailList.csv";
        //--------------------------------------------------------------------Inputs
        string unit;
        string material;
        double TQ;
        double CS;
        double LT;
        double ALT;
        double ADC;
        double MLT;
        double MDC;
        double ROP;
        double DC;
        //--------------------------------------------------------------------Outputs
        double RSDEF;
        double SSDEF;

        public CriteriaValidator()
        {
            InitializeComponent();


            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workbook1 = ExcelFile.Load("NAVIS OUT.csv");

            var dataTable1 = new DataTable();

            // Depending on the format of the input file, you need to change this:
            dataTable1.Columns.Add("Material", typeof(string));
            dataTable1.Columns.Add("GUID", typeof(string));
            dataTable1.Columns.Add("ELEMENT ID", typeof(string));
            dataTable1.Columns.Add("NAME", typeof(string));
            dataTable1.Columns.Add("GLOABALID", typeof(string));
            dataTable1.Columns.Add("START DATE", typeof(string));
            dataTable1.Columns.Add("END DATE", typeof(string));
            dataTable1.Columns.Add("DURATION", typeof(string));

            // Select the first worksheet from the file.
            var worksheet1 = workbook1.Worksheets[0];

            var options = new ExtractToDataTableOptions(0, 0, 10000);
            options.ExtractDataOptions = ExtractDataOptions.StopAtFirstEmptyRow;
            options.ExcelCellToDataTableCellConverting += (sender, e) =>
            {
                if (!e.IsDataTableValueValid)
                {
                    // GemBox.Spreadsheet doesn't automatically convert numbers to strings in ExtractToDataTable() method because of culture issues; 
                    // someone would expect the number 12.4 as "12.4" and someone else as "12,4".
                    e.DataTableValue = e.ExcelCell.Value == null ? null : e.ExcelCell.Value.ToString();
                    e.Action = ExtractDataEventAction.Continue;
                }
            };

            // Extract the data from an Excel worksheet to the DataTable.
            // Data is extracted starting at first row and first column for 10 rows or until the first empty row appears.
            worksheet1.ExtractToDataTable(dataTable1, options);


            //-----------------Again

            var workbook2 = ExcelFile.Load("REVIT output.xlsx");

            var dataTable2 = new DataTable();

            // Depending on the format of the input file, you need to change this:
            dataTable2.Columns.Add("Count", typeof(string));
            dataTable2.Columns.Add("ELEMENT ID", typeof(string));
            dataTable2.Columns.Add("Description", typeof(string));
            dataTable2.Columns.Add("Family", typeof(string));
            dataTable2.Columns.Add("Family and Type", typeof(string));
            dataTable2.Columns.Add("Material: Area", typeof(string));
            dataTable2.Columns.Add("Material: Volume", typeof(string));
            dataTable2.Columns.Add("Material: Name", typeof(string));

            // Select the first worksheet from the file.
            var worksheet2 = workbook2.Worksheets[0];

            // Extract the data from an Excel worksheet to the DataTable.
            // Data is extracted starting at first row and first column for 10 rows or until the first empty row appears.
            worksheet2.ExtractToDataTable(dataTable2, options);


            dataGridView1.DataSource = dataTable2;

            int i = 0;
            foreach (DataRow dt1_Row in dataTable1.Rows)
            {
                string dt1_EID = dt1_Row["ELEMENT ID"].ToString();
                if(dt1_EID == "ELEMENT ID") { continue; }

                foreach (DataRow dt2_Row in dataTable2.Rows)
                {
                    string dt2_EID = dt2_Row["ELEMENT ID"].ToString();
                    if (dt2_EID == "ELEMENT ID") { continue; }

                    if (dt1_EID == dt2_EID)
                    {
                        dgv_DataRecords.Rows.Add();
                        dgv_DataRecords.Rows[i].Cells["ElementID"].Value = dt1_Row["ELEMENT ID"].ToString();
                        dgv_DataRecords.Rows[i].Cells["Name"].Value = dt1_Row["NAME"].ToString();
                        dgv_DataRecords.Rows[i].Cells["FamilyandType"].Value = dt2_Row["Family and Type"].ToString();
                        dgv_DataRecords.Rows[i].Cells["Area"].Value = dt2_Row["Material: Area"].ToString();
                        dgv_DataRecords.Rows[i].Cells["Volume"].Value = dt2_Row["Material: Volume"].ToString();
                        dgv_DataRecords.Rows[i].Cells["StartDate"].Value = dt1_Row["START DATE"].ToString();
                        dgv_DataRecords.Rows[i].Cells["EndDate"].Value = dt1_Row["END DATE"].ToString();
                        i++;
                    }


                }

            }

            

        }

        private void btn_Validate_Click(object sender, EventArgs e)
        {
            unit = cmb_Input_Units.SelectedItem.ToString();
            material = cmb_Input_Material.SelectedItem.ToString();
            TQ = Convert.ToDouble(txt_Input_TQ.Text);
            CS = Convert.ToDouble(txt_Input_CS.Text);
            LT = Convert.ToDouble(txt_Input_LT.Text);
            ALT = Convert.ToDouble(txt_Input_ALT.Text);
            ADC = Convert.ToDouble(txt_Input_ADC.Text);
            MLT = Convert.ToDouble(txt_Input_MLT.Text);
            MDC = Convert.ToDouble(txt_Input_MDC.Text);
            ROP = Convert.ToDouble(txt_Input_ROP.Text);
            DC = Convert.ToDouble(txt_Input_DC.Text);

            //RSDEF
            if(DC > ADC)
            {
                RSDEF = (DC - ADC) * ALT;
            }
            else
            {
                RSDEF = 0;
            }

            //SSDEF
            if(LT > ALT)
            {
                SSDEF = (LT - ALT) * ADC;
            }
            else
            {
                SSDEF = 0;
            }


            lbl_Output_RSDEF.Text = RSDEF.ToString();
            lbl_Output_SSDEF.Text = SSDEF.ToString();

        }

        private void CriteriaValidator_Load(object sender, EventArgs e)
        {
            dgv_EmailList_BindData(EmailList_FileName);



        }

        private void dgv_EmailList_BindData(string filePath)
        {
            DataTable dt = new DataTable();
            string[] lines = System.IO.File.ReadAllLines(filePath);
            if (lines.Length > 0)
            {
                //first line to create header
                string firstLine = lines[0];
                string[] headerLabels = firstLine.Split(',');
                foreach (string headerWord in headerLabels)
                {
                    dt.Columns.Add(new DataColumn(headerWord));
                }
                //For Data
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] dataWords = lines[i].Split(',');
                    DataRow dr = dt.NewRow();
                    int columnIndex = 0;
                    foreach (string headerWord in headerLabels)
                    {
                        dr[headerWord] = dataWords[columnIndex++];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (dt.Rows.Count > 0)
            {
                dgv_EmailList.DataSource = dt;
            }

        }

        private void btn_EmailList_Refresh_Click(object sender, EventArgs e)
        {
            dgv_EmailList_BindData(EmailList_FileName);
        }

        private void btn_EmailList_Save_Click(object sender, EventArgs e)
        {
            writeCSV(dgv_EmailList, EmailList_FileName);
        }

        public void writeCSV(DataGridView gridIn, string outputFile)
        {
            //test to see if the DataGridView has any rows
            if (gridIn.RowCount > 0)
            {
                string value = "";
                DataGridViewRow dr = new DataGridViewRow();
                StreamWriter swOut = new StreamWriter(outputFile);

                //write header rows to csv
                for (int i = 0; i <= gridIn.Columns.Count - 1; i++)
                {
                    if (i > 0)
                    {
                        swOut.Write(",");
                    }
                    swOut.Write(gridIn.Columns[i].HeaderText);
                }

                swOut.WriteLine();

                //write DataGridView rows to csv
                for (int j = 0; j <= gridIn.Rows.Count - 2; j++)
                {
                    if (j > 0)
                    {
                        swOut.WriteLine();
                    }

                    dr = gridIn.Rows[j];

                    for (int i = 0; i <= gridIn.Columns.Count - 1; i++)
                    {
                        if (i > 0)
                        {
                            swOut.Write(",");
                        }

                        value = dr.Cells[i].Value.ToString();
                        //replace comma's with spaces
                        value = value.Replace(',', ' ');
                        //replace embedded newlines with spaces
                        value = value.Replace(Environment.NewLine, " ");

                        swOut.Write(value);
                    }
                }
                swOut.Close();
            }
        }

        private void btn_EmailList_PingAll_Click(object sender, EventArgs e)
        {

            for (int rows = 0; rows < dgv_EmailList.Rows.Count - 1; rows++)
            {
                string emailToAddress = dgv_EmailList.Rows[rows].Cells[2].Value.ToString(); // cells 2 as email is in column 2

                try
                {
                    using (MailMessage mail = new MailMessage())
                    {
                        mail.From = new MailAddress(emailFromAddress);

                        mail.To.Add(emailToAddress);

                        mail.Subject = subject;
                        mail.Body = body;
                        mail.IsBodyHtml = true;
                        //mail.Attachments.Add(new Attachment("D:\\TestFile.txt"));//--Uncomment this to send any attachment  
                        using (SmtpClient smtp = new SmtpClient(smtpAddress, portNumber))
                        {
                            smtp.Credentials = new System.Net.NetworkCredential(emailFromAddress, password);
                            smtp.EnableSsl = enableSSL;
                            smtp.Send(mail);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                    
            }

            MessageBox.Show("Emails were sent");

        }

        private void btn_DataRecords_Calculate_Click(object sender, EventArgs e)
        {

            ElementsIds.Clear();
            Areas.Clear();
            Volumes.Clear();
            TotalDurations.Clear();
            DateOfOrderings.Clear();


            for (int i = 0; i < dgv_DataRecords.RowCount - 1; i++)
            {
                double CurrentStock = Convert.ToDouble(dgv_DataRecords.Rows[i].Cells["CurrentStock"].Value);//For calcualtion
                double LeadTime = Convert.ToDouble(dgv_DataRecords.Rows[i].Cells["LeadTime"].Value);//For calcualtion
                double AverageDailyUsage = Convert.ToDouble(dgv_DataRecords.Rows[i].Cells["AverageDailyUsage"].Value);//For calcualtion
                double MaxDailyUsage = Convert.ToDouble(dgv_DataRecords.Rows[i].Cells["MaxDailyUsage"].Value);//For calcualtion
                double MaxLeadTime = Convert.ToDouble(dgv_DataRecords.Rows[i].Cells["MaxLeadTime"].Value);//For calcualtion
                double AverageLeadTime = Convert.ToDouble(dgv_DataRecords.Rows[i].Cells["AverageLeadTime"].Value);//For calcualtion


                int reorder_point = Convert.ToInt32((LeadTime * AverageDailyUsage) + ((MaxDailyUsage * MaxLeadTime) - (AverageDailyUsage * AverageLeadTime)));

                dgv_DataRecords.Rows[i].Cells["ReorderPoint"].Value = reorder_point.ToString();

                string Area = dgv_DataRecords.Rows[i].Cells["Area"].Value.ToString();
                string Volume = dgv_DataRecords.Rows[i].Cells["Volume"].Value.ToString();
                string start_date = dgv_DataRecords.Rows[i].Cells["StartDate"].Value.ToString().Split(' ')[0].Replace('-', '/');
                string end_date = dgv_DataRecords.Rows[i].Cells["EndDate"].Value.ToString().Split(' ')[0].Replace('-', '/');
                CultureInfo provider = CultureInfo.InvariantCulture;
                DateTime StartDate = DateTime.ParseExact(start_date, "dd/MM/yyyy", provider);//For calcualtion
                DateTime EndDate = DateTime.ParseExact(end_date, "dd/MM/yyyy", provider);//For calcualtion

                double TotalDuration = (EndDate - StartDate).TotalDays;
                dgv_DataRecords.Rows[i].Cells["TotalDuration"].Value = TotalDuration.ToString();

                DateTime DateOfOrdering = StartDate.AddDays(-LeadTime);
                dgv_DataRecords.Rows[i].Cells["DateofOrdering"].Value = DateOfOrdering.ToString();



                string YesNo;

                if(reorder_point > CurrentStock)
                {
                    YesNo = "Yes";
                    dgv_DataRecords.Rows[i].Cells["YesorNo"].Value = YesNo;

                    ElementsIds.Add(dgv_DataRecords.Rows[i].Cells["ElementId"].Value.ToString());
                    Areas.Add(dgv_DataRecords.Rows[i].Cells["Area"].Value.ToString());
                    Volumes.Add(dgv_DataRecords.Rows[i].Cells["Area"].Value.ToString());
                    TotalDurations.Add(TotalDuration);
                    DateOfOrderings.Add(DateOfOrdering);


                    btn_DataRecords_SendEmails.Enabled = true;
                }
                else
                {
                    YesNo = "No";
                    dgv_DataRecords.Rows[i].Cells["YesorNo"].Value = YesNo;
                }

            }
        }

        private void btn_DataRecords_SendEmails_Click(object sender, EventArgs e)
        {
            for (int rows = 0; rows < dgv_EmailList.Rows.Count - 1; rows++)
            {
                string emailToAddress = dgv_EmailList.Rows[rows].Cells[2].Value.ToString(); // cells 2 as email is in column 2

                try
                {
                    using (MailMessage mail = new MailMessage())
                    {
                        mail.From = new MailAddress(emailFromAddress);

                        mail.To.Add(emailToAddress);

                        mail.Subject = "REORDER POINT ALERT!";
                        mail.Body = "Here is the list of elements and their details:\n";


                        for(int i = 0; i < ElementsIds.Count; i++)
                        {
                            mail.Body += "ElementID: ";
                            mail.Body += ElementsIds[i].ToString();
                            mail.Body += "\t";

                            mail.Body += "Area: ";
                            mail.Body += Areas[i].ToString();
                            mail.Body += "\t";

                            mail.Body += "Volume: ";
                            mail.Body += Volumes[i].ToString();
                            mail.Body += "\t";

                            mail.Body += "Total Duration: ";
                            mail.Body += TotalDurations[i].ToString();
                            mail.Body += "\t";

                            mail.Body += "Date of Ordering: ";
                            mail.Body += DateOfOrderings[i].ToString();

                            mail.Body += "\n";
                        }

                        

                        mail.IsBodyHtml = true;
                        //mail.Attachments.Add(new Attachment("D:\\TestFile.txt"));//--Uncomment this to send any attachment  
                        using (SmtpClient smtp = new SmtpClient(smtpAddress, portNumber))
                        {
                            smtp.Credentials = new System.Net.NetworkCredential(emailFromAddress, password);
                            smtp.EnableSsl = enableSSL;
                            smtp.Send(mail);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            }

            MessageBox.Show("Emails were sent");
        }

        private void btn_DataRecords_Refresh_Click(object sender, EventArgs e)
        {
            btn_DataRecords_SendEmails.Enabled = false;

            dgv_DataRecords.DataSource = null;
            dgv_DataRecords.Rows.Clear();

            ElementsIds.Clear();
            Areas.Clear();
            Volumes.Clear();
            TotalDurations.Clear();
            DateOfOrderings.Clear();

            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workbook1 = ExcelFile.Load("NAVIS OUT.csv");

            var dataTable1 = new DataTable();

            // Depending on the format of the input file, you need to change this:
            dataTable1.Columns.Add("Material", typeof(string));
            dataTable1.Columns.Add("GUID", typeof(string));
            dataTable1.Columns.Add("ELEMENT ID", typeof(string));
            dataTable1.Columns.Add("NAME", typeof(string));
            dataTable1.Columns.Add("GLOABALID", typeof(string));
            dataTable1.Columns.Add("START DATE", typeof(string));
            dataTable1.Columns.Add("END DATE", typeof(string));
            dataTable1.Columns.Add("DURATION", typeof(string));

            // Select the first worksheet from the file.
            var worksheet1 = workbook1.Worksheets[0];

            var options = new ExtractToDataTableOptions(0, 0, 10000);
            options.ExtractDataOptions = ExtractDataOptions.StopAtFirstEmptyRow;
            options.ExcelCellToDataTableCellConverting += (sender_1, e_1) =>
            {
                if (!e_1.IsDataTableValueValid)
                {
                    // GemBox.Spreadsheet doesn't automatically convert numbers to strings in ExtractToDataTable() method because of culture issues; 
                    // someone would expect the number 12.4 as "12.4" and someone else as "12,4".
                    e_1.DataTableValue = e_1.ExcelCell.Value == null ? null : e_1.ExcelCell.Value.ToString();
                    e_1.Action = ExtractDataEventAction.Continue;
                }
            };

            // Extract the data from an Excel worksheet to the DataTable.
            // Data is extracted starting at first row and first column for 10 rows or until the first empty row appears.
            worksheet1.ExtractToDataTable(dataTable1, options);


            //-----------------Again

            var workbook2 = ExcelFile.Load("REVIT output.xlsx");

            var dataTable2 = new DataTable();

            // Depending on the format of the input file, you need to change this:
            dataTable2.Columns.Add("Count", typeof(string));
            dataTable2.Columns.Add("ELEMENT ID", typeof(string));
            dataTable2.Columns.Add("Description", typeof(string));
            dataTable2.Columns.Add("Family", typeof(string));
            dataTable2.Columns.Add("Family and Type", typeof(string));
            dataTable2.Columns.Add("Material: Area", typeof(string));
            dataTable2.Columns.Add("Material: Volume", typeof(string));
            dataTable2.Columns.Add("Material: Name", typeof(string));

            // Select the first worksheet from the file.
            var worksheet2 = workbook2.Worksheets[0];

            // Extract the data from an Excel worksheet to the DataTable.
            // Data is extracted starting at first row and first column for 10 rows or until the first empty row appears.
            worksheet2.ExtractToDataTable(dataTable2, options);


            dataGridView1.DataSource = dataTable2;

            int i = 0;
            foreach (DataRow dt1_Row in dataTable1.Rows)
            {
                string dt1_EID = dt1_Row["ELEMENT ID"].ToString();
                if (dt1_EID == "ELEMENT ID") { continue; }

                foreach (DataRow dt2_Row in dataTable2.Rows)
                {
                    string dt2_EID = dt2_Row["ELEMENT ID"].ToString();
                    if (dt2_EID == "ELEMENT ID") { continue; }

                    if (dt1_EID == dt2_EID)
                    {
                        dgv_DataRecords.Rows.Add();
                        dgv_DataRecords.Rows[i].Cells["ElementID"].Value = dt1_Row["ELEMENT ID"].ToString();
                        dgv_DataRecords.Rows[i].Cells["Name"].Value = dt1_Row["NAME"].ToString();
                        dgv_DataRecords.Rows[i].Cells["FamilyandType"].Value = dt2_Row["Family and Type"].ToString();
                        dgv_DataRecords.Rows[i].Cells["Area"].Value = dt2_Row["Material: Area"].ToString();
                        dgv_DataRecords.Rows[i].Cells["Volume"].Value = dt2_Row["Material: Volume"].ToString();
                        dgv_DataRecords.Rows[i].Cells["StartDate"].Value = dt1_Row["START DATE"].ToString();
                        dgv_DataRecords.Rows[i].Cells["EndDate"].Value = dt1_Row["END DATE"].ToString();
                        i++;
                    }


                }

            }
        }

        private void dgv_DataRecords_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            btn_DataRecords_SendEmails.Enabled = false;
        }
    }
}
