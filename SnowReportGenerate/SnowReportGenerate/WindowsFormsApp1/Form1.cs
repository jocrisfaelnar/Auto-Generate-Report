
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using msExcel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public static List<string> ApmList = new List<string>();
        public static string workBookName;
        public static string taskSource;
        string emailItem = "";
        public static string btnUsed;
        string tableEmailHead = "";
        string sentTo;
        string subject;
        string countsMessage = "<div  style = 'font-size:9pt; font-family:calibri;'>Hi Team,</div><br><div style = 'font-size:9pt; font-family:calibri;'>Please see open snow tickets for today. Kindly revisit your ticket/s and close all that can be closed.</div><br><br><p style = 'font-size:9pt; font-family:calibri;'><b>I. Tickets Summary</b></p>";
        string snowSourceFolder;
        object reportSourceFolder;
        Form2 form2 = new Form2();
        DataTable dtUnAssigned = new DataTable();
        DataTable dtApm = new DataTable();
        DataTable dtOps = new DataTable();
        DataTable dtApmEm = new DataTable();
        DataTable dtOpsEm = new DataTable();
        DateTime toDay = DateTime.Now;
        DateTime updatedDate;
        DateTime reportedDate;
        DateTime closedDate;
        public static DateTime lastDate;
        public static int openTickets = 0;
        public static int closedResolved = 0;
        public static int activeTickets = 0;
        public static int workInprogress = 0;
        public static int awaiting = 0;
        public static int totalTickets = 0;
        public static int severetyLow = 0;
        public static int severetyModerate = 0;
        public static int severetyHigh = 0;
        public static int severetyCritical = 0;

        bool isEvening = false;
        // Uncomment to show Console wimdow
        //internal sealed class NativeMethods
        //{
        //    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        //    public static extern bool AllocConsole();

        //    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        //    public static extern bool FreeConsole();
        //}
        public Form1()
        {
            InitializeComponent();
        }

        private void btnMorning_Click(object sender, EventArgs e)
        {
            isEvening = false;
            ResetData();

            // Read APM Members.Text ------------------------
            var lines = File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "APM Members.txt")); 
            for (var i = 0; i < lines.Length; i += 1)
            {
                var line = lines[i];
                ApmList.Add(line.ToString());
            }
            //  --------------------------------------------

            if (openFileDialog.ShowDialog() == DialogResult.OK) // -- Open IncTickets
            {
                string sourceFile = openFileDialog.FileName;
                snowSourceFolder = openFileDialog.FileName;
                if (openFileDialog.ShowDialog() == DialogResult.OK) // -- Open TaskTickets
                {
                    
                    taskSource = openFileDialog.FileName;
                    if (openFileDialog.ShowDialog() == DialogResult.OK) // -- Open Report
                    {

                        Hide();
                        btnUsed = " am";
                        //try
                        //{
                            if (cbTeamSelector.SelectedItem.ToString() != null)
                            {
                                
                                sourceFile = openFileDialog.FileName;
                                reportSourceFolder = openFileDialog.FileName;
                                //NativeMethods.AllocConsole();   // Uncomment - to release console window
                                Console.WriteLine("Getting Last Working Days");
                                Console.WriteLine("Start Gathering Data.");
                            PapulateExcel();
                            }
                            else
                            {
                                MessageBox.Show("Please Select Team!");
                            }
                        //}
                        //catch (NullReferenceException ex) { MessageBox.Show(ex.Message + "\n \n Must Select a Team!"); }
                        //catch (FormatException fe) { MessageBox.Show(fe.Message); }
                        //catch (DuplicateNameException dn) { MessageBox.Show(dn.Message + "\n \n Please Use Another name"); }
                        //catch (Exception allex) { MessageBox.Show(allex.Message); }
                        Show();
                    }
                }

            }
        }
        private void btnEvening_Click(object sender, EventArgs e)
        {
            isEvening = true;
            ResetData();

            // Read APM Members.Text ------------------------
            var lines = File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "APM Members.txt"));
            for (var i = 0; i < lines.Length; i += 1)
            {
                var line = lines[i];
                ApmList.Add(line.ToString());
            }
            //  --------------------------------------------

            if (openFileDialog.ShowDialog() == DialogResult.OK) // -- Open Inc Tickets
            {
                string sourceFile = openFileDialog.FileName;
                snowSourceFolder = openFileDialog.FileName;
                if (openFileDialog.ShowDialog() == DialogResult.OK) // -- Open TaskTickets
                {
                    taskSource = openFileDialog.FileName;
                    try
                    {
                        if (openFileDialog.ShowDialog() == DialogResult.OK) // -- Open Report
                        {
                            Hide();
                            btnUsed = " pm";
                            if (cbTeamSelector.SelectedItem.ToString() != null)
                            {
                                sourceFile = openFileDialog.FileName;
                                reportSourceFolder = openFileDialog.FileName;
                                //NativeMethods.AllocConsole();  // Uncomment - to release console window
                                Console.WriteLine("Start Gathering Data.");
                                PapulateExcel();
                            }
                            else
                            {
                                MessageBox.Show("Please Select Team!");
                            }
                        }
                    }
                    catch (NullReferenceException ex) { MessageBox.Show(ex.Message + "\n \n Must Select a Team!"); }
                    catch (FormatException fe) { MessageBox.Show(fe.Message); }
                    catch (DuplicateNameException dn) { MessageBox.Show(dn.Message + "\n \n Please Use Another name"); }
                    catch (Exception allex) { MessageBox.Show(allex.Message); }
                    Show();
                }
            }

        }
        private void EditReportMorning(string source)
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(source)))
            {
                int dtcoll = 0;
                int h = 0;
                dtApm.Clear();
                dtOps.Clear();
                dtUnAssigned.Clear();
                dtOpsEm.Clear();
                dtApmEm.Clear();
                var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                string[] row = new string[totalColumns];
                for (int rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                {
                    var cell = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                    foreach (var data in cell)
                    {
                        //Adding The Headder And Columns
                        if (h != 1)
                        {
                            if (dtApm.Columns.Count != totalColumns)
                            {
                                dtApm.Columns.Add(data);
                                dtUnAssigned.Columns.Add(data);
                                dtOps.Columns.Add(data);
                                dtOpsEm.Columns.Add(data);
                                dtApmEm.Columns.Add(data);
                            }
                            else
                            {
                                h = 1;
                            }
                        }
                        //Populating the Tables / Dividing Data
                        if (h == 1)
                        {
                            if (dtcoll != totalColumns - 1)
                            {
                                row[dtcoll] = data;
                                dtcoll++;
                            }
                            else
                            {

                                try
                                {
                                    updatedDate = ConvertDates(row[7].ToString());
                                    row[7] = updatedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                    closedDate = ConvertDates(row[11].ToString());
                                    row[11] = closedDate.ToString("MM/dd/yyyy hh:mm:ss");

                                }
                                catch (Exception ex)
                                {
                                    ex.ToString();
                                    updatedDate = new DateTime(1990, 01, 01);
                                    closedDate = new DateTime(1990, 01, 01);
                                }
                                dtcoll = 0;
                                Console.WriteLine(lastDate.Date.ToString("MM/dd/yyyy") +" " + closedDate.Date.ToString("MM/dd/yyyy"));
                                //Populating the Tables
                                if (lastDate.Date <= closedDate.Date || (row[11] == "" || row[11] == null))
                                {
                                    Console.WriteLine("True");
                                    string severety = row[4];
                                    switch (severety)
                                    {
                                        case "4 - Low":
                                            severetyLow = severetyLow + 1;
                                            break;
                                        case "3 - Moderate":
                                            severetyModerate = severetyModerate + 1;
                                            break;
                                        case "2 - High":
                                            severetyHigh = severetyHigh + 1;
                                            break;
                                        case "1 - Critical":
                                            severetyCritical = severetyCritical + 1;
                                            break;
                                        default:
                                            break;
                                    }
                                    if (row[2] == "Closed" || row[2] == "Resolved")
                                    {
                                        closedResolved = closedResolved + 1;
                                    }
                                    else if (row[2] == "Awaiting User Info" || row[2] == "Awaiting External 3rd Party Action" ||
                                        row[2] == "Awaiting Evidence" || row[2] == "Awaiting Change" || row[2] == "Awaiting Problem"
                                        )
                                    {
                                        awaiting = awaiting + 1;
                                    }
                                    else if (row[2] == "Work In Progress")
                                    {
                                        workInprogress = workInprogress + 1;
                                    }
                                    else if (row[2] == "Active" || row[2] == "New")
                                    {
                                        activeTickets = activeTickets + 1;
                                    }
                                    openTickets = workInprogress + awaiting + activeTickets;
                                    totalTickets = closedResolved + openTickets;

                                    if (cbTeamSelector.SelectedItem.ToString() == "Collab Report")
                                    {
                                        sentTo = "Manila-KMAD-PeopleSearch-Operations <Manila-KMAD-PeopleSearch-Operations@accenture.com>; Collab-ManilaAPM <Collab-ManilaAPM@accenture.com>";
                                        subject = "Collab Operations SNOW Status Report - " + toDay.ToString("MMMM dd, yyyy");
                                        //Populating the UnAssigned Table
                                        reportedDate = ConvertDates(row[5].ToString());
                                        row[5] = reportedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                        if (row[3] == "" || row[3] == null)
                                        {
                                            dtUnAssigned.Rows.Add(row);
                                        }
                                        else if (ApmList.Contains(row[3].ToString()))
                                        {
                                            dtApm.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {

                                            }
                                            else
                                            {
                                                dtApmEm.Rows.Add(row);
                                            }
                                            //Console.WriteLine(dtApm.Rows.Count);
                                        }
                                        //Populating the Ops Table
                                        else
                                        {
                                            dtOps.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {

                                            }
                                            else
                                            {
                                                dtOpsEm.Rows.Add(row);
                                            }
                                        }
                                    }
                                    else if (cbTeamSelector.SelectedItem.ToString() == "KX Report")
                                    {
                                        sentTo = "Manila-KMAD-KXGroups-Operations <Manila-KMAD-KXGroups-Operations@accenture.com>; Collab-ManilaAPM <Collab-ManilaAPM@accenture.com>";
                                        subject = "KX Operations SNOW Status Report - " + toDay.ToString("MMMM dd, yyyy");
                                        reportedDate = ConvertDates(row[5].ToString());
                                        row[5] = reportedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                        Console.WriteLine("True");
                                        if (row[3] == "" || row[3] == null)
                                        {
                                            dtUnAssigned.Rows.Add(row);
                                        }
                                        else if (ApmList.Contains(row[3].ToString()))
                                        {
                                            dtApm.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {

                                            }
                                            else
                                            {
                                                dtApmEm.Rows.Add(row);
                                            }
                                            //Console.WriteLine(dtApm.Rows.Count);
                                        }
                                        //Populating the Ops Table
                                        else
                                        {
                                            dtOps.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {
                                            }
                                            else
                                            {
                                                dtOpsEm.Rows.Add(row);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Please Select Team");
                                    }
                                }
                            }
                        }
                    }
                    
                }
            }
            // Sort Tables
            dtApm = SortTables(dtApm);
            dtOps = SortTables(dtOps);
            dtUnAssigned = SortTables(dtUnAssigned);
            dtApmEm = SortTables(dtApmEm);
            dtOpsEm = SortTables(dtOpsEm);

            // Combine Inc Table to Task Tables

            CombineTables(dtApm, TaskClass.dtTaskApm);
            CombineTables(dtOps, TaskClass.dtTaskOps);
            CombineTables(dtUnAssigned, TaskClass.dtTaskUnAssigned);
            CombineTables(dtApmEm, TaskClass.dtTaskApmEm);
            CombineTables(dtOpsEm, TaskClass.dtTaskOpsEm);
        }
        public void EditReportEvening(string source)
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(source)))
            {
                int dtcoll = 0;
                int h = 0;
                dtApm.Clear();
                DateTime rawResolvedDate;
                var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                string[] row = new string[totalColumns];
                for (int rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                {
                    var cell = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                    foreach (var data in cell)
                    {
                        //Adding The Headder And Columns
                        if (h != 1)
                        {
                            if (dtApm.Columns.Count != totalColumns)
                            {
                                dtApm.Columns.Add(data);
                                dtUnAssigned.Columns.Add(data);
                                dtOps.Columns.Add(data);
                                dtOpsEm.Columns.Add(data);
                                dtApmEm.Columns.Add(data);
                            }
                            else
                            {
                                h = 1;
                            }
                        }
                        //Populating the Tables / Dividing Data
                        if (h == 1)
                        {
                            if (dtcoll != totalColumns - 1)
                            {

                                row[dtcoll] = data;

                                //Console.WriteLine(row[dtcoll] + "...Ok");
                                dtcoll++;
                            }
                            else
                            {

                                updatedDate = ConvertDates(row[7].ToString());
                                row[7] = updatedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                try
                                {
                                    string rawResolved = data;
                                    rawResolvedDate = ConvertDates(rawResolved.ToString());

                                }
                                catch (Exception ex)
                                {
                                    rawResolvedDate = new DateTime(1990, 01, 01);
                                    ex.ToString();
                                }
                                try
                                {
                                    closedDate = ConvertDates(row[11].ToString());
                                    row[11] = closedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                }
                                catch (Exception ex)
                                {
                                    closedDate = new DateTime(1990, 01, 01);
                                    ex.ToString();
                                }
                                dtcoll = 0;
                                Console.WriteLine(rawResolvedDate.Date.ToString() + "' " + row[2].ToString() + "' " + closedDate.Date.ToString() + "' " + row[11].ToString());
                                //Populating the Tables
                                if ((rawResolvedDate.Date == DateTime.Now.Date && row[2].ToString() == "Resolved") ||
                                    row[2] == "Awaiting User Info" ||
                                    row[2] == "Awaiting External 3rd Party Action" ||
                                    row[2] == "Awaiting Evidence" ||
                                    row[2] == "Awaiting Change" ||
                                    row[2] == "Awaiting Problem" ||
                                    row[2] == "Work In Progress" ||
                                   (row[2].ToString() == "Closed"
                                   && closedDate.Date == DateTime.Now.Date))
                                {
                                    Console.WriteLine("True");
                                    string severety = row[4];
                                    switch (severety)
                                    {
                                        case "4 - Low":
                                            severetyLow = severetyLow + 1;
                                            break;
                                        case "3 - Moderate":
                                            severetyModerate = severetyModerate + 1;
                                            break;
                                        case "2 - High":
                                            severetyHigh = severetyHigh + 1;
                                            break;
                                        case "1 - Critical":
                                            severetyCritical = severetyCritical + 1;
                                            break;
                                        default:
                                            break;
                                    }
                                    if (row[2] == "Closed" || row[2] == "Resolved")
                                    {
                                        closedResolved = closedResolved + 1;
                                    }
                                    else if (row[2] == "Awaiting User Info" || row[2] == "Awaiting External 3rd Party Action" ||
                                        row[2] == "Awaiting Evidence" || row[2] == "Awaiting Change" || row[2] == "Awaiting Problem"
                                        )
                                    {
                                        awaiting = awaiting + 1;
                                    }
                                    else if (row[2] == "Work In Progress")
                                    {
                                        workInprogress = workInprogress + 1;
                                    }
                                    else if (row[2] == "Active" || row[2] == "New")
                                    {
                                        activeTickets = activeTickets + 1;
                                    }
                                    openTickets = workInprogress + awaiting + activeTickets;
                                    totalTickets = closedResolved + openTickets;
                                    if (cbTeamSelector.SelectedItem.ToString() == "Collab Report")
                                    {
                                        sentTo = "Manila-KMAD-PeopleSearch-Operations <Manila-KMAD-PeopleSearch-Operations@accenture.com>; Collab-ManilaAPM <Collab-ManilaAPM@accenture.com>";
                                        subject = "Collab Operations SNOW Status Report - " + toDay.ToString("MMMM dd, yyyy");
                                        //Populating the UnAssigned Table
                                        reportedDate = ConvertDates(row[5].ToString());
                                        row[5] = reportedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                        //Console.WriteLine("True");
                                        if (row[3] == "" || row[3] == null)
                                        {
                                            dtUnAssigned.Rows.Add(row);
                                        }
                                        else if (ApmList.Contains(row[3].ToString()))
                                        {
                                            dtApm.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {

                                            }
                                            else
                                            {
                                                dtApmEm.Rows.Add(row);
                                            }
                                            //Console.WriteLine(dtApm.Rows.Count);
                                        }
                                        //Populating the Ops Table
                                        else
                                        {
                                            dtOps.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {

                                            }
                                            else
                                            {
                                                dtOpsEm.Rows.Add(row);
                                            }
                                        }
                                    }
                                    else if (cbTeamSelector.SelectedItem.ToString() == "KX Report")
                                    {
                                        subject = "KX Operations SNOW Status Report - " + toDay.ToString("MMMM dd, yyyy");
                                        sentTo = "Manila-KMAD-KXGroups-Operations <Manila-KMAD-KXGroups-Operations@accenture.com>; Collab-ManilaAPM <Collab-ManilaAPM@accenture.com>";
                                        //Populating the UnAssigned Table
                                        reportedDate = ConvertDates(row[5].ToString());
                                        row[5] = reportedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                        //Console.WriteLine("True");
                                        if (row[3] == "" || row[3] == null)
                                        {
                                            dtUnAssigned.Rows.Add(row);
                                        }
                                        else if (ApmList.Contains(row[3].ToString()))
                                        {
                                            dtApm.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {

                                            }
                                            else
                                            {
                                                dtApmEm.Rows.Add(row);
                                            }
                                            //Console.WriteLine(dtApm.Rows.Count);
                                        }
                                        //Populating the Ops Table
                                        else
                                        {
                                            dtOps.Rows.Add(row);
                                            if (row[2] == "Resolved" || row[2] == "Closed")
                                            {

                                            }
                                            else
                                            {
                                                dtOpsEm.Rows.Add(row);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Please Select Team");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("false");
                                }

                            }
                        }
                    }
                    
                }

            }

            // Sort Tables
            dtApm = SortTables(dtApm);
            dtOps = SortTables(dtOps);
            dtUnAssigned = SortTables(dtUnAssigned);
            dtApmEm = SortTables(dtApmEm);
            dtOpsEm = SortTables(dtOpsEm);

            // Combine IncTable to Task Tables

            CombineTables(dtApm, TaskClass.dtTaskApm);
            CombineTables(dtOps, TaskClass.dtTaskOps);
            CombineTables(dtUnAssigned, TaskClass.dtTaskUnAssigned);
            CombineTables(dtApmEm, TaskClass.dtTaskApmEm);
            CombineTables(dtOpsEm, TaskClass.dtTaskOpsEm);
        }

        public void PapulateExcel()
        {
            Console.WriteLine("Copying..");
            Console.WriteLine("Starting Exel App..");
            //Microsoft Excel Variables
            Object oMissing = System.Reflection.Missing.Value;
            msExcel.Application excelApp = new msExcel.Application();
            Object oTemplatePath = reportSourceFolder;
            msExcel.Workbook currentWorkbook = excelApp.Workbooks.Add(oTemplatePath);
            msExcel.Worksheet currentWorksheet = (msExcel.Worksheet)currentWorkbook.ActiveSheet;

            //Remove 'am' and 'pm' and get Last date using sheet name
            lastDate = Convert.ToDateTime(Regex.Replace(currentWorksheet.Name.Trim(), "[apm]", string.Empty)); 
            Console.WriteLine("Last Date Data is: " + lastDate.ToShortDateString());
           
            //false value will hide excel while ploting the data
            excelApp.Visible = false;

            //Start Processing Task Ticket.
            TaskClass.GetTaskData(taskSource);
            //Start Processing INC Tickets.
            if (!isEvening) { EditReportMorning(snowSourceFolder); }
            else { EditReportEvening(snowSourceFolder); }

            
            currentWorksheet = (msExcel.Worksheet)currentWorkbook.Sheets[1];
            currentWorkbook.Sheets[currentWorkbook.Sheets.Count].Name = toDay.ToString("MM-dd-yyyy") + btnUsed;
            currentWorksheet = currentWorkbook.Sheets[currentWorkbook.Sheets.Count];
            currentWorksheet.Activate();
            excelApp.DisplayAlerts = false;
            msExcel.Range sourceRange = currentWorkbook.Sheets[1].Range["A16:L17"];
            sourceRange.Copy();

            CopyTables(dtUnAssigned.Rows.Count + 19, currentWorksheet, excelApp, "OPS");
            CopyTables(dtOps.Rows.Count + dtUnAssigned.Rows.Count + 22, currentWorksheet, excelApp, "APM");

            currentWorksheet.Cells[15, 1] = "Unassigned";
            currentWorksheet.Cells[dtUnAssigned.Rows.Count + 18, 1] = "OPS";
            ColorText(currentWorksheet, dtUnAssigned.Rows.Count + 18, 1);
            currentWorksheet.Cells[dtOps.Rows.Count + dtUnAssigned.Rows.Count + 21, 1] = "APM";
            ColorText(currentWorksheet, dtOps.Rows.Count + dtUnAssigned.Rows.Count + 21, 1);
            tableEmailHead = "<table border='1' bordercolor='solid black' style='font-family:calibri; font-size:9pt;'><tr bgcolor='#008080' style = 'color:White'>";
            //------------------------
            for (int i = 0; i < dtApm.Columns.Count - 2; i++)
            {
                var header = dtApm.Columns[i].ColumnName.ToString();
                tableEmailHead = tableEmailHead + "<td><center><b>" + header + "</b></center></td>";
            }
            if (dtUnAssigned.Rows.Count != 0)
            {
                tableEmailHead = tableEmailHead + "</tr>";
                //Sending Unassiagned Table Data To Excel
                emailItem = emailItem + "<div style='color:red; font-family:calibri;font-size:12pt;'><b>Unassigned</b></div>" + tableEmailHead;
                for (int i = 0; i <= dtUnAssigned.Rows.Count - 1; i++)
                {
                    currentWorksheet.Cells[i + 17, 1] = " ";
                    currentWorksheet.Cells[17, 1].Offset[i].Resize[1, dtUnAssigned.Columns.Count].Value =
                    dtUnAssigned.Rows[i].ItemArray;
                    emailItem = emailItem + "<tr>";
                    for (int j = 0; j <= dtUnAssigned.Columns.Count - 3; j++)
                    {
                        var cell = dtUnAssigned.Rows[i].Field<string>(j);
                        emailItem = emailItem + "<td>" + cell + "</td>";

                    }
                    emailItem = emailItem + "</tr>";
                }

                emailItem = emailItem + "</table> <br> <br>";
            }

            //Sending Apm Table Data To Excel
            emailItem = emailItem + "<div style='color:red; font-family:calibri;font-size:12pt;'><b>OPS</b></div>" + tableEmailHead;

            for (int i = 0; i <= dtOps.Rows.Count - 1; i++)
            {
                currentWorksheet.Cells[i + dtUnAssigned.Rows.Count + 20, 1] = " ";
                currentWorksheet.Cells[dtUnAssigned.Rows.Count + 20, 1].Offset[i].Resize[1, dtOps.Columns.Count].Value = dtOps.Rows[i].ItemArray;
            }

            for (int i = 0; i <= dtOpsEm.Rows.Count - 1; i++)
            {
                emailItem = emailItem + "<tr>";
                for (int j = 0; j <= dtOpsEm.Columns.Count - 3; j++)
                {
                    var cell = dtOpsEm.Rows[i].Field<string>(j);
                    emailItem = emailItem + "<td>" + cell + "</td>";
                }
                emailItem = emailItem + "</tr>";
            }

            emailItem = emailItem + "</table> <br> <br>";
            //Sending Ops Table Data To Excel
            emailItem = emailItem + "<div style='color:red; font-family:calibri;font-size:12pt;'><b>APM</b></div>" + tableEmailHead;

            for (int i = 0; i <= dtApm.Rows.Count - 1; i++)
            {
                currentWorksheet.Cells[i + dtOps.Rows.Count + dtUnAssigned.Rows.Count + 23, 1] = " ";
                currentWorksheet.Cells[dtOps.Rows.Count + dtUnAssigned.Rows.Count + 23, 1].Offset[i].Resize[1, dtApm.Columns.Count].Value = dtApm.Rows[i].ItemArray;
            }
            for (int i = 0; i <= dtApmEm.Rows.Count - 1; i++)
            {
                emailItem = emailItem + "<tr>";
                for (int j = 0; j <= dtApmEm.Columns.Count - 3; j++)
                {
                    var cell = dtApmEm.Rows[i].Field<string>(j);
                    emailItem = emailItem + "<td>" + cell + "</td>";
                }
                emailItem = emailItem + "</tr>";
            }

            emailItem = emailItem + "</table> <br> <br></body></html>";

            CreateSumarry(currentWorksheet);
            //Creating Email item
            CreateEmailItem();
            if (!isEvening)
            {
                try
                {
                    CreateMailItem();
                    //Save Excel file
                    excelApp.DisplayAlerts = false;
                    currentWorksheet.SaveAs(reportSourceFolder.ToString());
                    currentWorkbook.Close();
                    excelApp.Quit();
                    System.Diagnostics.Process.Start(reportSourceFolder.ToString());
                }
                catch (Exception ex)
                {
                    //Save Excel file
                    Console.Write(ex.Message);
                    totalTickets = 0;
                    closedResolved = 0;
                    openTickets = 0;
                    workInprogress = 0;
                    activeTickets = 0;
                    countsMessage = "<div  style = 'font-size:9pt; font-family:calibri;'>Hi Team,</div>" +
                                       "<br<div style = 'font-size:9pt; font-family:calibri;'>Please see open snow tickets for today. Kindly revisit your ticket/s and close all that can be closed.</div>" +
                                       "<br><br><p><div style = 'font-size:9pt; font-family:calibri;'>I. Tickets Summary</b></div>";
                    excelApp.DisplayAlerts = false;
                    currentWorksheet.SaveAs(reportSourceFolder.ToString());
                    currentWorkbook.Close();
                    excelApp.Quit();
                    System.Diagnostics.Process.Start(reportSourceFolder.ToString());
                }
            }
            // if evening  system will not generate Email item
            else if (isEvening)
            {
                excelApp.DisplayAlerts = false;
                currentWorksheet.SaveAs(reportSourceFolder.ToString());
                currentWorkbook.Close();
                excelApp.Quit();
                System.Diagnostics.Process.Start(reportSourceFolder.ToString());
            }
            //<----------------------
            totalTickets = 0;
            closedResolved = 0;
            openTickets = 0;
            workInprogress = 0;
            activeTickets = 0;
            countsMessage = "<div style = 'font-size:9pt; font-family:calibri;'>Hi Team,</div>" +
                               "<br><div style = 'font-size:9pt; font-family:calibri;'>Please see open snow tickets for today. Kindly revisit your ticket/s and close all that can be closed.</div>" +
                               "<br><br><div style = 'font-size:9pt; font-family:calibri;'><b>I. Tickets Summary</b></div>";
        }

        public void ColorText(msExcel.Worksheet currentWorksheet, int row, int col)
        {
            msExcel.Range range = currentWorksheet.Cells[row, col];
            range.HorizontalAlignment = msExcel.XlHAlign.xlHAlignCenter;
            range.Font.Color = ColorTranslator.ToOle(Color.Red);
        }

        public void CreateMailItem()
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = subject;
            mailItem.To = sentTo;
            mailItem.HTMLBody = "<div>" + countsMessage + emailItem + "</div>";
            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            mailItem.Display(true);
            mailItem.Close(Outlook.OlInspectorClose.olDiscard);
        }

        private void CopyTables(int row, msExcel.Worksheet currentWorksheet, msExcel.Application excelApp, string tableName)
        {
            msExcel.Range destinationRange = currentWorksheet.Cells[row,1];
            currentWorksheet.ListObjects.Add(msExcel.XlListObjectSourceType.xlSrcRange, destinationRange,
            Type.Missing, msExcel.XlYesNoGuess.xlYes, Type.Missing).Name = tableName;
            currentWorksheet.ListObjects[tableName].TableStyle = "TableStyleMedium22";
            destinationRange.PasteSpecial(msExcel.XlPasteType.xlPasteAll);
            excelApp.Visible = true;
        }

        //Create summarry values in excel
        private void CreateSumarry(msExcel.Worksheet currentWorksheet)
        {
            currentWorksheet.Cells[6, 2] = "Snow Tickets as of " + toDay.ToString("M/dd/yyyy");
            currentWorksheet.Cells[3, 2] = totalTickets.ToString();
            currentWorksheet.Cells[3, 3] = closedResolved.ToString();
            currentWorksheet.Cells[3, 4] = openTickets.ToString();

            currentWorksheet.Cells[4, 2] = TaskClass.tasktotalTickets.ToString();
            currentWorksheet.Cells[4, 3] = TaskClass.taskclosedResolved.ToString();
            currentWorksheet.Cells[4, 4] = TaskClass.taskopenTickets.ToString();

            currentWorksheet.Cells[8, 2] = severetyCritical.ToString();
            currentWorksheet.Cells[8, 3] = severetyHigh.ToString();
            currentWorksheet.Cells[8, 4] = severetyModerate.ToString();
            currentWorksheet.Cells[8, 5] = severetyLow.ToString();

            currentWorksheet.Cells[9, 2] = TaskClass.taskCritical.ToString();
            currentWorksheet.Cells[9, 3] = TaskClass.taskHigh.ToString();
            currentWorksheet.Cells[9, 4] = TaskClass.taskModerate.ToString();
            currentWorksheet.Cells[9, 5] = TaskClass.taskLow.ToString();

            currentWorksheet.Cells[12, 2] = activeTickets.ToString();
            currentWorksheet.Cells[12, 3] = workInprogress.ToString();
            currentWorksheet.Cells[12, 4] = awaiting.ToString();
            currentWorksheet.Cells[12, 5] = closedResolved.ToString();

            currentWorksheet.Cells[13, 2] = TaskClass.taskactiveTickets.ToString();
            currentWorksheet.Cells[13, 3] = TaskClass.taskworkInprogress.ToString();
            currentWorksheet.Cells[13, 4] = TaskClass.taskawaiting.ToString();
            currentWorksheet.Cells[13, 5] = TaskClass.taskclosedResolved.ToString();
        }

        //Create Summarry Tables in outlook
        private void CreateEmailItem()
        {
            countsMessage = countsMessage + "<table border='1' bordercolor='solid black' table style='font-family:calibri;font-size:9pt;'>" +
                                               "<tr bgcolor='#008080' style = 'color:White; text-align:center;'>" +
                                                  "<td><b>SUMMARY:</b></td>" +
                                                  "<td><b>Total</b></td>" +
                                                  "<td><b>Closed</b></td>" +
                                                  "<td><b>Open Tickets</b></td>" +
                                                "</tr>" +
                                                "<tr style = 'color:red; text-align:center;'>" +
                                                   "<td>Incident</td>" +
                                                   "<td>" + totalTickets.ToString() + "</td>" +
                                                   "<td>" + closedResolved.ToString() + "</td>" +
                                                   "<td>" + openTickets.ToString() + "</td>" +
                                                "</tr>" +
                                                "<tr style = 'color:red; text-align:center;'>" +
                                                   "<td>RITM</td>" +
                                                   "<td>" + TaskClass.tasktotalTickets.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskclosedResolved.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskopenTickets.ToString() + "</td>" +
                                                "</tr>" +
                                            "</table><br><br>";
            countsMessage = countsMessage + "<table border='1' bordercolor='solid black' style='font-family:calibri;font-size:9pt;'>" +
                                                "<tr bgcolor='#008080' style = 'color:White; text-align:center;'>" +
                                                   "<td><b> Ticket </b></td>" +
                                                   "<td><b>1 - Critical</b></td>" +
                                                   "<td><b>2 - High</b></td>" +
                                                   "<td><b>3 - Medium</b></td>" +
                                                   "<td><b>4 - Low</b></td>" +
                                                "</tr>" +
                                                "<tr style = 'color:red; text-align:center;'>" +
                                                   "<td>Incident</td>" +
                                                   "<td>" + severetyCritical.ToString() + "</td>" +
                                                   "<td>" + severetyHigh.ToString() + "</td>" +
                                                   "<td>" + severetyModerate.ToString() + "</td>" +
                                                   "<td>" + severetyLow.ToString() + "</td>" +
                                                "</tr>" +
                                                "<tr style = 'color:red; text-align:center;'>" +
                                                   "<td>RITM</td>" +
                                                   "<td>" + TaskClass.taskCritical.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskHigh.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskModerate.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskLow.ToString() + "</td>" +
                                                "</tr>" +
                                             "</table><br><br>";
            countsMessage = countsMessage + "<table border='1' bordercolor='solid black' style='font-family:calibri;font-size:9pt;'>" +
                                                "<tr bgcolor='#008080' style = 'color:White; text-align:center;'>" +
                                                   "<td><b> Ticket </b></td>" +
                                                   "<td><b>Active</b></td>" +
                                                   "<td><b>Work In Progress</b></td>" +
                                                   "<td><b>Awaiting</b></td>" +
                                                   "<td><b>Closed/Resolved</b></td>" +
                                                "</tr>" +
                                                "<tr style = 'color:red; text-align:center;'>" +
                                                   "<td>Incident</td>" +
                                                   "<td>" + activeTickets.ToString() + "</td>" +
                                                   "<td>" + workInprogress.ToString() + "</td>" +
                                                   "<td>" + awaiting.ToString() + "</td>" +
                                                   "<td>" + closedResolved.ToString() + "</td>" +
                                                "</tr>" +
                                                "<tr style = 'color:red; text-align:center;'>" +
                                                   "<td>RITM</td>" +
                                                   "<td>" + TaskClass.taskactiveTickets.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskworkInprogress.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskawaiting.ToString() + "</td>" +
                                                   "<td>" + TaskClass.taskclosedResolved.ToString() + "</td>" +
                                                "</tr>" +
                                             "</table><br><br>";
        }

        public void CombineTables (DataTable incTable, DataTable taskTable)
        {
            string[] row = new string[taskTable.Columns.Count];
            foreach (DataRow trow in taskTable.Rows)
            {
                for (int i = 0; i < taskTable.Columns.Count; i++)
                {
                    row[i] = trow.ItemArray.GetValue(i).ToString();
                }
                incTable.Rows.Add(row);
            }
        }

        public static DataTable SortTables(DataTable dt)
        {
            dt.DefaultView.Sort = "Updated DESC";
            dt = dt.DefaultView.ToTable();
            dt.DefaultView.Sort = "Priority";
            dt = dt.DefaultView.ToTable();
            return dt;
        }

        public static DateTime ConvertDates(string rowData)
        {
            double temp = Convert.ToDouble(rowData);
            DateTime convertedDate = DateTime.FromOADate(temp);
            return convertedDate;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Path.Combine(Directory.GetCurrentDirectory(), "Read Me.txt"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //form2.ShowDialog();
            System.Diagnostics.Process.Start(Path.Combine(Directory.GetCurrentDirectory(), "APM Members.txt"));
        }
        private void ResetData()
        {
            emailItem = "";
            tableEmailHead = "";
            openTickets = 0;
            closedResolved = 0;
            activeTickets = 0;
            workInprogress = 0;
            awaiting = 0;
            totalTickets = 0;
            severetyLow = 0;
            severetyModerate = 0;
            severetyHigh = 0;
            severetyCritical = 0;
            countsMessage = "<div  style = 'font-size:9pt; font-family:calibri;'>" +
                                "Hi Team," +
                             "</div>" +
                             "<br>" +
                             "<div style = 'font-size:9pt; font-family:calibri;'>" +
                                 "Please see open snow tickets for today. Kindly revisit your ticket/s and close all that can be closed." +
                             "</div>" +
                             "<br><br>" +
                             "<p style = 'font-size:9pt; font-family:calibri;'>" +
                                 "<b>I. Tickets Summary</b>" +
                             "</p>";
        }
    }
}
