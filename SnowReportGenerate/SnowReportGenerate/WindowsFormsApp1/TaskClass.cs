using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;
using WindowsFormsApp1;

namespace WindowsFormsApp1
{
    class TaskClass
    {

        public static DataTable dtTaskUnAssigned = new DataTable();
        public static DataTable dtTaskApm = new DataTable();
        public static DataTable dtTaskOps = new DataTable();
        public static DataTable dtTaskApmEm = new DataTable();
        public static DataTable dtTaskOpsEm = new DataTable();

        public static DateTime toDay = DateTime.Now;
        public static DateTime updatedDate;
        public static DateTime reportedDate;
        public static DateTime closedDate;

        public static int taskopenTickets = 0;
        public static int taskclosedResolved = 0;
        public static int taskactiveTickets = 0;
        public static int taskworkInprogress = 0;
        public static int taskawaiting = 0;
        public static int tasktotalTickets = 0;
        public static int taskLow = 0;
        public static int taskModerate = 0;
        public static int taskHigh = 0;
        public static int taskCritical = 0;


        public static void GetTaskData(string source)
        {
            ResetTicketCounts();

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(source)))
            {

                int dtcoll = 0;
                int heading = 0;
                dtTaskApm.Clear();
                dtTaskOps.Clear();
                dtTaskApmEm.Clear();
                dtTaskOpsEm.Clear();
                dtTaskUnAssigned.Clear();
                var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                string[] row = new string[totalColumns];

                for (int rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                {
                    //get values per cell, data is the values from one cell
                    var cell = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                    foreach (var data in cell)
                    {
                        //creates headings for data tables, first row in excel will be the headings.
                        if (heading != 1)
                        {
                            if (dtTaskApm.Columns.Count != totalColumns)
                            {
                                dtTaskApm.Columns.Add(data);
                                dtTaskOps.Columns.Add(data);
                                dtTaskApmEm.Columns.Add(data);
                                dtTaskOpsEm.Columns.Add(data);
                                dtTaskUnAssigned.Columns.Add(data);
                            }
                            else { heading = 1; }
                        }
                        if (heading == 1)
                        {
                            //Create row to add, dtcol is numbers of columns to be fill need 12 column to create 1 row
                            if (dtcoll != 12)
                            {
                                row[dtcoll] = data;
                                dtcoll++;
                            }
                            //Add the row created to data tables
                            else
                            {
                                try
                                {
                                    //Convert excel Updated Date 
                                    Console.WriteLine(row[7]);
                                    updatedDate = Form1.ConvertDates(row[7]);
                                    row[7] = updatedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                }
                                catch (Exception ex)
                                {
                                    updatedDate = new DateTime(1990, 01, 01);
                                    ex.ToString();
                                }
                                try
                                {
                                    //Convert excel reported Date
                                    Console.WriteLine(row[5]);
                                    reportedDate = Form1.ConvertDates(row[5]);
                                    row[5] = reportedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                }
                                catch (Exception ex)
                                {
                                    reportedDate = new DateTime(1990, 01, 01);
                                    ex.ToString();
                                }
                                try
                                {
                                    //Convert excel Closed Date
                                    Console.WriteLine(row[11]);
                                    closedDate = Form1.ConvertDates(row[11]);
                                    row[11] = closedDate.ToString("MM/dd/yyyy hh:mm:ss");
                                }
                                catch (Exception ex)
                                {
                                    closedDate = new DateTime(1990, 01, 01);
                                    ex.ToString();
                                }

                                //if ((Form1.lastDate.Date <= closedDate.Date || row[11] == null || row[11] == "") && (row[10] != "" || row[10] != null) && row[2] != "Cancelled")
                                if (Valid(Form1.lastDate, closedDate, row[11], row[10], row[2])) // Determines if button used is am or pm and if the ticket is have a valid dates
                                {
                                    Console.WriteLine(Form1.lastDate.Date.ToString("MM/dd/yyyy") + closedDate.Date + "True !!");
                                    CountTickets(row[2]); // Count Status of the ticket
                                    CountPriority(row[4]); // Count the number of priorities
                                    
                                    //Populating the Unassigned data Table
                                    if ((row[3] == "" || row[3] == null) && (row[2] != "Closed Complete" || row[2] != "Resolved"))
                                    { dtTaskUnAssigned.Rows.Add(row); }

                                    //Populating the APM data Table
                                    else if (Form1.ApmList.Contains(row[3].ToString()))
                                    {

                                        dtTaskApm.Rows.Add(row);
                                        if (row[2] == "Closed Complete" || row[2] == "Resolved") { }
                                        else { dtTaskApmEm.Rows.Add(row); }

                                    }
                                    //Populating the Ops data Table
                                    else
                                    {
                                        dtTaskOps.Rows.Add(row);
                                        if (row[2] == "Closed Complete" || row[2] == "Resolved") { }
                                        else { dtTaskOpsEm.Rows.Add(row); }

                                    }
                                }
                                dtcoll = 0;
                            }
                        }
                    }

                }
            }

            //Sort table Updated DESC then sort to Priority
            dtTaskApm = Form1.SortTables(dtTaskApm);
            dtTaskApmEm = Form1.SortTables(dtTaskApmEm);
            dtTaskOps = Form1.SortTables(dtTaskOps);
            dtTaskOpsEm = Form1.SortTables(dtTaskOpsEm);
            dtTaskUnAssigned = Form1.SortTables(dtTaskUnAssigned);
        }

        public static bool Valid (DateTime lastdate, DateTime closeddate, string closedate, string reference, string status)
        {
            if (Form1.btnUsed == " am")
            {
                if ((lastdate.Date <= closedDate.Date ||
                closedate == null || closedate == "") &&
                (reference != "" || reference != null) &&
                status != "Cancelled")
                { return true; }
                else { return false; }
            }
            else 
            {
                closeddate = toDay;
                if ((lastdate.Date == closedDate.Date ||
                closedate == null || closedate == "") &&
                (reference != "" || reference != null) &&
                status != "Cancelled")
                { return true; }
                else { return false; }
            }
            

        }

        public static void CountTickets(string status)
        {
            switch (status)
            {
                case "Resolved": { taskclosedResolved++; break; }
                case "Closed Complete": { taskclosedResolved++; break; }
                case "Open": { taskactiveTickets++; break; }
                case "Work in Progress": { taskworkInprogress++; break; }
                case "For Review": { taskworkInprogress++; break; }
                case "Pending": { taskworkInprogress++; break; }
                case "Additional Information Requested": { taskawaiting++; break; }
            }
            taskopenTickets = taskworkInprogress + taskawaiting + taskactiveTickets;
            tasktotalTickets = taskopenTickets + taskclosedResolved;
            
        }

        public static void CountPriority(string priority)
        {
            switch (priority)
            {
                case "4 - Low": { taskLow++; break; }
                case "3 - Moderate": { taskModerate++; break; }
                case "2 - High": { taskHigh++; break; }
                default: { taskCritical++; break; }
            }
        }
        
        public static void ResetTicketCounts()
        {
            taskopenTickets = 0;
            taskclosedResolved = 0;
            taskactiveTickets = 0;
            taskworkInprogress = 0;
            taskawaiting = 0;
            tasktotalTickets = 0;
            taskLow = 0;
            taskModerate = 0;
            taskHigh = 0;
            taskCritical = 0;
        }
    }

}
