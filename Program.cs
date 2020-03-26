using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using WorkItemPublish.Class;
using Excel = Microsoft.Office.Interop.Excel;
namespace WorkItemPublish
{
    class Program
    {
        static string Url;//= "https://dev.azure.com/Organisation";
        static string UserPAT;
        static string ProjectName ;
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static public string OldTeamProject= "HOLMES-TrainingStudio";
        static string ExcelPath;
        static void Main(string[] args)
        {
            Console.WriteLine("Enter The Server Url(https://dev.azure.com/{Organisation}): ");
            Url = Console.ReadLine();
            Console.WriteLine("Enter The Personal Access Token: ");
            UserPAT = Console.ReadLine();
            Console.WriteLine("Enter The Project Nmae: ");
            ProjectName = Console.ReadLine();
            Console.WriteLine("Enter The Old Project Nmae: ");
            OldTeamProject = Console.ReadLine();
            Console.Write("Enter The Ecel File Path:");
            ExcelPath = Console.ReadLine();
            WIOps.ConnectWithPAT(Url, UserPAT);
            DT = ReadExcel();
            List<WorkitemFromExcel> WiList = GetWorkItems();
            CreateLinks(WiList);
            Console.WriteLine("Successfully Migrated WorkItems");
            Console.ReadLine();
        }
        /// <summary>
        /// Iterates The Datatable for Creating the Workitems.
        /// </summary>
        /// <returns> list of workItems</returns>
        public static List<WorkitemFromExcel> GetWorkItems()
        {
            List<WorkitemFromExcel> workitemlist = new List<WorkitemFromExcel>();
            if (DT.Rows.Count > 0)
            {
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    DataRow dr = DT.Rows[i];
                    WorkitemFromExcel item = new WorkitemFromExcel();
                    if (DT.Rows[i] != null)
                    {
                        item.id = createWorkItem(dr);
                        dr["ID"] = item.id.ToString();
                        item.WiState = dr["State"].ToString();
                        int columnindex = 0;
                        foreach (var col in TitleColumns)
                        {
                            if (!string.IsNullOrEmpty(col))
                            {
                                if (!string.IsNullOrEmpty(dr[col].ToString()))
                                {
                                    item.tittle = dr[col].ToString();
                                    if (i > 0 && columnindex > 0)
                                        item.parent = getParentData(DT, i - 1, columnindex);
                                    break;
                                }
                            }
                            columnindex++;
                        }
                        workitemlist.Add(item);
                    }
                    else
                    {

                    }
                    
                }
            }

            return workitemlist;
        }
        /// <summary>
        /// Updates The WorKItem With The Parent and the State
        /// </summary>
        /// <param name="WiList"></param>
        public static void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            Dictionary<string, object> Fields ;
            List<string> newStates = new List<string>(){ "New", "To Do" };
            foreach (var wi in WiList)
            {
                Fields = new Dictionary<string, object>();
                if (wi.parent != null)
                    WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");
                if (!newStates.Contains(wi.WiState.ToString()) )
                    Fields.Add("State", wi.WiState.ToString());           
                if (Fields.Count!=0)
                    WIOps.UpdateWorkItemFields(wi.id, Fields);
            }
        }

        /// <summary>
        /// Method to Get The Parent oF the WoirkItem
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="rowindex"></param>
        /// <param name="columnindex"></param>
        /// <returns> Parent WorkItem</returns>
        public static ParentWorkItem getParentData(DataTable dt, int rowindex, int columnindex)
        {
            ParentWorkItem workItem = new ParentWorkItem();
    

            if (columnindex > 0)
            {
                for (int i = rowindex; i >= 0; i--)
                {
                    
                    DataRow dr = dt.Rows[i];
                    int colindex = columnindex;
                    while (colindex > 0)
                    {
                        int index = colindex - 1;
                        if (!string.IsNullOrEmpty(dr[TitleColumns[index]].ToString()))
                        {
                            workItem.Id = int.Parse(dr["ID"].ToString());
                            workItem.tittle = dr[TitleColumns[index]].ToString();
                            break;
                        }
                        colindex--;
                    }
                    if (!string.IsNullOrEmpty(workItem.tittle))
                    { break; }          
                        
                }
            }
            return workItem;

        }
        /// <summary>
        /// Creates Fields and Values with the excel data And create Wis with that fields
        /// </summary>
        /// <param name="Dr"></param>
        /// <returns> ID of created WI</returns>
        static int createWorkItem(DataRow Dr)
        {
            
            Dictionary<string, object> fields = new Dictionary<string, object>();
            foreach (DataColumn column in DT.Columns)
            {
                if (Dr[column.ToString()].ToString() != "")
                {                    
                        if (column.ToString().StartsWith("Title"))
                            fields.Add("Title", Dr[column.ToString()]);
                        else if (column.ColumnName == "Iteration")
                            {
                                fields.Add("Iteration Path", Dr[column.ToString()]);
                            }
                        else if (column.ToString() != "State")
                            fields.Add(column.ToString(), Dr[column.ToString()]);
                }

            }
            var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(),fields);
            return newWi.Id.Value;
        }
        /// <summary>
        /// Method to read The Excel sheet
        /// </summary>
        /// <returns> Data Table</returns>
        public static DataTable ReadExcel()
        {
            DataTable Dt = new DataTable();

            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"" + ExcelPath);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                DataRow row;

                string ColName = "";
                for (int i = 1; i <= rowCount; i++)
                {
                    row = Dt.NewRow();
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (i == 1)
                        {
                            ColName = xlRange.Cells[j][i].Value.ToString();
                            if (ColName.StartsWith("Title"))
                            {
                                TitleColumns.Add(ColName);
                            }
                            DataColumn column = new DataColumn(ColName);
                            Dt.Columns.Add(column);

                            continue;
                        }
                        ColName = xlRange.Cells[j][1].Value.ToString();
                        if (xlRange.Cells[j][i].Value != null)
                        {
                            string val = xlRange.Cells[j][i].Value.ToString().TrimStart('\\');
                            val=val.Replace(OldTeamProject, ProjectName);
                            if(ColName== "Iteration")
                            { }
                            row[ColName] = val;

                        }
                    }
                    if (i != 1)
                        Dt.Rows.Add(row);
                    /*string teststring =row.ItemArray[3].ToString();*/
                }

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return Dt;
        }
       
    }

}

