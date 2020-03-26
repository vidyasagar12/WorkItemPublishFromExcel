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
        static string Url = "https://dev.azure.com/sagorg1";
        static string UserPAT = "44tfdkhh7t2yzztombfdbzisjs7laljwpo5sbhngbfeyr57e5pta";
        static string ProjectName = "AttacmentImport";
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static public string OldTeamProject;

        static void Main(string[] args)
        {
           /* Console.WriteLine("Enter The Server Url(https://{Instance Name}/{Organisation}): ");
            Url = Console.ReadLine();
            Console.WriteLine("Enter The Personal Access Token: ");
            UserPAT = Console.ReadLine();
            Console.WriteLine("Enter The Project Nmae: ");
            ProjectName = Console.ReadLine();*/
            WIOps.ConnectWithPAT(Url, UserPAT);
            DT = ReadExcel();
            List<WorkitemFromExcel> WiList = GetWorkItems();
            CreateLinks(WiList);
            Console.WriteLine("Successfully Migrated WorkItems");
            Console.ReadLine();
        }
        public static List<WorkitemFromExcel> GetWorkItems()
        {
            List<WorkitemFromExcel> workitemlist = new List<WorkitemFromExcel>();
            if (DT.Rows.Count > 0)
            {
                for (int i = 0; i < DT.Rows.Count; i++)
                {
                    DataRow dr = DT.Rows[i];
                    string ID = dr["ID"].ToString();
                    if (!string.IsNullOrEmpty(ID))
                    {
                        WorkitemFromExcel item = new WorkitemFromExcel();
                        //item.id = ID;
                        item.id = createWorkItem(dr);
                        item.Old_ID= int.Parse(dr["ID"].ToString());
                        dr["ID"] = item.id.ToString();
                        item.WiState = dr["State"].ToString();
                        item.AreaPath = dr["Area Path"].ToString();
                        item.Itertation = dr["Iteration"].ToString();
                        OldTeamProject = dr["Team Project"].ToString();
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
                }
            }

            return workitemlist;
        }
        public static void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            Dictionary<string, object> Fields ;
            List<string> newStates = new List<string>(){ "New", "To Do" };
            string Areapath;
            string iteration;
            foreach (var wi in WiList)
            {
                Fields = new Dictionary<string, object>();
                if (wi.parent != null)
                    WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");
                if (!newStates.Contains(wi.WiState.ToString()) )
                    Fields.Add("State", wi.WiState.ToString());
                Areapath = wi.AreaPath.ToString().Replace(OldTeamProject, ProjectName);
                Fields.Add("System.AreaPath", Areapath);
                iteration = wi.Itertation.ToString().Replace(OldTeamProject, ProjectName);
                Fields.Add("System.IterationPath", iteration);
                Fields.Add("System.TeamProject", ProjectName);
                Fields.Add("Old_ID", wi.Old_ID);
                WIOps.UpdateWorkItemFields(wi.id, Fields);
            }
        }
        public static ParentWorkItem getParentData(DataTable dt, int rowindex, int columnindex)
        {
            ParentWorkItem workItem = new ParentWorkItem();
           // bool hasParent;

            if (columnindex > 0)
            {
                for (int i = rowindex; i >= 0; i--)
                {
                   // hasParent = false;
                    DataRow dr = dt.Rows[i];
                    int colindex = columnindex;
                    while (colindex > 0)
                    {
                        int index = colindex - 1;
                        if (!string.IsNullOrEmpty(dr[TitleColumns[index]].ToString()))
                        {
                            //hasParent = true;
                            workItem.Id = int.Parse(dr["ID"].ToString());
                            workItem.tittle = dr[TitleColumns[index]].ToString();
                            break;
                        }
                        colindex--;
                    }
                    if (!string.IsNullOrEmpty(workItem.tittle))
                    { break; }
                    /*if (hasParent == false)
                        return null;*/
                        
                }
            }
            return workItem;

        }

        public static List<string> inavlidCoumns = new List<string>();
        static int createWorkItem(DataRow Dr)
        {
            Dictionary<string, object> fields = new Dictionary<string, object>();
            foreach (DataColumn column in DT.Columns)
            {
                if (Dr[column.ToString()].ToString() != "")
                {
                    if (column.ToString().StartsWith("Title"))
                        fields.Add("Title", Dr[column.ToString()]);
                    /*if (column.ToString()== "Work Item Type")
                    {          
                        fields.Add(column.ToString(), Dr[column.ToString()]);
                    }*/
                }
                if (fields.Count != 0)
                    break;
            }
            WorkItem newWi=new WorkItem();
            if (fields.Count != 0)
            {
                newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);
            }
            return newWi.Id.Value;
        }

        public static DataTable ReadExcel()
        {
            Excel.Application xlApp = new Excel.Application();
            //Console.Write("Enter The Ecel File Path:");
            /*string ExcelPath=Console.ReadLine();*/           
            string ExcelPath = @"C:\Users\vidyasagarp\Documents\naveenkunder-SM-Epic18-03-2020 10_54_40.xlsx";
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelPath);//@""+
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            DataTable Dt = new DataTable();
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
                        row[ColName] = xlRange.Cells[j][i].Value.ToString();
                }
                if (i != 1)
                    Dt.Rows.Add(row);
                /*string teststring =row.ItemArray[3].ToString();*/
            }
            return Dt;
        }
       
    }

}

