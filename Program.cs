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
        static string Url;// = "https://dev.azure.com/sagorg1";
        static string UserPAT;// = "44tfdkhh7t2yzztombfdbzisjs7laljwpo5sbhngbfeyr57e5pta";
        static string ProjectName;// = "HOLMES-TrainingStudio";
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static public string OldTeamProject= "HOLMES-TrainingStudio";

        static void Main(string[] args)
        {
            Console.WriteLine("Enter The Server Url(https://dev.azure.com/{Organisation}): ");
            Url = Console.ReadLine();
            Console.WriteLine("Enter The Personal Access Token: ");
            UserPAT = Console.ReadLine();
            Console.WriteLine("Enter The Project Nmae: ");
            ProjectName = Console.ReadLine();
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
                   // string ID = dr["ID"].ToString();                    
                        WorkitemFromExcel item = new WorkitemFromExcel();
                    //item.id = ID;
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
        public static void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            Dictionary<string, object> Fields ;
            List<string> newStates = new List<string>(){ "New", "To Do" };
            /*string Areapath;
            string iteration;*/
            foreach (var wi in WiList)
            {
                Fields = new Dictionary<string, object>();
                if (wi.parent != null)
                    WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");
               if (!newStates.Contains(wi.WiState.ToString()) )
                    Fields.Add("State", wi.WiState.ToString());
                /* Areapath = wi.AreaPath.ToString().Replace(OldTeamProject, ProjectName);
                 Fields.Add("System.AreaPath", Areapath);
                 iteration = wi.Itertation.ToString().Replace(OldTeamProject, ProjectName);
                 Fields.Add("System.IterationPath", iteration);
                 Fields.Add("System.TeamProject", ProjectName);*/
                 if(Fields.Count!=0)
                    WIOps.UpdateWorkItemFields(wi.id, Fields);
            }
        }
        public static ParentWorkItem getParentData(DataTable dt, int rowindex, int columnindex)
        {
            ParentWorkItem workItem = new ParentWorkItem();
            //bool hasParent;

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
                        else if (column.ToString() != "State")
                            fields.Add(column.ToString(), Dr[column.ToString()]);
                    

                    /* else if(column.ToString()=="State")
                 { }*/
                    /* else if(Dr[column.ToString()].ToString()=="Story")
                 {
                     string val = Dr[column.ToString()].ToString();
                     val.Replace("Story", "User Story");
                     fields.Add(column.ToString(), val);

                 }*/

                }

            }
            var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);
            return newWi.Id.Value;
        }

        public static DataTable ReadExcel()
        {
            DataTable Dt = new DataTable();

            try
            {
                Excel.Application xlApp = new Excel.Application();
                //Console.Write("Enter The Ecel File Path:");
                /*string ExcelPath=Console.ReadLine();*/
                string ExcelPath = @"C:\Users\vidyasagarp\Documents\Training Studio - Final - Template.xlsx";
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelPath);//@""+
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
                            row[ColName] = xlRange.Cells[j][i].Value.ToString();
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

