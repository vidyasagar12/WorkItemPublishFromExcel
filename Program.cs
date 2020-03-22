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
        static string Url = null;
        static string UserPAT = null;
        static string ProjectName = null;
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();

        static void Main(string[] args)
        {
            Console.WriteLine("Enter The Server Url(https://{Instance Name}/{Organisation}): ");
            Url = Console.ReadLine();
            Console.WriteLine("Enter The Personal Access Token: ");
            UserPAT = Console.ReadLine();
            Console.WriteLine("Enter The Project Nmae: ");
            ProjectName = Console.ReadLine();
            WIOps.ConnectWithPAT(Url, UserPAT);
            DT = ReadExcel();
            List<WorkitemFromExcel> WiList = GetWorkItems();
            CreateLinks(WiList);
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
                        dr["ID"] = item.id.ToString();
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
            foreach (var wi in WiList)
            {
                if (wi.parent != null)
                    WIOps.UpdateWorkItem(wi.parent.Id, wi.id, "");
            }
        }
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

        public static List<string> inavlidCoumns = new List<string>();
        static int createWorkItem(DataRow Dr)
        {
            inavlidCoumns.AddRange(new string[] { "ID", "Team Project", "Area Path", "Iteration", "State" });
            Dictionary<string, object> fields = new Dictionary<string, object>();

            foreach (DataColumn column in DT.Columns)
            {
                if (Dr[column.ToString()].ToString() != "")
                {
                    if (!inavlidCoumns.Contains(column.ToString()))
                    {
                        if (column.ToString().StartsWith("Title"))
                            fields.Add("Title", Dr[column.ToString()]);
                        else
                            fields.Add(column.ToString(), Dr[column.ToString()]);
                    }
                }

            }
            var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);
            return newWi.Id.Value;
        }

        public static DataTable ReadExcel()
        {
            Excel.Application xlApp = new Excel.Application();
            Console.Write("Enter The Ecel File Path:");
            string ExcelPath=Console.ReadLine();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@""+ExcelPath);
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

