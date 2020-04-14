using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using WorkItemPublish.Class;
using WorkItemPublish.Model.Mapper;
using Excel = Microsoft.Office.Interop.Excel;
namespace WorkItemPublish
{
    class Program
    {
        static string Url;//= "https://dev.azure.com/sagorg1";
        static string UserPAT;// = "qbi2it66pkjvlj7p4whh7efbkdjqzemume5xazf7ogspqmcieosa";
        static string ProjectName;// = "Agile Project";//"HOLMES-TrainingStudio";
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static public string OldTeamProject;// = "HOLMES-AutomationStudio";

        static void Main(string[] args)
        {
            try
            {

                /*Console.WriteLine("Enter Azure DevOps Organisation Name: ");
                Url = "https://dev.azure.com/"+Console.ReadLine();
                Console.WriteLine("Enter Personal Access Token: ");
                UserPAT = Console.ReadLine();
                Console.WriteLine("Enter Source Project name: ");
                OldTeamProject = Console.ReadLine();
                while(string.IsNullOrEmpty(OldTeamProject))
                {
                    Console.WriteLine("Please Enter the Source Project Name");
                    OldTeamProject = Console.ReadLine();
                }
                Console.WriteLine("Enter Azure DevOps Destination Project name: ");
                ProjectName = Console.ReadLine();
                while (string.IsNullOrEmpty(ProjectName))
                {
                    Console.WriteLine("Please Enter the Destination Project Name");
                    ProjectName = Console.ReadLine();
                }*/
                Console.Write("Enter The Map File Path:");
                string MapFilePath = Console.ReadLine();
                string MapFileContent = File.ReadAllText(MapFilePath);
                Mapper mapperObject = JsonConvert.DeserializeObject<Mapper>(MapFileContent);
                WIOps.ConnectWithPAT(Url, UserPAT);
                DT = ReadExcel();
                List<WorkitemFromExcel> WiList = GetWorkItems();
                CreateLinks(WiList);
                UpdateWIFields();
                Console.WriteLine("Successfully Migrated WorkItems");
                Console.ReadLine();
            }
            catch(Exception E)
            {
                Console.WriteLine(E.Message);
            }
        }
        public static List<WorkitemFromExcel> GetWorkItems()
        {
            try
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
                            Console.WriteLine("WorkItemPublish Created= " + item.id);
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
            catch (Exception E)
            {
                Console.WriteLine(E.Message);
                return null;
            }

        }
        public static void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            if (WiList == null)
                Console.WriteLine();
            
           foreach (var wi in WiList)
            {
                
                Console.WriteLine("Updating Links of" + wi.id);
                if (wi.parent != null)
                    WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");
               
            }
        }
        public static ParentWorkItem getParentData(DataTable dt, int rowindex, int columnindex)
        {
            try
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
            catch(Exception E)
            {
                Console.WriteLine(E.Message);
                return null;
            }

        }

    
        static int createWorkItem(DataRow Dr)
        {

            Dictionary<string, object> fields = new Dictionary<string, object>();

            foreach (DataColumn column in DT.Columns)
            {
                if (!string.IsNullOrEmpty(Dr[column].ToString()))
                {
                    if (column.ToString().StartsWith("Title"))
                        fields.Add("Title", Dr[column].ToString());                                                                 

                }

            }
            var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);
            if (newWi == null)
            {
                Console.WriteLine("Failed to create WorkItem please check The Inputs And try Again");
                return -1;
            }
            else
            return newWi.Id.Value;
        }
        public static void UpdateWIFields()
        {
            try
            {             

                foreach (DataRow row in DT.Rows)
                {
                    Console.WriteLine("Updating Fields of" + row["ID"]);
                    Dictionary<string, object> Updatefields = new Dictionary<string, object>();
                    foreach (DataColumn col in DT.Columns)
                    {
                        if (!string.IsNullOrEmpty(row[col].ToString()))
                        {
                            if (col.ToString() != "ID"&& col.ToString() != "Reason" && col.ToString() != "Work Item Type" && !col.ToString().StartsWith("Title"))
                            {
                                string val = row
                                    [col.ToString()].ToString().Replace(OldTeamProject, ProjectName).TrimStart('\\');
                                if (!string.IsNullOrEmpty(val))
                                    Updatefields.Add(col.ToString(), val);
                            }
                        }
                    }
                    WIOps.UpdateWorkItemFields(int.Parse(row["ID"].ToString()), Updatefields);
                }
            }catch(Exception E)
            {
                Console.WriteLine(E.Message);
            }

        }

        public static DataTable ReadExcel()
        {
            DataTable Dt = new DataTable();

            try
            {
                Excel.Application xlApp = new Excel.Application();
                Console.Write("Enter file path:");
                string ExcelPath = Console.ReadLine();
                //ExcelPath = @"C:\Users\naveenak\Downloads\Copy of SampleWorkItems-AgileProj.xlsx";
                while(!System.IO.File.Exists(ExcelPath)||(ExcelPath.EndsWith("xls")&&ExcelPath.EndsWith("xlsx")))
                {
                    Console.WriteLine("File Not Found Or File Is not in supported format Please Enter A valid Path");
                    ExcelPath = Console.ReadLine();
                }
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
                    {
                        Console.WriteLine("Reading excel data row " + i+"/"+ rowCount);
                        Dt.Rows.Add(row);
                    }
                    
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return Dt;
        }

    }

}

