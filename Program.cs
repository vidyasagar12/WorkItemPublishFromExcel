using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
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
        static string Url;
        static string UserPAT;
        static string ProjectName;
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static public string OldTeamProject;
        static public string ExcelUniqueField;
        public static Mapper mapperObject;
        static string ExcelPath = "";
        static string CreatedlogFilePath = "";
        static string UpdatedlogFilePath = "";
        static string MapFilePath = "";
        static string createdWIsstring;
        static string updatedWIsstring;
        static Dictionary<string, string> CreatedWIS= new Dictionary<string, string>();
        static Dictionary<string, string> UpdatedWIS= new Dictionary<string, string>();
        static Dictionary<String, string> FieldMapper = new Dictionary<string, string>();
        static Dictionary<string, String> DataMapper = new Dictionary<string, string>();
        static List<string> UpdateAfterCreating = new List<string>() { "ID" };
        static List<string> RequiredFields = new List<string>();
        static string CurrentWorkItem = "";

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Enter Azure DevOps Organisation Name: ");
                Url = "https://dev.azure.com/" + Console.ReadLine();
                Console.WriteLine("Enter Personal Access Token: ");
                UserPAT = Console.ReadLine();                
                do
                {
                    Console.Write("Enter Excel file path:");
                    ExcelPath = Console.ReadLine();                    

                } while (!File.Exists(ExcelPath));
                CheckLogFile();
                do
                {
                    Console.Write("Enter The Map File Path:");
                    MapFilePath = Console.ReadLine();                    
                } while (!File.Exists(MapFilePath)&&MapFilePath.ToLower().EndsWith(".json"));
                string MapFileContent = File.ReadAllText(MapFilePath);
                mapperObject = JsonConvert.DeserializeObject<Mapper>(MapFileContent);
                OldTeamProject = mapperObject.source_project;
                ProjectName = mapperObject.target_project;
                ExcelUniqueField = mapperObject.Unique_ID_Field;
                MapperMethod();
                DT = ReadExcel();
                WIOps.ConnectWithPAT(Url, UserPAT);
                List<WorkitemFromExcel> WiList = GetWorkItems();
                CreateLinks(WiList);
                UpdateWIFields();
                Console.WriteLine("Successfully Migrated WorkItems");               
            }
            catch(Exception E)
            {
                Console.WriteLine("Exception Message:"+E.Message);
                Console.WriteLine(" Inner Exception Message:"+E.InnerException);
            }
            finally
            {
                if(!string.IsNullOrEmpty(createdWIsstring))
                File.AppendAllText(CreatedlogFilePath, createdWIsstring);
                if(!string.IsNullOrEmpty(updatedWIsstring))
                File.AppendAllText(UpdatedlogFilePath, updatedWIsstring);
                Console.WriteLine("Please Enter Any Key To Exit");
                Console.ReadLine();

            }
        }
        static void CheckLogFile()
        {
            try
            {
                String[] SplitPath = ExcelPath.Split('.');
                CreatedlogFilePath = SplitPath[0] + "-CreatedWIs.txt";
                UpdatedlogFilePath = SplitPath[0] + "-UpdatedWIs.txt";
                if (!File.Exists(CreatedlogFilePath) || !File.Exists(UpdatedlogFilePath))
                {
                    if (!File.Exists(CreatedlogFilePath))
                    {
                        File.Create(CreatedlogFilePath);
                    }
                    else if (!File.Exists(UpdatedlogFilePath))
                    {
                        File.Create(UpdatedlogFilePath);
                    }
                }
                else
                {
                    String logFileContent = File.ReadAllText(CreatedlogFilePath);
                    ReadLogFile(logFileContent, CreatedWIS);
                    logFileContent = File.ReadAllText(UpdatedlogFilePath);
                    ReadLogFile(logFileContent, UpdatedWIS);
                }
            }
            catch(Exception E)
            {
                Console.WriteLine("Error Occured While Reading Log File");
                throw E;
            }
        }
        public static void ReadLogFile(string Content,Dictionary<string,string> dictionary)
        {
            try
            {
                if (!string.IsNullOrEmpty(Content))
                {
                    string[] logs = Content.Split(',');
                    foreach (string log in logs)
                    {
                        if (!string.IsNullOrEmpty(log))
                        {
                            string[] Ids = log.Split('-');
                            if (!dictionary.ContainsValue(Ids[1]))
                                dictionary.Add(Ids[0], Ids[1]);
                        }
                    }
                }
            }
            catch(Exception E)
            {
                throw E;
            }
        }
        
        public static void MapperMethod()
        {
            try
            {
                //For Mapping Columns To Azure Devops Fields
                foreach(Field field in mapperObject.field_map.field)
                {
                    if(field.ShouldUpdateAfterCreating!=null)
                        if (field.ShouldUpdateAfterCreating.ToLower() == "true")
                            UpdateAfterCreating.Add(field.source);
                    if(!FieldMapper.ContainsKey(field.source))
                    FieldMapper.Add(field.source, field.target);
                    if (field.mapping != null)
                    {
                        foreach (var data in field.mapping.values)
                        {
                            if(!DataMapper.ContainsKey(data.source))
                                DataMapper.Add(data.source, data.target);
                        }
                    }
                }
               
            }
            catch(Exception E)
            {
                Console.WriteLine("Error Occured While Mapping Fields");
                throw (E);
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
                            if (CreatedWIS.ContainsKey(dr[ExcelUniqueField].ToString()))
                            {
                                item.id = int.Parse(CreatedWIS[dr[ExcelUniqueField].ToString()]);                               
                            }
                            else
                            {
                                item.id = createWorkItem(dr);
                                Console.WriteLine("WorkItem Created With ID= " + item.id);
                                createdWIsstring += dr[ExcelUniqueField] + "-" + item.id + ",";
                            }
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
                Console.WriteLine("Error Occured While Iterating The Data");
                throw E;                
            }

        }
        public static void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            try
            {
                if (WiList == null)
                    Console.WriteLine("No WorkItems Were Created To Updated Links");
                foreach (var wi in WiList)
                {
                    if (!UpdatedWIS.ContainsValue(wi.id.ToString()))
                    {
                        CurrentWorkItem = wi.tittle;
                        Console.WriteLine("Updating Links of" + wi.id);
                        if (wi.parent != null)
                            WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");
                    }
                }
            }
            catch(Exception E)
            {
                Console.WriteLine("Error Occured While Creating Parent-Child Relations For"+CurrentWorkItem);
                throw E;
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
                Console.WriteLine("Error Occured While Finding Parent Of WorkItems");
                throw E;
            }

        }
        static int createWorkItem(DataRow Dr)
        {
            string WIType="";

            try
            {
                Dictionary<string, object> fields = new Dictionary<string, object>();

                foreach (DataColumn column in DT.Columns)
                {
                    string ColumnName = column.ColumnName;
                    string value = Dr[ColumnName].ToString();                                       
                    /*if (value.StartsWith(@"\"))
                    {
                        if (value == @"\")
                            value = ProjectName;
                        value = value.TrimStart('\\');
                    }
                    if (value.StartsWith(OldTeamProject))
                        value = value.Replace(OldTeamProject, ProjectName);*/
                    if (!string.IsNullOrEmpty(value) && !UpdateAfterCreating.Contains(ColumnName))
                    {
                        if (ColumnName == mapperObject.WorkItem_Type_Field)
                        {                            
                                if (DataMapper.ContainsKey(value))
                                    WIType = DataMapper[value];
                                else
                                    WIType = value;
                            continue;
                        }
                        if (FieldMapper.ContainsKey(ColumnName))
                        {
                            if (ColumnName.StartsWith("Title"))
                            {
                                CurrentWorkItem = value;
                                fields.Add("Title", value);
                                continue;
                            }
                            if (DataMapper.ContainsKey(value))
                                fields.Add(FieldMapper[ColumnName], DataMapper[value]);
                            else
                                fields.Add(FieldMapper[ColumnName], value);
                        }                        
                        else
                        {
                            if (ColumnName.StartsWith("Title"))
                            {
                                fields.Add("Title", value);
                                continue;
                            }
                            if (DataMapper.ContainsKey(value))
                                fields.Add(column.ColumnName, DataMapper[value]);
                            else
                                fields.Add(column.ColumnName, value);
                        }

                    }

                }
                var newWi = WIOps.CreateWorkItem(ProjectName, WIType, fields);                
                    
                    return newWi.Id.Value;
            }
            catch(Exception E)
            {
                Console.WriteLine("Error  Occured While Creatig WorkItem"+CurrentWorkItem);
                throw E;                
            }
        }
        public static void UpdateWIFields()
        {
            try
            {       

                foreach (DataRow row in DT.Rows)
                {
                    if (!UpdatedWIS.ContainsKey(row[ExcelUniqueField].ToString()))
                    {
                        CurrentWorkItem = row["ID"].ToString();
                        Console.WriteLine("Updating Fields of" +CurrentWorkItem);
                        Dictionary<string, object> Updatefields = new Dictionary<string, object>();
                        foreach (DataColumn col in DT.Columns)
                        {
                            if (!string.IsNullOrEmpty(row[col].ToString()))
                            {
                                if (col.ToString() != "ID" && UpdateAfterCreating.Contains(col.ColumnName))
                                {
                                    string val = row[col.ToString()].ToString();//.TrimStart('\\');
                                    if (!string.IsNullOrEmpty(val))
                                        Updatefields.Add(col.ToString(), val);
                                }
                            }
                        }
                        WorkItem UpdatedWI = WIOps.UpdateWorkItemFields(int.Parse(row["ID"].ToString()), Updatefields);
                        if (UpdatedWI != null)
                        {
                            updatedWIsstring += row[ExcelUniqueField] + "-" + UpdatedWI.Id.ToString() + ",";
                        }
                    }
                }
            }
            catch(Exception E)
            {
                Console.WriteLine("Error Occured While Updating WorkItem With ID="+CurrentWorkItem);
                throw E;
            }
           
        }

        public static DataTable ReadExcel()
        {
            DataTable Dt = new DataTable();

            try
            {
                Excel.Application xlApp = new Excel.Application();                
               
                while(!System.IO.File.Exists(ExcelPath)||(ExcelPath.EndsWith("xls")&&ExcelPath.EndsWith("xlsx")))
                {
                    Console.WriteLine("File Not Found Or File Is not in supported format Please Enter A valid Path");
                    ExcelPath = Console.ReadLine();
                }
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelPath);
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
                    if(!Dt.Columns.Contains("ID"))
                    {
                        DataColumn column = new DataColumn("ID");
                        Dt.Columns.Add("ID");
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
                Console.WriteLine("Error Occured While Reading Excel File");
                throw ex;
            }
            return Dt;
        }

    }

}

