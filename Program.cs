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
        public static Mapper mapperObject;
        static string ExcelPath = "";
        static string MapFilePath = "";
        static Dictionary<string, FieldsMapper> FieldMapper = new Dictionary<string, FieldsMapper>();

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
                do
                {
                    Console.Write("Enter The Map File Path:");
                    MapFilePath = Console.ReadLine();
                } while (!File.Exists(MapFilePath) && MapFilePath.ToLower().EndsWith(".json"));
                string MapFileContent = File.ReadAllText(MapFilePath);
                mapperObject = JsonConvert.DeserializeObject<Mapper>(MapFileContent);
                OldTeamProject = mapperObject.source_project;
                ProjectName = mapperObject.target_project;

                MapperMethod();
                DT = ReadExcel();
                WIOps.ConnectWithPAT(Url, UserPAT);
                List<WorkitemFromExcel> WiList = GetWorkItems();
                CreateLinks();
                UpdateWIFields();
                Console.WriteLine("Successfully Migrated WorkItems");
            }
            catch (Exception E)
            {
                Console.WriteLine("Exception Message:" + E.Message);
                Console.WriteLine(" Inner Exception Message:" + E.InnerException);
            }
            finally
            {
                ExportDataSetToExcel();
                Console.WriteLine("Please Enter Any Key To Exit");
                Console.ReadLine();

            }
        }

        public static void MapperMethod()
        {
            try
            {
                FieldsMapper FMObject;
                //For Mapping Columns To Azure Devops Fields
                foreach (Field field in mapperObject.field_map.field)
                {
                    if (field.ShouldUpdateAfterCreating != null)
                        if (field.ShouldUpdateAfterCreating.ToLower() == "true")
                            UpdateAfterCreating.Add(field.source);
                    if (!FieldMapper.ContainsKey(field.source))
                    {
                        FMObject = new FieldsMapper();
                        FMObject.TargetFieldName = field.target;
                        if (field.mapping != null)
                        {
                            foreach (var data in field.mapping.values)
                            {
                                FMObject.FieldSupprotedValues = new Dictionary<string, string>();
                                if (!FMObject.FieldSupprotedValues.ContainsKey(data.source))
                                {
                                    FMObject.FieldSupprotedValues.Add(data.source, data.target);
                                }
                            }
                        }
                        FieldMapper.Add(field.source, FMObject);
                    }
                }

            }
            catch (Exception E)
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
                            if (!string.IsNullOrEmpty(dr["ID"].ToString()))
                            {
                                item.id = int.Parse(dr["ID"].ToString());
                            }
                            else
                            {
                                item.id = createWorkItem(dr);
                                Console.WriteLine("WorkItem Created With ID= " + item.id);
                                dr["ID"] = item.id.ToString();
                            }

                            int columnindex = 0;
                            foreach (var col in TitleColumns)
                            {
                                if (!string.IsNullOrEmpty(col))
                                {
                                    if (!string.IsNullOrEmpty(dr[col].ToString()))
                                    {
                                        item.tittle = dr[col].ToString();
                                        if (i > 0 && columnindex > 0)
                                        {
                                            item.parent = getParentData(DT, i - 1, columnindex);
                                            dr["ParentID"] = item.parent.Id;
                                        }
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
        public static void CreateLinks()
        {
            try
            {
                foreach (DataRow row in DT.Rows)
                {
                    if (row["Updated"].ToString().ToLower() != "true")
                    {
                        if (!string.IsNullOrEmpty(row["ParentID"].ToString()))
                        {
                            CurrentWorkItem = row["ID"].ToString();
                            Console.WriteLine("Updating Links of" + CurrentWorkItem);
                            WIOps.UpdateWorkItemLink(int.Parse(row["ParentID"].ToString()), int.Parse(row["ID"].ToString()), "");
                        }
                    }
                }
            }
            catch (Exception E)
            {
                Console.WriteLine("Error Occured While Creating Parent-Child Relations For" + CurrentWorkItem);
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
            catch (Exception E)
            {
                Console.WriteLine("Error Occured While Finding Parent Of WorkItems");
                throw E;
            }

        }
        static int createWorkItem(DataRow Dr)
        {
            string WIType = "";

            try
            {
                Dictionary<string, object> fields = new Dictionary<string, object>();

                foreach (DataColumn column in DT.Columns)
                {
                    string ColumnName = column.ColumnName;
                    string value = Dr[ColumnName].ToString();
                    if (!string.IsNullOrEmpty(value) && !UpdateAfterCreating.Contains(ColumnName))
                    {
                        if (ColumnName == mapperObject.WorkItem_Type_Field)
                        {
                            if (FieldMapper[ColumnName].FieldSupprotedValues.ContainsKey(value))
                                WIType = FieldMapper[ColumnName].FieldSupprotedValues[value];
                            else
                                WIType = value;
                            continue;
                        }

                        if (ColumnName.StartsWith("Title"))
                        {
                            CurrentWorkItem = value;
                            fields.Add("Title", value);
                            continue;
                        }
                        if (FieldMapper.ContainsKey(ColumnName))
                        {
                            if (FieldMapper[ColumnName].FieldSupprotedValues != null)
                            {
                                if (FieldMapper[ColumnName].FieldSupprotedValues.ContainsKey(value))
                                    fields.Add(FieldMapper[ColumnName].TargetFieldName, FieldMapper[ColumnName].FieldSupprotedValues[value]);
                            }
                            else
                                fields.Add(FieldMapper[ColumnName].TargetFieldName, value);
                        }
                        else
                            fields.Add(column.ColumnName, value);
                    }

                }
                var newWi = WIOps.CreateWorkItem(ProjectName, WIType, fields);

                return newWi.Id.Value;
            }
            catch (Exception E)
            {
                Console.WriteLine("Error  Occured While Creatig WorkItem" + CurrentWorkItem);
                throw E;
            }
        }
        public static void UpdateWIFields()
        {
            try
            {
                foreach (DataRow row in DT.Rows)
                {
                    if (row["Updated"].ToString().ToLower() != "true")
                    {
                        CurrentWorkItem = row["ID"].ToString();
                        Console.WriteLine("Updating Fields of" + CurrentWorkItem);
                        Dictionary<string, object> Updatefields = new Dictionary<string, object>();
                        foreach (DataColumn col in DT.Columns)
                        {
                            if (!string.IsNullOrEmpty(row[col].ToString()))
                            {
                                if (col.ToString() != "ID" && col.ToString() != "Reason"&& UpdateAfterCreating.Contains(col.ColumnName))
                                {
                                    string val = row[col.ToString()].ToString();
                                    if (!string.IsNullOrEmpty(val))
                                    {
                                        if (FieldMapper.ContainsKey(col.ColumnName))
                                        {
                                            if (FieldMapper[col.ColumnName].FieldSupprotedValues.ContainsKey(val))
                                                Updatefields.Add(FieldMapper[col.ColumnName].TargetFieldName, FieldMapper[col.ColumnName].FieldSupprotedValues[val]);
                                            else
                                                Updatefields.Add(FieldMapper[col.ColumnName].TargetFieldName, val);
                                        }
                                        else
                                            Updatefields.Add(col.ColumnName, val);
                                    }

                                }
                            }
                        }
                        WorkItem UpdatedWI = WIOps.UpdateWorkItemFields(int.Parse(row["ID"].ToString()), Updatefields);
                        if (UpdatedWI != null)
                            row["Updated"] = "true";
                    }
                }
            }
            catch (Exception E)
            {
                Console.WriteLine("Error Occured While Updating WorkItem With ID=" + CurrentWorkItem);
                throw E;
            }

        }

        public static DataTable ReadExcel()
        {
            DataTable Dt = new DataTable();

            try
            {
                Excel.Application xlApp = new Excel.Application();

                while (!File.Exists(ExcelPath) || (ExcelPath.EndsWith("xls") && ExcelPath.EndsWith("xlsx")))
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
                    if (!Dt.Columns.Contains("ID"))
                    {
                        DataColumn column = new DataColumn("ID");
                        Dt.Columns.Add(column);
                    }
                    if (!Dt.Columns.Contains("ParentID"))
                    {
                        DataColumn column = new DataColumn("ParentID");
                        Dt.Columns.Add(column);
                    }
                    if (!Dt.Columns.Contains("Updated"))
                    {
                        DataColumn column = new DataColumn("Updated");
                        Dt.Columns.Add(column);
                    }

                    if (i != 1)
                    {
                        Console.WriteLine("Reading excel data row " + i + "/" + rowCount);
                        Dt.Rows.Add(row);
                    }

                }
                xlWorkbook.Close(true);
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error Occured While Reading Excel File");
                throw ex;
            }
            return Dt;
        }
        static DataSet DS = new DataSet();
        private static void ExportDataSetToExcel()
        {
            DS.Tables.Add(DT);
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();
            object file = System.Reflection.Missing.Value;
            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(ExcelPath, file, false, file, file, file, true, file, file, true, file, file, file, file, file);


            foreach (DataTable table in DS.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name

                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets[1];
                if (excelWorkSheet.Name != table.TableName)
                {
                    excelWorkSheet = excelWorkBook.Sheets.Add();
                    excelWorkSheet.Name = table.TableName;
                }
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            excelWorkBook.Save();
            excelWorkBook.Close(true);
            excelApp.Quit();
        }
    }

}

