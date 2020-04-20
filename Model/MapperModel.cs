
using System.Collections.Generic;

namespace WorkItemPublish.Model.Mapper
{
    public class WIType
    {
        public string source { get; set; }
        public string target { get; set; }
    }

    public class TypeMap
    {
        public List<WIType> type { get; set; }
    }

    public class Value
    {
        public string source { get; set; }
        public string target { get; set; }
    }

    public class Mapping
    {
        public List<Value> values { get; set; }
    }

    public class Field
    {
        public string source { get; set; }
        public string target { get; set; }
        public string required { get; set; }
        public Mapping mapping { get; set; }
        public string ShouldUpdateAfterCreating { get; set; }

    }

    public class FieldMap
    {
        public List<Field> field { get; set; }
    }

    public class Mapper
    {
        public string source_project { get; set; }
        public string target_project { get; set; }    
        public string Unique_ID_Field { get; set; }   
        public string WorkItem_Type_Field { get; set; }

        public TypeMap type_map { get; set; }
        public FieldMap field_map { get; set; }
    }
}