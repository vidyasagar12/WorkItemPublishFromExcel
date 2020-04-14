
using System.Collections.Generic;

namespace WorkItemPublish.Model.Mapper
{
    public class Link
    {
        public string source { get; set; }
        public string target { get; set; }
    }

    public class LinkMap
    {
        public List<Link> link { get; set; }
    }

    public class Type
    {
        public string source { get; set; }
        public string target { get; set; }
    }

    public class TypeMap
    {
        public List<Type> type { get; set; }
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
        public string type { get; set; }
        public string required { get; set; }
        public string mapper { get; set; }
        public Mapping mapping { get; set; }
        public string @for { get; set; }
        public string sourceType { get; set; }
        public string not_for { get; set; }
    }

    public class FieldMap
    {
        public List<Field> field { get; set; }
    }

    public class Mapper
    {
        public string source_project { get; set; }
        public string target_project { get; set; }
        public string workspace { get; set; }
        public string sprint_field { get; set; }
        public string base_area_path { get; set; }
        public string base_iteration_path { get; set; }
        public string process_template { get; set; }
        public LinkMap link_map { get; set; }
        public TypeMap type_map { get; set; }
        public FieldMap field_map { get; set; }
    }
}