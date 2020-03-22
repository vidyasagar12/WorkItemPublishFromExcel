using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkItemPublish.Model
{
    /*public class WorkItems
    {
        public List<Epic> Epics { get; set; }
    }
    public class Epic
    {
        public List<Feature> Features { get; set; }
    }
    public class Features*/
    public class WIs
    {
        public List<WIFields> Workitems { get; set; }
    }
   public class WIFields
    {
        public object ID { get; set; }
        public string Title1 { get; set; }
        public string Title2 { get; set; }
        public string Title3 { get; set; }
        public string Title4 { get; set; }

        [JsonProperty(PropertyName = "Area Path")]
        public string Area_Path { get; set; }

        public string Iteration { get; set; }

        [JsonProperty(PropertyName = "Epic Category")]
        public string Epic_Category { get; set; }

        public string Themes { get; set; }
        public string Frameworks { get; set; }


        [JsonProperty(PropertyName = "Epic Type")]
        public string Epic_Type { get; set; }


        [JsonProperty(PropertyName = "Requirement Souce")]
        public string Requirement_Souce { get; set; }


        [JsonProperty(PropertyName = "CLM Created By")]
        public string CLM_Created_By { get; set; }


        [JsonProperty(PropertyName = "CLM Planned For")]
        public string CLM_Planned_For { get; set; }

        public string Bot_or_framework { get; set; }


        [JsonProperty(PropertyName = "CLM_Owned by")]
        public string CLM_Owned_by { get; set; }

        public string CLM_ID { get; set; }
        public string CLM_Priority { get; set; }


        [JsonProperty(PropertyName = "CLM_Filed Against")]
        public string CLM_Filed_Against { get; set; }


        [JsonProperty(PropertyName = "Story_Origin")]
        public string Story_Origin { get; set; }


        [JsonProperty(PropertyName = "Task Category")]
        public string Task_Category { get; set; }


        [JsonProperty(PropertyName = "Task Origin")]
        public string Task_Origin { get; set; }


        [JsonProperty(PropertyName = "Original Estimate")]
        public string Original_Estimate { get; set; }


        [JsonProperty(PropertyName = "Completed Work")]
        public string Completed_Work { get; set; }



        [JsonProperty(PropertyName = "Work Item Type")]
        public string Work_Item_Type { get; set; }

        public string State { get; set; }
        public string AssignedTo { get; set; }

        [JsonProperty(PropertyName = "Story point")]
        public string Story_point { get; set; }

        public string Description { get; set; }
    }
}
