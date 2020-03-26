using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkItemPublish
{
    class WorkitemFromExcel
    {
        public  int id { get; set; }
        public string tittle { get; set; }
        //public string createdID { get; set; }
        public ParentWorkItem parent { get; set; }
        public string  WiState { get; set; }
        public string  AreaPath { get; set; }
        public string  Itertation { get; set; }
        //public string  WiState { get; set; }
    }
    class ParentWorkItem
    {
        public int Id { get; set; }
        public string tittle { get; set; }
        //public string createdID { get; set; }
    }
}
