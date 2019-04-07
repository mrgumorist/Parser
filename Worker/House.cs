using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Worker
{
    public class House
    {
        public string Adress { get; set; }
        public string Price { get; set; }
        public string CountOfRooms { get; set; }
        public string Metrazh { get; set; }
        public string Link { get; set; }
        public string Created { get; set; }
        public string Updated { get; set; }
        public House()
        {
            Adress = "Pusto";
            Price = "Pusto";
            CountOfRooms = "Pusto";
            Metrazh = "Pusto";
            Link = "Pusto";
            Created = "Pusto";
            Updated = "Pusto";
        }
    }
}
