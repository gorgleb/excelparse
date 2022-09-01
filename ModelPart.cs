using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParse
{
    public class ModelPart
    {
        public string Catalog = "";

        public string Model = "";

        public string OEM_Number = "";
        
        public string BodyParts_Number = "";
        
        public string Year = "";
        
        public string Description = "";

        public override string ToString()
        {
            return $"{Catalog} {Model} {OEM_Number} {BodyParts_Number} {Year} {Description}";
        }
    }
}
