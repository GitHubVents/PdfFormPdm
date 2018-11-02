using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormPdf
{
   public class BomShell
    {
        public string Configuration { get; set; }
        public int Version { get; set; }
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public int IdPdm { get; set; }
        public string FileName { get; set; }
        public string FolderPath { get; set; }
        public string ObjectType { get; set; }
        public string Partition { get; set; }
    }
}
