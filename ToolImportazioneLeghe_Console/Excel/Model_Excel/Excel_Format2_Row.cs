using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    public class Excel_Format2_Row
    {
        /// <summary>
        /// Lista degli attributi di lega che è possibile leggere dal foglio
        /// </summary>
        public List<Excel_PropertyWrapper> ReadLegheParams { get; set; }

        
        /// <summary>
        /// Lista degli elementi relativi alle concentrazioni
        /// </summary>
        public List<Excel_Format2_ConcColumns> ReadConcentrations { get; set; }
    }
}
