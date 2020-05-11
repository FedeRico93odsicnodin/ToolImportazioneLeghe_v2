using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    /// <summary>
    /// Oggetto utilizzato per la mappatura delle informazioni di riga contenute nella tipologia di foglio 1 per
    /// il formato utilizzato in lettura delle informazioni dall'excel database leghe 
    /// </summary>
    public class Excel_Format1_Sheet1_Rows
    {
        /// <summary>
        /// Lista di tutte le proprieta settate attraverso la definizione del property wrapper
        /// </summary>
        public List<Excel_PropertyWrapper> ExcelSheet1_LegheProperties { get; set; }


        /// <summary>
        /// Istanza di lista per le proprieta obbligatorie e opzionali in lettura corrente 
        /// </summary>
        public Excel_Format1_Sheet1_Rows()
        {
            ExcelSheet1_LegheProperties = new List<Excel_PropertyWrapper>();
        }
    }
}
