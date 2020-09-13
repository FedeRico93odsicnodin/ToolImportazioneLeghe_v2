using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel.Model_Excel;

namespace ToolImportazioneLeghe_Console.Utils
{
    /// <summary>
    /// In questa classe sono inserite tutte le lista comuni nelle quali andare a effettuare i diversi recuperi dalle fonti di origine 
    /// e l'eventuale seconda validazione 
    /// </summary>
    public static class CommonMemList
    {
        /// <summary>
        /// Tutte le informazioni relativamente al foglio excel correntemente aperto
        /// accesso comune dalle diverse classi di recupero informazioni validazione e scrittura 
        /// </summary>
        public static List<Excel_AlloyInfo_Sheet> InformazioniFoglioExcelOrigine { get; set; }
    }
}
