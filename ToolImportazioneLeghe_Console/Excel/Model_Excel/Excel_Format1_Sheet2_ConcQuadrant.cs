using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    /// <summary>
    /// Oggetto utilizzato per la mattura dei quadranti per le concentrazioni lette dalla seconda tipologia di foglio per il primo formato excel
    /// relativo al database di leghe
    /// </summary>
    public class Excel_Format1_Sheet2_ConcQuadrant
    {
        /// <summary>
        /// Nome relativo al materiale, questa stringa è quella recuperata direttamente dal foglio
        /// </summary>
        public string NomeMateriale { get; set; }


        /// <summary>
        /// Nome relativo alla lega, questa stringa è inferita rispetto al valore precedente 
        /// </summary>
        public string NomeLega { get; set; }

        
        /// <summary>
        /// Set delle concentrazioni lette per il quadrante corrente 
        /// </summary>
        public List<Excel_PropertyWrapper> Concentrations { get; set; }
    }
}
