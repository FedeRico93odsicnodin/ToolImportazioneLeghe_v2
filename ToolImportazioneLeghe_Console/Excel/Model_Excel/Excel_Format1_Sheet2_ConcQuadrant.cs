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


        #region COORDINATE EXCEL

        /// <summary>
        /// Indica la posizione di riga iniziale dalla quale si legge per il quadrante corrente 
        /// </summary>
        public int StartingRow_Title { get; set; }


        /// <summary>
        /// Indica la posizione di colonna iniziale dalla quale si legge per il quadrante corrente 
        /// </summary>
        public int StartigCol { get; set; }


        /// <summary>
        /// Indica la posizione di riga per la lettura degli headers
        /// </summary>
        public int StartingRow_Headers { get; set; }


        /// <summary>
        /// Inidica la posizione di riga per la lettura delle concentrazioni
        /// </summary>
        public int StartingRow_Concentrations { get; set; }


        /// <summary>
        /// Indica la posizione di riga per riga di fine lettura delle concentrazioni per il quadrante corrente
        /// </summary>
        public int EndingRow_Concentrations { get; set; }


        /// <summary>
        /// Indica la posizione di riga finale per il quadrante excel corrente
        /// </summary>
        public int EndingCol { get; set; }
                
        #endregion


        /// <summary>
        /// Set delle concentrazioni lette per il quadrante corrente 
        /// </summary>
        public List<Excel_PropertyWrapper> Concentrations { get; set; }
    }
}
