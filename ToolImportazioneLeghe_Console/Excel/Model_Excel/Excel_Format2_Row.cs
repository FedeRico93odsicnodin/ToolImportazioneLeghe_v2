using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    public class Excel_Format2_Row_LegaProperties
    {
        /// <summary>
        /// Lista degli attributi di lega che è possibile leggere dal foglio
        /// </summary>
        public Excel_PropertyWrapper ReadLegheParams { get; set; }

        
        /// <summary>
        /// Indicazione di indice di riga per il quale sto leggendo le proprieta correnti generali di lega e le sue concentrazioni
        /// </summary>
        public int RowIndexLegaProperties { get; set; }



        /// <summary>
        /// Informazione delle colonne nelle quali sono contenute le diverse informazioni
        /// di concentrazione per le leghe in lettura corrente 
        /// </summary>
        public List<Excel_Format2_ConcColumns> ColonneConcentrazioni { get; set; }


        /// <summary>
        /// Mi dice se le proprieta per la lega corrente sono state correttamente lette durante la fase di lettura 
        /// per le informazioni correnti
        /// Il controllo passa se ho letto tutto il set delle proprieta obbligatorie per la lega corrente 
        /// </summary>
        public bool HoLettoProprietaLega { get; set; }


        /// <summary>
        /// Mi dice se ho letto almeno un valore di concentrazione per uno degli elementi (ho letto tutte le proprietà obbligatorie
        /// indispensabili per la concentrazione)
        /// </summary>
        public bool HoLettoConcentrazioni { get; set; }

    }
}
