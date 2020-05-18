using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    /// <summary>
    /// Oggetto che identifica le colonne nelle quali andare a leggere le diverse proprieta legate alle concentrazioni
    /// per la seconda tipologia di foglio excel in lettura corrente 
    /// </summary>
    public class Excel_Format2_ConcColumns
    {
        /// <summary>
        /// Nome per l'elemento corrente, letto nell'intestazione di colonna 
        /// </summary>
        public string NomeElemento { get; set; }
        

        /// <summary>
        /// Riga di partenza per la lettura, corrisponde alla riga nella quale 
        /// è presente la definizione per l'elemento corrente 
        /// </summary>
        public int startingRow_Elemento { get; set; }

        
        /// <summary>
        /// Colonna di partenza per la lettura delle proprieta lette 
        /// </summary>
        public int startingCol_Header { get; set; }


        /// <summary>
        /// Colonna di fine per la lettura delle proprieta lette
        /// </summary>
        public int endingCol_Header { get; set; }


        /// <summary>
        /// Indicazione dei valori per tutte le proprieta lette per l'analisi corrente 
        /// </summary>
        public Excel_PropertyWrapper ReadProperties { get; set; }
        
    }
}
