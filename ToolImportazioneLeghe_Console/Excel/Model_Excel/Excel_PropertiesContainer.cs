using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    /// <summary>
    /// Riferimento ad una categorizzazione per un certo set di proprieta in lettura, queste possono essere inputabili ad una lettura per una particolare concentrazione
    /// oppure ad una determinata lega, in ogni caso viene inserita per questa proprieta la riga di riferimento per la sua individuazione, cosi come le eventuali colonne di 
    /// inizio fine lettura per le proprieta correnti
    /// </summary>
    public class Excel_PropertiesContainer
    {
        #region SET DI PROPRIETA CORRENTEMENTE IN ANALISI PER ISTANZA CORRENTE

        /// <summary>
        /// Lista di tutte le proprieta lette per l'istanza corrente
        /// </summary>
        public List<Excel_PropertyWrapper> PropertiesDefinition { get; set; }

        #endregion


        #region IDENTIFICAZIONE DELL'ISTANZA CORRENTE 

        /// <summary>
        /// Nome per l'istanza corrente (può essere il nome di una lega oppure il nome per un elemento)
        /// </summary>
        public string NameInstance { get; set; }


        /// <summary>
        /// Indice di riga sul quale si stanno leggendo le proprieta correntemente in analisi
        /// </summary>
        public int StartingRowIndex { get; set; }


        /// <summary>
        /// Indice di riga sul quale si finiscono di leggere le proprieta correntemente in analisi
        /// </summary>
        public int EndingRowIndex { get; set; }


        /// <summary>
        /// Indice di colonna iniziale sul quale si stanno leggendo le proprieta correntemente in analisi
        /// </summary>
        public int StartColIndex { get; set; }


        /// <summary>
        /// Indice di colonna finale sul quale si stanno leggendo le proprieta correntemente in analisi
        /// </summary>
        public int EndingColIndex { get; set; }

        #endregion


        #region VALIDAZIONI SU PROPRIETA CORRENTE - MI SERVONO PER LA STAMPA DEL RELATIVO LOG
        

        public bool ValidatedElem { get; set; }


        public bool ValidatedAssociation { get; set; }

        #endregion
    }
}
