using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Excel.Model_Excel
{
    /// <summary>
    /// Wrapper per proprieta relative alla lettura che è possibile eseguire sui 2 fogli 
    /// </summary>
    public class Excel_PropertyWrapper
    {

        #region COSTRUTTORE

        /// <summary>
        /// Valorizzazione della proprieta correntemente in analisi per il caso excel, tutti i parametri sono per l'identificazione 
        /// della proprieta in maniera primaria sono inseriti nel momento in cui la proprieta viene effettivamente creata per il caso corrente 
        /// </summary>
        /// <param name="rowPosition"></param>
        /// <param name="colPosition"></param>
        /// <param name="propertyName"></param>
        /// <param name="isOptional"></param>
        public Excel_PropertyWrapper(int rowPosition, int colPosition, string propertyName, bool isOptional)
        {
            Row_Position = rowPosition;
            Col_Position = colPosition;
            PropertyName = propertyName;
            IsOptional = isOptional;
        }

        #endregion


        #region INDICI PER LA LETTURA DELLA PROPRIETA CORRENTE 

        /// <summary>
        /// Riga posizione per la proprieta corrente nel file excel in lettura 
        /// </summary>
        public int Row_Position { get; set; }


        /// <summary>
        /// Colonna posizione per la proprieta corrente nel file excel in lettura  
        /// </summary>
        public int Col_Position { get; set; }


        /// <summary>
        /// Valore di stringa per il nome della proprieta letta, questa proprieta viene presa dalle definizioni di proprieta 
        /// obbligatoria o opzionale per i 2 formati disponibili di foglio 
        /// </summary>
        public string PropertyName { get; set; }


        /// <summary>
        /// Indicazione di proprieta opzionale per istanza corrente di proprieta
        /// </summary>
        public bool IsOptional { get; set; }


        /// <summary>
        /// Valore di stringa per la proprieta corrente, questo valore puo corrispondere alla definizione di stringa 
        /// che ne verrà già data oppure alla definizione di un qualche valore numerico di cui dovrà essere corretta 
        /// la trasformazione
        /// </summary>
        public string StringValue { get; set; }

        #endregion


        #region VALIDAZIONI SU PROPRIETA CORRENTE - MI SERVONO PER LA STAMPA DEL RELATIVO LOG

        /// <summary>
        /// Indica se l'instanza corrente passa la validazione 1: questa validazione corrisponde 
        /// alla lettura corretta di un qualche valore presente nella cella
        /// </summary>
        public bool Validation1_OK { get; set; }


        /// <summary>
        /// Indica se l'instanza corrente passa la validazione 2: questa validazione corrisponde 
        /// all'associazione con un valore netto per l'instanza corrente (per esempio le stringhe 
        /// convertite correttamente in valori numerici)
        /// </summary>
        public bool Validation2_OK { get; set; }


        /// <summary>
        /// Indica se l'istanza corrente passa la validazione 3: questa validazione corrisponde 
        /// all'associazione fatta SULLO STESSO EXCEL rispetto alle informazioni recuperate 
        /// (per esempio trovo che tutte le definizioni per una particolare concentrazione sono associate correttamente 
        /// alle informazioni di lega e viceversa)
        /// </summary>
        public bool Validation3_OK { get; set; }


        /// <summary>
        /// Passaggio per la validazione 4: questa validazione è fatta rispetto alle definizioni della destinazione 
        /// rispetto alla quale si va a persistere l'informazione (per esempio nella destinazione esiste già la definizione 
        /// per l'istanza corrente, quindi non ha senso persisterla)
        /// </summary>
        public bool Validation4_OK { get; set; }

        #endregion

    }
}
