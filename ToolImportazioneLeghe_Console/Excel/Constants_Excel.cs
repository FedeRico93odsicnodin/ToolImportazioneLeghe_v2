using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Excel
{
    /// <summary>
    /// Inserimento di tutte le costati di lettura del foglio excel, che sia per il primo o il secondo formato e per i diversi 
    /// fogli disponibili all'interno del file 
    /// </summary>
    public static class Constants_Excel
    {
        #region PARAMETRI COMUNI AI 2 FOGLI 

        /// <summary>
        /// Mappatura dei possibili casi per quanto riguarda la creazione di un nuovo wrapper di proprieta
        /// per la tipologia di foglio corrente
        /// </summary>
        public enum TipologiaPropertiesFoglio
        {
            Format1_Foglio1_Leghe = 1,
            Format1_Foglio2_Concentrazioni = 2,
            Format2_Leghe = 3,
            Format2_ConcentrazioniElemento = 4
        }

        #endregion


        #region PARAMETRI PER LA LETTURA DEL FORMATO 1 PRESO DAL DATABASE DI LEGHE

        /// <summary>
        /// Indica se il foglio correntemente in analisi è relativo alla lettura delle informazioni di lega 
        /// piuttosto che dei quadranti per le diverse concentrazioni
        /// </summary>
        public enum TipologiaFoglio_Format1
        {
            NotDefined = 0,
            FoglioLeghe = 1,
            FoglioConcentrazioni = 2
        }


        /// <summary>
        /// Esito della validazione per il formato excel 1 - le informazioni lette e validate possono rientrare in una della seguente casistica 
        /// - per la prima tipologia di lettura si prosegue senza chiedere nessuna conferma a user
        /// - per la seconda tipologia si chiede conferma allo user anche in base alla modalita impostata di sovrascrittura o accodamento informazioni
        /// - per la terza tipologia si ferma il tool
        /// </summary>
        public enum ValidazioneFoglio
        {
            LetturaCompleta = 1,
            LetturaParziale = 2,
            NessunaCorrispondenza = 3
        }


        /// <summary>
        /// Segnalazioni rispettivamente al recupero delle informazioni dal foglio excel corrente
        /// questo puo produrre Warnings oppure Eccezioni vere e proprie per quanto riguarda la lettura delle informazioni di partenza
        /// </summary>
        public enum EsitoRecuperoInformazioniFoglio
        {
            RecuperoCorretto = 1,
            RecuperoConWarnings = 2,
            RecuperoConErrori = 3
        }


        /// <summary>
        /// Indicazione di quale sia lo step di lettura per l'analisi del foglio corrente in base ai parametri passati nelle configurazioni
        /// serve per andare a inserire adeguata messaggistica circa l'analisi del foglio corrente 
        /// </summary>
        public enum StepLetturaFoglio
        {
            Riconoscimento = 1,
            RecuperoInformazioni_Validazione1 = 2,
            Validazione2_StessoFile = 3
        }


        /// <summary>
        /// Proprietà obbligatorie per la lettura delle informazioni dal primo foglio (FORMAT1)
        /// </summary>
        public static string[] PROPRIETAOBBLIGATORIE_FORMAT1_SHEET1 = { "MATERIALE", "NORMATIVA", "TIPO" };


        /// <summary>
        /// Proprietà opzionali per la lettura delle informazioni dal primo foglio (FORMAT1)
        /// </summary>
        public static string[] PROPRIETAOPZIONALI_FORMAT1_SHEET1 = { "#", "PAESE / PRODUTTORE" };


        /// <summary>
        /// Proprietà obbligatorie per la lettura delle informazioni dal secondo foglio (FORMAT1)
        /// </summary>
        public static string[] PROPRIETAOBBLIGATORIE_FORMAT1_SHEET2 = { "CRITERI", "MIN", "MAX" };


        /// <summary>
        /// Proprietà opzionali per la lettura delle informazioni per il seconto foglio (FORMAT1)
        /// </summary>
        public static string[] PROPRIETAOPZIONALI_FORMAT1_SHEET2 = { "APPROSS", "COMMENTO" };

        #endregion


        #region PARAMETRI PER LA LETTURA DEL FORMATO 2 CLIENTE

        /// <summary>
        /// Proprieta obbligatorie per la lettura delle informazioni di base e per il formato 2
        /// </summary>
        public static string[] PROPRIETAOBBLIGATORIE_FORMAT2_LEGHE = { "NORMATIVA", "BASE", "MAT. NO.", "ALLOY NAME" };


        /// <summary>
        /// Proprieta opzionali per la lettura delle inforamzioni di base e per il formato 2
        /// </summary>
        public static string[] PROPRIETAOPZIONALI_FORMAT2_LEGHE = { "CATEGORIA", "DESCRIZIONE", "TRATTAMENTO" };


        /// <summary>
        /// Proprieta obbligatorie per la lettura delle informazioni di elemento e per il formato 2
        /// </summary>
        public static string[] PROPRIETAOBBLIGATORIE_ELEM_FORMAT2 = { "MIN.", "MAX." };


        /// <summary>
        /// Proprieta opzionali per la lettura delle informazioni di elemento e per il formato 2
        /// </summary>
        public static string[] PROPRIETAOPZIONALI_ELEM_FORMAT2 = { "DEROGAMIN.", "DEROGAMAX.", "OBIETTIVO" };

        #endregion
    }
}
