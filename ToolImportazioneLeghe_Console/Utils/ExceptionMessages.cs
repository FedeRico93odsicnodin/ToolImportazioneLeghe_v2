using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Utils
{
    /// <summary>
    /// Messaggi di errore che è possibile incontrare utilizzando il programma nelle 2 versioni
    /// </summary>
    public static class ExceptionMessages
    {
        #region LOGGING

        /// <summary>
        /// Segnalazione di non poter continuare per via di una mancata ricostruzione di parte del path per 
        /// la ricreazione del file di log 
        /// </summary>
        public static string LOGGING_NONHOTROVATOPARTEDELPATH = "impossibile ricreare il file di log, non è stata trovata parte del path '{0}'";


        /// <summary>
        /// Segnalazione di non aver trovato il file di log sul quale andare a scrivere una certa procedura 
        /// </summary>
        public static string LOGGING_INVALIDLOGPATH = "non ho trovato il log '{0}' file sul quale si è chiesto di scrivere un comando";


        /// <summary>
        /// Segnalazione generica di errore nel tentativo di configurazione della procedura di logging
        /// </summary>
        public static string LOGGING_VERIFICADIUNCERTOERRORE = "ho riscontrato il seguente errore durante la configurazione del LOGGING\n\n{0}";

        #endregion


        #region UTIL FUNCTIONS

        /// <summary>
        /// Segnalazione di un path nullo per l'ottenimento del nome proprio del file 
        /// </summary>
        public static string UTILFUNCTIONS_PATHNULL = "il path per un certo file è nullo o vuoto";


        /// <summary>
        /// Segnalazione di un nome nullo attribuito al file a un certo path che comunque è stato analizzato
        /// </summary>
        public static string UTILFUNCTIONS_NOMEVUOTO = "il nome attribuito al file è nullo";


        /// <summary>
        /// Segnalazione di errore nel tentativo di ricavare il nome per un determinato file 
        /// </summary>
        public static string UTILFUNCTIONS_ERRORENELLALETTURANOMEFILE = "ho riscontrato un errore nel tentativo di ricavare un nome per un file\n\n{0}";

        #endregion


        #region EXCEL 

        /// <summary>
        /// Eccezione relativa all'assenza di path per il file excel corrente 
        /// </summary>
        public static string EXCEL_EMPTYPATH = "il path passato per la lettura del file excel non corrisponde a nessuna informazione";


        /// <summary>
        /// Eccezione di inesistenza del file excel che viene passato in apertura per leggere le informazioni
        /// </summary>
        public static string EXCEL_SOURCENOTEXISTING = "il file excel passato come sorgente non esiste";


        /// <summary>
        /// Eccezione relativa alla non definizione per il formato relativo al file excel in apertura corrente 
        /// </summary>
        public static string EXCEL_FORMATNOTDEFINED = "il formato per il file excel in apertura corrente non è stato definito";


        /// <summary>
        /// Eccezione mancata apertura corretta per il foglio excel passato in input
        /// </summary>
        public static string EXCEL_PROBLEMAAPERTURAFILE = "si è verificato il seguente problema durante l'apertura del file excel '{0}':\n\n{1}";


        /// <summary>
        /// Eccezione relativa al fatto che il file excel corrente non è stato correttamente predisposto in memoria per l'analisi
        /// </summary>
        public static string EXCEL_FILENOTINMEMORY = "il file excel per l'analisi non è stato correttamente caricato in memoria";


        /// <summary>
        /// Eccezione utilizzata a livello di creazione e inserimento proprietà excel in lettura all'interno del wrapper
        /// </summary>
        public static string EXCEL_READINGPROPERTIES = "si è verificato un errore di implementazione durante la creazione istanza di proprieta excel";


        /// <summary>
        /// Ricerca per proprieta contenuta nel wrapper non andata a buon fine 
        /// </summary>
        public static string EXCEL_PROPERTYNOTDEFINED = "la proprieta ricercata durante la fase di analisi non è stata definita all'interno del wrapper per '{0}'";


        /// <summary>
        /// Eccezione relativa al fatto che manca l'header di colonna nelle definizioni obbligatorie e sulla quale andare a riconoscere la presenza di eventuali 
        /// elementi sottostanti
        /// </summary>
        public static string EXCEL_COLCRITERINONPRESENTE = "non posso proseguire l'analisi per individuare se il foglio corrente è di concentrazioni: manca la definizione della colonna su cui distinguere gli elementi (CRITERI)";


        /// <summary>
        /// Eccezione lanciata nel caso in cui, per un motivo sconosciuto, la lista finale dei fogli dai quali lo step 3 andrebbe a leggere le informazioni, è stata trovata null o empty
        /// </summary>
        public static string EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT1 = "non posso continuare con il recupero di tutte le informazioni dal file EXCEL sorgente, riconosciuto come essere di primo formato, in quanto la lista dei FOGLI è NULL o EMPTY";


        /// <summary>
        /// Eccezione che si verifica nel caso in cui, durante la decisione del tipo di analisi per il recupero delle informazioni dal foglio corrente, la tipologia di foglio è trovata come NOT DEFINED
        /// </summary>
        public static string EXCEL_READERINFO_TIPOLOGIANONDEFINITAFOGLIOCORRENTE = "non posso continuare con l'analisi e il recupero delle informazioni perché uno dei fogli nel set è di TIPOLOGIA NON DEFINITA";


        /// <summary>
        /// Eccezione che si verifica nel caso in cui la posizione nel file excel di origine per il foglio in analisi corrente è stata trovata come = 0 cioe non valida per l'analisi e il recupero delle informazioni successivi
        /// </summary>
        public static string EXCEL_READERINFO_NESSUNAPOSIZIONETROVATAPERFOGLIOCORRENTE = "non posso continuare con l'analisi e il recupero delle informazioni perché uno dei fogli ha una posizione = 0 per il file EXCEL e quindi non valida";


        /// <summary>
        /// Eccezione lanciata nel caso in cui, per un motivo sconosciuto, la lista finale dei fogli dai quali lo step 3 andrebbe a leggere le informazioni, è stata trovata null o empty
        /// </summary>
        public static string EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT2 = "non posso continuare con il recupero di tutte le informazioni dal file EXCEL sorgente, riconosciuto come essere di secondo formato, n quanto la lista dei FOGLI è NULL o EMPTY";


        ////////////////////////////////////////////////////////////////////////////// ECCEZIONI LETTURA INFORMAZIONI DA FOGLI VALIDATI ////////////////////////////////////////////////////////////////////////////// 


        /// <summary>
        /// Eccezione lanciata nel caso in cui il foglio in input dal quale andare a leggere le informazioni è NULL per il caso corrente 
        /// </summary>
        public static string EXCEL_READERINFO_FOGLIONULLPERLETTURA = "non posso continuare con la lettura corrente in quanto il foglio in input è NULL";


        /// <summary>
        /// Eccezione lanciata prima della lettura effettiva per le informazioni del foglio corrente, gli indici mi definiscono il perimetro di azione rispetto alle celle di lettura se uno di questi è ZERO non posso continuare con l'analisi
        /// </summary>
        public static string EXCEL_READERINFO_INDICIDILETTURAZERO = "non posso continuare con la lettura delle informazioni per il foglio corrente in quanto uno degli indici di lettura è stato trovato = ZERO";


        /// <summary>
        /// Eccezione lanciata prima della lettura effettiva per le informazioni dal foglio corrente, gli indici non rispettano i vincoli di maggiore minore e quindi la lettura a priori sarebbe non valida
        /// </summary>
        public static string EXCEL_READERINFO_INDICINONVALIDI = "non posso continuare con la lettura delle informazioni per il foglio corrente in quanto il valore attribuito agli INDICI non è valido";


        /// <summary>
        /// Eccezione lanciata prima della lettura effettiva delle informazioni dal foglio corrente, una proprieta interna di lettura è stata trovata nulla per il caso corrente 
        /// </summary>
        public static string EXCEL_READERINFO_PROPRIETAINTERNANULLA = "non posso continuare con la lettura delle informazioni per il foglio interno in quanto una proprieta interna è NULLA";


        /// <summary>
        /// Eccezione lanciata prima della lettura delle informazioni effettiva per il foglio, non sono lette correttamente tutte le proprieta di headers sulle quali eseguire la successiva lettura valori
        /// </summary>
        public static string EXCEL_READERINFO_VINCOLILETTURAPROPRIETANONRISPETTATI = "non posso continuare con la lettura delle informazioni per il foglio, in quanto non sono rispettati tutti i vincoli di validazione sulle proprieta del foglio corrente";
        
        #endregion


        #region STEPS FROM EXCEL TO DATABASE

        /// <summary>
        /// Eccezione su STEP 1 nel caso in cui il file excel non sia stato aperto correttamente per la lettura successiva delle informazioni
        /// </summary>
        public static string ERRORESTEP1_APERTURAFILEEXCEL = "STEPS: non posso continuare con lo STEP 2 di validazione fogli da EXCEL perché questo non è stato aperto correttamente";


        /// <summary>
        /// Eccezione su STEP 2 nel caso in cui i fogli per il file excel corrente non siano stati validati correttamente per la lettura di certe informazioni da questi
        /// </summary>
        public static string ERRORESTEP2_VALIDAZIONEFOGLIEXCEL = "STEPS: non posso continuare con lo STEP 3 di lettura informazioni da EXCEL perché i fogli non sono stati validati correttamente";

        #endregion
    }
}
