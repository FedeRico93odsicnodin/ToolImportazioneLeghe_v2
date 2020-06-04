using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Messaging_Console
{
    public static class ConsoleService
    {
        #region TUTTI I MESSAGGI PER LA CONSOLE 

        /// <summary>
        /// Servizio corrente per i messaggi di console provenienti dal servizio di logging
        /// </summary>
        private static ConsoleService_UtilityFunctionsMessages _consoleService_LoggingMessage; 


        /// <summary>
        /// Servizio corrente per i messaggi di console provenienti dal servizio excel
        /// </summary>
        private static ConsoleService_Excel _consoleService_Excel;


        /// <summary>
        /// Servizio corrente con tutti i messaggi relativi alla segnalazione sui vari steps per le diverse importazione 
        /// in esecuzione
        /// </summary>
        private static STEPS_FromExcelToDatabase_ConsoleServices _steps_ConsoleMessages;

        #endregion


        #region METODI PER L'EFFETTIVO UTILIZZO DELLA CONSOLE 

        /// <summary>
        /// Permette la formattazione per il messaggio visualizzato in console e il suo log nel file 
        /// di log globale (nel caso sia specificato come parametro di input)
        /// </summary>
        /// <param name="currentMessage"></param>
        /// <param name="logMessage"></param>
        public static void FormatMessageConsole(string currentMessage, bool logMessage)
        {
            Console.WriteLine(currentMessage);

            if (logMessage)
                ServiceLocator.GetLoggingService.LogMessage(currentMessage, Constants.GlobalLoggingPath);
        }


        /// <summary>
        /// Servie per ottenere il separatore delle diverse attivita (ad esempio per i diversi steps legati all'importazione 
        /// </summary>
        public static void GetSeparatoreAttivita()
        {
            string currentMessage = "----------------------------------------";

            Console.WriteLine(currentMessage);
            ServiceLocator.GetLoggingService.LogMessage(currentMessage, Constants.GlobalLoggingPath);
        }


        /// <summary>
        /// Permette di fare passare un po di tempo all'interno della console
        /// </summary>
       public static void GetSomeTimeOnConsole()
        {

            for (int i = 0; i < 3; i++)
            {
                Thread.Sleep(1000);
                Console.Write(".");
            }

            Console.WriteLine("\n");
        }


        #endregion


        #region TUTTI I SERVIZI DI CONSOLE

        public static ConsoleService_UtilityFunctionsMessages ConsoleLogging
        {
            get
            {
                if (_consoleService_LoggingMessage == null)
                    _consoleService_LoggingMessage = new ConsoleService_UtilityFunctionsMessages();

                return _consoleService_LoggingMessage;
            }
        }


        /// <summary>
        /// Getter per messaging console excel services
        /// </summary>
        public static ConsoleService_Excel ConsoleExcel
        {
            get
            {
                if (_consoleService_Excel == null)
                    _consoleService_Excel = new ConsoleService_Excel();

                return _consoleService_Excel;
            }
        }


        /// <summary>
        /// Getter per messaging console per i vari steps di importazione
        /// </summary>
        public static STEPS_FromExcelToDatabase_ConsoleServices ConsoleStepsMessages
        {
            get
            {
                if (_steps_ConsoleMessages == null)
                    _steps_ConsoleMessages = new STEPS_FromExcelToDatabase_ConsoleServices();

                return _steps_ConsoleMessages;
            }
        }


        /// <summary>
        /// Getter per messaging console relativa agli STEPS eseguiti per l'importazione da excel a database
        /// </summary>
        public static STEPS_FromExcelToDatabase_ConsoleServices STEPS_FromExcelToDatabase
        {
            get
            {
                if (_steps_ConsoleMessages == null)
                    _steps_ConsoleMessages = new STEPS_FromExcelToDatabase_ConsoleServices();

                return _steps_ConsoleMessages;
            }
                
        }


        #endregion
    }


    /// <summary>
    /// Tutti i serivizi di messaging console per il servizio di logging
    /// </summary>
    public class ConsoleService_UtilityFunctionsMessages
    {
        /// <summary>
        /// Marker per i messaggi provienti dal servizio di log 
        /// </summary>
        private const string utilityFunctions_Marker = "UTILITY: ";


        /// <summary>
        /// Segnalazione di esistenza per il file di log attuale, la procedura verrà accodata 
        /// a questo file di log 
        /// </summary>
        /// <param name="logFilePath"></param>
        public void EsistenzaFileLog_Message(string logFilePath)
        {
            string currentMessage = String.Format(utilityFunctions_Marker + "il file di LOG al path \n'{0}'\nesiste già, la procedura verrà loggata in accodamento.", logFilePath);
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }


        /// <summary>
        /// Segnalazione di ricreazione della folder per inserire successivamente il log nel percorso 
        /// esplicitato
        /// </summary>
        /// <param name="analyzedFolder"></param>
        /// <param name="targetLogFile"></param>
        public void RicreazioneFolder(string analyzedFolder, string targetLogFile)
        {
            string currentMessage = String.Format(utilityFunctions_Marker + "ho appena creato la seguente cartella '{0}' per l'inserimneto del log '{1}'", analyzedFolder, targetLogFile);
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }


        /// <summary>
        /// Segnalazione di una ricreazione del file nel quale verranno loggate parti della procedura 
        /// </summary>
        /// <param name="targetLogFile"></param>
        public void RicreazioneLogFile(string targetLogFile)
        {
            string currentMessage = String.Format(utilityFunctions_Marker + "ho appena ricreato il seguente file di log '{0}', la procedura verrà loggata da 0 in questo file", targetLogFile);
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }
    }


    /// <summary>
    /// Tutti i servizi di messaging console per il servizio excel 
    /// </summary>
    public class ConsoleService_Excel
    {
        #region ATTRIBUTI CREAZIONE INTESTAZIONE MESSAGGIO

        /// <summary>
        /// Excel per il marker principale per i servizi excel
        /// </summary>
        private const string excelService_Marker = "EXCEL - ";


        /// <summary>
        /// Marker per i reader di inserimento excel 
        /// </summary>
        private const string readerExcel_Marker = "READER: ";


        /// <summary>
        /// Marker per identificare il formato 1 proveniente dal database di leghe 
        /// </summary>
        private const string intestazioneFormat1 = "FORMAT 1 (DATABASE LEGHE) - ";


        /// <summary>
        /// Marker per identificare il formato 2 proveninete dal cliente 
        /// </summary>
        private const string intestazioneFormat2 = "FORMAT 2 (CLIENTE) - ";

        #endregion


        #region APERTURA FILE EXCEL CORRENTE 

        /// <summary>
        /// Segnalazione esistenza per il file excel corrente 
        /// </summary>
        /// <param name="logFilePath"></param>
        public void EsistenzaFileExcel_Message(string logFilePath)
        {
            string currentMessage = String.Format(excelService_Marker + "ho appena aperto il file excel '{0}'", logFilePath);
            ConsoleService.FormatMessageConsole(currentMessage, false);
        }

        #endregion


        #region VALIDAZIONE FOGLI PER FILE EXCEL CORRENTE RISPETTO AI FORMATI E ALLE DIVERSE TIPOLOGIE DI FOGLIO DISPONIBILI

        /// <summary>
        /// Segnalazione che il foglio è stato trovato di una certa tipologia per l'istanza di reader corrente e per il formato corrente
        /// questo metodo è riferito al primo formato proveniente dal database di tutte le leghe
        /// </summary>
        /// <param name="currentFoglio"></param>
        /// <param name="currentPosizione"></param>
        /// <param name="currentTipologia"></param>
        public void ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format1(string currentFoglio, int currentPosizione, string currentTipologia)
        {
            string currentMessage = String.Format(excelService_Marker + intestazioneFormat1 + readerExcel_Marker + "il foglio di nome '{0}' in posizione in excel {1} è stato riconosciuto come {2}", currentFoglio, currentPosizione, currentTipologia);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione che il foglio è stato trovato come contenente delle informazioni valide per la lettura di leghe e concentrazioni 
        /// inerentemente la seconda tipologia di foglio
        /// </summary>
        /// <param name="currentFoglio"></param>
        /// <param name="currentPosizione"></param>
        public void ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format2(string currentFoglio, int currentPosizione)
        {
            string currentMessage = String.Format(excelService_Marker + intestazioneFormat2 + readerExcel_Marker + "il foglio di nome '{0}' in posizione excel {1} è stato riconosciuto come contenere delle informazioni valide", currentFoglio, currentPosizione);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione che il foglio non è stato riconosciuto come contenere delle informazioni valide per il file corrente e per le informazioni 
        /// di lega o di concentrazione
        /// </summary>
        /// <param name="currentFoglio"></param>
        /// <param name="currentPosizione"></param>
        public void ExcelReaders_Message_FoglioNonRiconosciuto(string currentFoglio, int currentPosizione)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "il seguente foglio '{0}' in posizione {1} non è stato riconosciuto come foglio di informazioni valide, si prega di controllarne nuovamente il contenuto", currentFoglio, currentPosizione);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }

        #endregion


        #region LETTURA INFORMAZIONI + VALIDAZIONE 1 (EMPTY SPACES)

        /// <summary>
        /// Segnalazione di avvenuta lettura informazioni corretta per la tipologia di formato 1 excel e la natura del foglio passato in input
        /// nel messaggio verrà indicata anche la posizione e il nome per il foglio in analisi
        /// </summary>
        /// <param name="foglioLettura"></param>
        /// <param name="posizioneFoglio"></param>
        /// <param name="tipologiaFoglioFormato1"></param>
        public void ExcelReaders_Message_LetturaFoglioTipoFormato1AvvenutaCorrettamente(string foglioLettura, int posizioneFoglio, Constants_Excel.TipologiaFoglio_Format tipologiaFoglioFormato1)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "PRIMO FORMATO: ho letto correttamente le informazioni per tipologia '{0}' per il foglio '{1}' in posizione {2}", tipologiaFoglioFormato1, foglioLettura, posizioneFoglio);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione sempre per il foglio corrente che la lettura è avvenuta con dei warnings rispetto alla validazione contenimento dei valori per il foglio di cui si sono recuperati i valori
        /// questi warnings non portano al momento al blocco dell'applicazione ma potrebbero non coincidere con la situazione voluta dall'utente
        /// </summary>
        /// <param name="foglioLettura"></param>
        public void ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConWarnings(string foglioLettura)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "PRIMO FORMATO: per il foglio '{0}' la lettura è pero avvenuta con WARNINGS (consultare il relativo log per maggiori informazioni)", foglioLettura);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione che il recupero per il certo foglio excel del primo formato e per le informazioni contenute per lega / concentrazioni non è avvenuto correttamente 
        /// cioe non ha passato la fase 1 di validazione rispetto al contenimento delle informazioni di base per poter continuare con l'analisi successiva
        /// </summary>
        /// <param name="foglioLetturaCorrente"></param>
        /// <param name="posizioneFoglio"></param>
        /// <param name="tipologiaFoglioFOrmato1"></param>
        public void ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConErrori(string foglioLetturaCorrente, int posizioneFoglio, Constants_Excel.TipologiaFoglio_Format tipologiaFoglioFormato1)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "PRIMO FORMATO: ERRORE nella lettura delle informazioni del foglio '{0}' in posizione {1} e per il formato '{2}', non potrò proseguire con l'analisi di questo foglio", foglioLetturaCorrente, posizioneFoglio, tipologiaFoglioFormato1);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnaalzione di prossima interruzione per l'analisi corrente, non sono stati riconosciuti correttamente tutte le tipologie minime per poter andare successivamente 
        /// a leggere per concentrazioni e materiali per il file excel corrente
        /// </summary>
        /// <param name="fileExcelName"></param>
        public void ExcelReaders_Message_InterruzioneAnalisi_SetFogliRiconosciutiInsufficiente(string fileExcelName)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "PRIMO FORMATO: non posso proseguire con l'ANALISI e l'IMPORT per il file '{0}' in quanto non sono state riconosciute le tipologie di FOGLIO per LEGHE E CONCENTRAZIONI minimo e indispensabile per il proseguimento", fileExcelName);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione di recupero corretto per le informazioni contenute nel secondo formato disponibile per la lettura di leghe e concentrazioni
        /// </summary>
        /// <param name="foglioLettura"></param>
        /// <param name="posizioneFoglio"></param>
        public void ExcelReaders_Message_LetturaFoglioTipoFormato2AvvenutaCorrettamente(string foglioLettura, int posizioneFoglio)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "SECONDO FORMATO: ho letto correttamente le informazioni per tipologia per il foglio '{0}' in posizione {1}", foglioLettura, posizioneFoglio);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione che per il secondo formato il recupero delle informazioni al momento non è invalidante ma è avvenuto con dei warnings 
        /// che devono essere presi in considerazione
        /// </summary>
        /// <param name="foglioLettura"></param>
        public void ExcelReaders_Message_LetturaFoglioFormato2AvvenutaConWarnings(string foglioLettura)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "SECONDO FORMATO: per il foglio '{0}' la lettura è pero avvenuta con WARNINGS (consultare il relativo log per maggiori informazioni)", foglioLettura);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Segnalazione di errori durante la lettura per il foglio del secondo formato disponibile, l'analisi e successivi import 
        /// non saranno più disponibili per il foglio in analisi 
        /// </summary>
        /// <param name="foglioLetturaCorrente"></param>
        /// <param name="posizioneFoglio"></param>
        /// <param name="tipologiaFoglioFOrmato1"></param>
        public void ExcelReaders_Message_LetturaFoglioFormato2AvvenutaConErrori(string foglioLetturaCorrente, int posizioneFoglio)
        {
            string currentMessage = String.Format(excelService_Marker + readerExcel_Marker + "SECONDO FORMATO: ERRORE nella lettura delle informazioni del foglio '{0}' in posizione {1} e per il formato 2, non potrò proseguire con l'analisi di questo foglio", foglioLetturaCorrente, posizioneFoglio);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }

        #endregion

    }


    /// <summary>
    /// Tutti i messaggi per l'esecuzione di un import dal file excel sorgente aperto
    /// al database (inserito come destinazione)
    /// </summary>
    public class STEPS_FromExcelToDatabase_ConsoleServices
    {
        #region MARKERS

        /// <summary>
        /// Marker per step 1 messages
        /// </summary>
        private const string STEP1 = "STEP1 : ";


        /// <summary>
        /// Marker per step 2 messages
        /// </summary>
        private const string STEP2 = "STEP2 : ";


        /// <summary>
        /// Marker per step 3 messages
        /// </summary>
        private const string STEP3 = "STEP3 : ";

        #endregion

        #region FROM EXCEL TO DATABASE STEPS

        #region MESSAGGI STEP 1 - APERTURA FILE EXCEL

        /// <summary>
        /// Inizio apertura foglio excel sorgente di cui viene passato il nome in input
        /// </summary>
        /// <param name="foglioExcelSorgente"></param>
        public void STEP1_InizioAperturaFoglioExcelSorgente(string foglioExcelSorgente)
        {
            string currentMessage = String.Format(STEP1 + "sto aprendo il file excel '{0}' in modalita LETTURA", foglioExcelSorgente);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Indicazione che il file excel sorgente è stato aperto correttamente 
        /// </summary>
        public void STEP1_FileExcelSorgenteApertoCorrettamente()
        {
            string currentMessage = String.Format(STEP1 + "l'apertura è avvenuta correttamente");
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }

        #endregion


        #region MESSAGGI STEP 2 - VALIDAZIONE FOGLI PER IL FILE EXCEL CORRENTE IN BASE AL FORMATO IMPOSTATO IN FASE DI CONFIGURAZIONE

        /// <summary>
        /// Inizio di validazione dei diversi fogli contenuti all'interno del file excel corrente 
        /// </summary>
        /// <param name="foglioExcelSorgente"></param>
        public void STEP2_InizioValidazioneFogliContenutiInExcel(string foglioExcelSorgente)
        {
            string currentMessage = String.Format(STEP2 + "sto VALIDANDO i fogli contenuti nel file excel '{0}' in base al formato impostato", foglioExcelSorgente);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// Indicazione di fine validazione in maniera corretta per i fogli excel contenuti all'interno del file
        /// </summary>
        public void STEP2_FineSorgenteValidatoCorrettamente()
        {
            string currentMessage = String.Format(STEP2 + "la validazione dei diversi fogli excel è avvenuta correttamente");
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }

        #endregion


        #region MESSAGGI STEP 3 - RECUPERO DI TUTTE LE INFORMAZIONI PER IL FILE EXCEL CORRENTEMENTE APERTO E VALIDATO

        /// <summary>
        /// Indicazione di inizio recupero di tutte le informazioni contenute all'interno dei fogli excel validati correttamente 
        /// per la tipologia passata in input
        /// </summary>
        /// <param name="foglioExcelSorgente"></param>
        public void STEP3_InizioRecuperoDiTutteLeInformazioniPerExcelCorrente(string foglioExcelSorgente)
        {
            string currentMessage = String.Format(STEP3 + "inizio della lettura di tutte le informazioni contenute nel file excel '{0}'", foglioExcelSorgente);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }


        /// <summary>
        /// indicazione di ultimato recupero delle informazioni per il file excel corrente, si fa riferimento alla consultazione dei log 
        /// prodotti se eventualmente sono stati prodotti dei warnings durante questa operazione
        /// </summary>
        public void STEP3_RecuperoDelleInformazioniUltimato(string foglioExcelSorgente)
        {
            string currentMessage = String.Format(STEP3 + "lettura delle informazioni per il file '{0}' avvenuta con successo, consultare il log prodotto per maggiori informazioni", foglioExcelSorgente);
            ConsoleService.FormatMessageConsole(currentMessage, true);
        }

        #endregion

        #endregion

    }
}
