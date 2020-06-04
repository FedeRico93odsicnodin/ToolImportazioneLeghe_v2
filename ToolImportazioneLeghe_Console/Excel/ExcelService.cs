using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel.Excel_Algorithms;
using ToolImportazioneLeghe_Console.Excel.Model_Excel;
using ToolImportazioneLeghe_Console.Messaging_Console;
using ToolImportazioneLeghe_Console.Utils;

namespace ToolImportazioneLeghe_Console.Excel
{
    /// <summary>
    /// Servizi excel per l'analisi del foglio 
    /// </summary>
    public class ExcelService
    {
        #region ATTRIBUTI PRIVATI

        /// <summary>
        /// Servizi di lettura excel 
        /// </summary>
        private ExcelReaders _excelReaders;


        /// <summary>
        /// Servizi di scrittura excel 
        /// </summary>
        private ExcelWriters _excelWriters;


        /// <summary>
        /// Nome per il file excel utilizzato come sorgente 
        /// </summary>
        private string _excelName;


        /// <summary>
        /// Stream contenuto del file excel aperto 
        /// </summary>
        private ExcelPackage _openedExcel;

        #endregion
        

        #region METODI PRIVATI

        /// <summary>
        /// Apertura del file excel in lettura / scrittura, viene passato anche il formato con il quale si identificheranno le informazioni
        /// sul file excel 
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="formatoExcel"></param>
        /// <param name="modalitaApertura"></param>
        /// <returns></returns>
        public bool OpenFileExcel(string excelPath, Constants.FormatFileExcel formatoExcel, Constants.ModalitaAperturaExcel modalitaApertura)
        {
            try
            {
                // validazione su formnato e path
                if (excelPath == String.Empty)
                    throw new Exception(ExceptionMessages.EXCEL_EMPTYPATH);

                // validazione sul formato in input
                if (formatoExcel == Constants.FormatFileExcel.NotDefined)
                    throw new Exception(ExceptionMessages.EXCEL_FORMATNOTDEFINED);

                if (modalitaApertura == Constants.ModalitaAperturaExcel.Lettura && !File.Exists(excelPath))
                    throw new Exception(ExceptionMessages.EXCEL_SOURCENOTEXISTING);
                else
                {
                    bool esistenza = false;

                    // ricreazione del file 
                    ServiceLocator.GetUtilityFunctions.BuildFilePath(excelPath, out esistenza);

                    if (esistenza)
                        ConsoleService.ConsoleExcel.EsistenzaFileExcel_Message(excelPath);

                }

                // set licenza corrente
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                _excelName = ServiceLocator.GetUtilityFunctions.GetFileName(excelPath);

                FileStream currentFileExcel = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                _openedExcel = new ExcelPackage(currentFileExcel);

                // decisione su cosa viene istanziato in base alla modalità su cui agire sul file excel 
                if (modalitaApertura == Constants.ModalitaAperturaExcel.Lettura)
                    _excelReaders = new ExcelReaders(ref _openedExcel, formatoExcel);
                else if (modalitaApertura == Constants.ModalitaAperturaExcel.Scrittura)
                    _excelWriters = new ExcelWriters(ref _openedExcel, formatoExcel);

                return true;
            }
            catch (Exception e)
            {
                throw new Exception(String.Format(ExceptionMessages.EXCEL_PROBLEMAAPERTURAFILE, ServiceLocator.GetUtilityFunctions.GetFileName(excelPath), e.Message));
            }
            
        }

        #endregion


        #region GETTERS READER WRITER EXCEL 

        /// <summary>
        /// Ottengo il reader per una apertura precedente di un file excel 
        /// </summary>
        public ExcelReaders GetExcelReaders => _excelReaders;


        /// <summary>
        /// Ottengo il writer per una apertura / eventuale creazione di un file excel 
        /// </summary>
        public ExcelWriters GetExcelWriters => _excelWriters;

        #endregion
    }


    /// <summary>
    /// Servizi di lettura per il foglio excel corrente 
    /// </summary>
    public class ExcelReaders
    {
        #region ATTRIBUTI PRIVATI

        /// <summary>
        /// Riferimento all'excel aperto dal servizio principale
        /// </summary>
        private ExcelPackage _openedExcel;


        /// <summary>
        /// Formato excel in apertura corrente 
        /// </summary>
        private Constants.FormatFileExcel _formatoExcel;


        /// <summary>
        /// Set di tutti i fogli in lettura corrente per il file excel aperto
        /// </summary>
        private List<Excel_AlloyInfo_Sheet> _fogliLetturaCorrente;
        
        #endregion


        #region COSTRUTTORE 

        /// <summary>
        /// Passaggio dello stream excel aperto con annesso anche il formato excel in apertura corrente 
        /// </summary>
        /// <param name="openedExcel"></param>
        /// <param name="formatoExcel"></param>
        public ExcelReaders(ref ExcelPackage openedExcel, Constants.FormatFileExcel formatoExcel)
        {
            _openedExcel = openedExcel;
            _formatoExcel = formatoExcel;
            
        }

        #endregion


        #region AZIONI SU FILE EXCEL 

        /// <summary>
        /// Riconoscimento delle informazioni contenute nei diversi fogli per il file excel corrente 
        /// se non riconosco almeno una tipologia di informazioni di lega ritorno false
        /// </summary>
        /// <returns></returns>
        public bool RecognizeSheetsOnExcel()
        {
            // inizializzazione lista di tutti i fogli letti
            _fogliLetturaCorrente = new List<Excel_AlloyInfo_Sheet>();


            #region RICONOSCIMENTO TIPOLOGIA PER DATABASE LEGHE 

            // inizializzazione di una delle 2 liste in base al formato
            if (_formatoExcel == Constants.FormatFileExcel.DatabaseLeghe)
            {
                foreach(ExcelWorksheet currentWorksheet in _openedExcel.Workbook.Worksheets)
                {
                    
                    
                    // riconoscimento della tipologia foglio per il primo formato
                    Constants_Excel.TipologiaFoglio_Format tipologiaRiconoscita = RecognizeTipoFoglio_Format1(currentWorksheet);
                    


                    if (!(tipologiaRiconoscita == Constants_Excel.TipologiaFoglio_Format.NotDefined))
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format1(currentWorksheet.Name, currentWorksheet.Index, tipologiaRiconoscita.ToString());
                    else
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_FoglioNonRiconosciuto(currentWorksheet.Name, currentWorksheet.Index);
                        
                }

                // se ho riconosciuto almeno un foglio per materiali e almeno uno per concentrazioni ritorno true, altrimenti false
                if (_fogliLetturaCorrente.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioLeghe).Count() > 0 &&
                    _fogliLetturaCorrente.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioConcentrazioni).Count() > 0)
                    return true;


                return false;

            }

            #endregion


            #region RICONOSCIMENTO TIPOLOGIA PER CLIENTE 

            else if(_formatoExcel == Constants.FormatFileExcel.Cliente)
            {
                foreach (ExcelWorksheet currentWorksheet in _openedExcel.Workbook.Worksheets)
                {
                    bool hoRiconosciutoSecondaTipologia = RecognizeTipoFoglio_Format2(currentWorksheet);
                    

                    if (hoRiconosciutoSecondaTipologia)
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format2(currentWorksheet.Name, currentWorksheet.Index);
                    else
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_FoglioNonRiconosciuto(currentWorksheet.Name, currentWorksheet.Index);
                    
                }
            }

            #endregion

            // come il caso precedente 
            if (_fogliLetturaCorrente.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioLegheConcentrazioni).Count() == 0)
                return false;

            return true;
        }


        /// <summary>
        /// Lettura delle informazioni contenute nel file excel corrente in base al formato inserito per il file excel corrente,
        /// in particolare in base alla tipologia si deciderà se recuperare le informazioni in un modo piuttosto che in un altro
        /// </summary>
        /// <returns></returns>
        public bool ReadExcelInformation()
        {

            #region RECUPERO INFORMAZIONI PER IL PRIMO FORMATO


            // fogli compilati con le nuove informazioni
            List<Excel_AlloyInfo_Sheet> newFilledSheets = new List<Excel_AlloyInfo_Sheet>();


            // caso in cui il file è di primo formato
            if (_formatoExcel == Constants.FormatFileExcel.DatabaseLeghe)
            {
                foreach (Excel_AlloyInfo_Sheet currentFoglioExcel in _fogliLetturaCorrente)
                {

                    // recupero del foglio corrente contenuto nel file di riferimento e dal quale continuare la lettura delle informazioni
                    ExcelWorksheet excelSheetReference = _openedExcel.Workbook.Worksheets[currentFoglioExcel.GetPosSheet];

                    // foglio relativo alle informazioni compilate 
                    Excel_AlloyInfo_Sheet filledInfo;

                    // riconoscimenti per la lettura del foglio delle leghe 
                    if (currentFoglioExcel.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioLeghe)
                    {

                        // esito di recupero informazioni per il foglio corrente 
                        Constants_Excel.EsitoRecuperoInformazioniFoglio esitoRecuperoInformazioniFoglioFormato1Corrente = ExcelReaderInfo.ReadLegheInfo(
                                excelSheetReference,
                                currentFoglioExcel,
                                out filledInfo);
                        

                        if ((esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto) ||
                            (esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings))
                        {
                            // segnalazione di avvenuta lettura per il foglio corrente a console
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioTipoFormato1AvvenutaCorrettamente(excelSheetReference.Name, currentFoglioExcel.GetPosSheet, currentFoglioExcel.GetTipologiaFoglio);

                            // segnalazione di eventuale recupero con warnings 
                            if (esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
                                ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConWarnings(excelSheetReference.Name);

                            // aggiunta del foglio con le nuove informazioni all'interno della lista finale dei fogli informazioni
                            newFilledSheets.Add(filledInfo);
                        }
                        // segnalazione che le informazioni di lega corrente sono state recuperate con degli errori
                        else
                        {
                            // oltre la segnalazione non c'è bisogno di aggiungere il foglio al nuovo set in quanto non piu valido
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConErrori(excelSheetReference.Name, currentFoglioExcel.GetPosSheet, currentFoglioExcel.GetTipologiaFoglio);
                        }
                    }
                    // riconoscimento per la lettura del foglio delle concentrazioni
                    else if (currentFoglioExcel.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioConcentrazioni)
                    {
                        

                        // esito di recupero informazioni per il foglio corrente 
                        Constants_Excel.EsitoRecuperoInformazioniFoglio esitoRecuperoInformazioniFoglioFormato1Corrente = ExcelReaderInfo.ReadConcentrationsInfo(
                            excelSheetReference,
                            currentFoglioExcel,
                            out filledInfo);

                        // nel caso in cui ho recuperato il foglio ma con degli errori allora non vado a leggerne ulteriormente le informazioni
                        if (!(esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori))
                        {
                            // segnalazione di avvenuta lettura per il foglio corrente a console
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioTipoFormato1AvvenutaCorrettamente(excelSheetReference.Name, currentFoglioExcel.GetPosSheet, currentFoglioExcel.GetTipologiaFoglio);

                            // segnalazione di eventuale recupero con warnings 
                            if (esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
                                ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConWarnings(excelSheetReference.Name);

                            // aggiunta del foglio con le nuove informazioni all'interno della lista finale dei fogli informazioni
                            newFilledSheets.Add(filledInfo);
                        }
                        // segnalazione che le informazioni di concentrazione corrente sono state recuperate con degli errori
                        else
                        {
                            // oltre la segnalazione non c'è bisogno di aggiungere il foglio al nuovo set in quanto non piu valido
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConErrori(excelSheetReference.Name, currentFoglioExcel.GetPosSheet, currentFoglioExcel.GetTipologiaFoglio);
                        }
                    }
                }


                // fine iterazione corrente: vado ad eseguire il controllo da quanti fogli ho recuperato effettivamente informazioni utili 
                // se il numero dei fogli per i quali le informazioni non sono state recuperate correttamente è 0, non posso proseguire con l'analisi
                if (newFilledSheets.Count() == 0)
                    return false;



                // tra il set di fogli recuperati deve essercene almeno uno di materiali e uno di concentrazioni per poter continuare correttamente l'analisi
                // TODO: mettere questo controllo anche in fase di validazione per i fogli sul formato corrente 
                if (
                    newFilledSheets.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioLeghe).Count() == 0 ||
                    newFilledSheets.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format.FoglioConcentrazioni).Count() == 0)
                {
                    // segnalazione di prossima interruzione del tool per mancanza di tutti i fogli necessari
                    ConsoleService.ConsoleExcel.ExcelReaders_Message_InterruzioneAnalisi_SetFogliRiconosciutiInsufficiente(ServiceLocator.GetUtilityFunctions.GetFileName(Constants.ExcelSourcePath));

                    return false;
                }

                // attribuzione del nuovo valore per i files recuperati per il formato corrente 
                _fogliLetturaCorrente = newFilledSheets;

                // per gli altri casi ritorno true in quanto il recupero si è concluso correttamente 
                return true;
            }

            #endregion


            #region RECUPERO INFORMAZIONI DAL SECONDO FORMATO

            // caso in cui il file è di secondo formato
            else if (_formatoExcel == Constants.FormatFileExcel.Cliente)
            {

                // introduzione della variabile per inserire i fogli 2 con le informazioni correntemente recuperate per questo step
                List<Excel_AlloyInfo_Sheet> newValuesFormat2 = new List<Excel_AlloyInfo_Sheet>();

                foreach (Excel_AlloyInfo_Sheet currentFoglioExcel in _fogliLetturaCorrente)
                {
                    
                    // dichiarazione delle possibili stringhe di errore / warnings per il caso corrente per la seconda tipologia in analisi
                    string errorsMessages_Format2 = String.Empty;
                    string warningMessages_Format2 = String.Empty;


                    // recupero del foglio corrente contenuto nel file di riferimento e dal quale continuare la lettura delle informazioni
                    ExcelWorksheet excelSheetReference = _openedExcel.Workbook.Worksheets[currentFoglioExcel.GetPosSheet];

                    // oggetto nel quale inserisco le informazioni recuperate per il foglio corrente
                    Excel_AlloyInfo_Sheet filledInfo;

                    // recupero informazioni per il foglio di formato 2
                    Constants_Excel.EsitoRecuperoInformazioniFoglio esitoRecuperoInformazioniFoglioFormato2Corrente = ExcelReaderInfo.ReadInfoFormat2(excelSheetReference, currentFoglioExcel
                        ,out filledInfo);
                    

                    if (esitoRecuperoInformazioniFoglioFormato2Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto ||
                        esitoRecuperoInformazioniFoglioFormato2Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
                    {
                        // segnalazione di recupero corretto per le informazioni contenute in questo secondo formato
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioTipoFormato2AvvenutaCorrettamente(excelSheetReference.Name, currentFoglioExcel.GetPosSheet);

                        // segnalazione di eventuale recupero con warnings
                        if (esitoRecuperoInformazioniFoglioFormato2Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato2AvvenutaConWarnings(excelSheetReference.Name);

                        // inserimento del foglio con le informazioni correntemente recuperate all'interno del set attuale
                        newValuesFormat2.Add(filledInfo);

                    }
                    else
                    {
                        // segnalazione di recupero per il foglio corrente con degli errori per i quali non se ne potrà proseguire la successiva analisi
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato2AvvenutaConErrori(excelSheetReference.Name, currentFoglioExcel.GetPosSheet);
                    }

                }


                // controllo dopo iterazione su questi fogli 
                if (newValuesFormat2.Count() == 0)
                    return false;

                _fogliLetturaCorrente = newValuesFormat2;

                return true;
            }

            #endregion

            // se esco dalle 2 validazioni disponibili per i diversi fogli senza ancora aver recuperato niente e non aver riconosciuto tipologia foglio lancio una eccezione 
            else
                throw new Exception(ExceptionMessages.EXCEL_READERINFO_NESSUNAATTRIBUZIONETIPIRICONOSCIUTAAFOGLI);
            

        }
        

        /// <summary>
        /// Vegnono validate le informazioni presenti all'interno del foglio excel corrente in particolare 
        /// 1) le informazioni possono essere completamente corrette (valida la definizione degli elementi per i quadranti e dei diversi elementi di stringa e numerici)
        /// 2) le informazioni possono essere state inserite parzialmente corrette (alcuni quadranti / informazioni di lega sono stati scartati perché non validi)
        /// 3) le informazioni sono state trovate completamente scorrette 
        /// 
        /// in base alla decisione effettuata nel momento delle configurazioni, se ci si trova nel caso (2) il tool deciderà in automatico se continuare l'analisi o fermare 
        /// l'esecuzione 
        /// </summary>
        /// <returns></returns>
        public Constants_Excel.ValidazioneFoglio ValidateExcel()
        {
            return Constants_Excel.ValidazioneFoglio.NessunaCorrispondenza;
        }

        #endregion


        #region METODI PRIVATI - RICONOSCIMENTO TIPOLOGIA FOGLIO

        /// <summary>
        /// Riconoscimento di una delle 2 tipologie per il formato 1 di fogli presenti nel file excel 
        /// quindi si puo trattare di un foglio di informazioni di lega o di concentrazioni per queste 
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="listaParametriRiconosciuti"></param>
        /// <returns></returns>
        private Constants_Excel.TipologiaFoglio_Format RecognizeTipoFoglio_Format1(ExcelWorksheet currentSheet)
        {
            // inizializzazione del foglio in lettura corrente 
            Excel_AlloyInfo_Sheet recognizedInfoOnSheet;

            // tentativo di lettura delle informazioni relativamente al formato di lega 
            Constants_Excel.EsitoRecuperoInformazioniFoglio riconoscimentoLeghe = ExcelRecognizers.Recognize_Format1_InfoLeghe(ref currentSheet, out recognizedInfoOnSheet);

            // ho riconosciuto correttamente il foglio come foglio di leghe 
            if(riconoscimentoLeghe == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings || 
                riconoscimentoLeghe == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto)
            {
                // attribuzione tipologia lega 
                recognizedInfoOnSheet.GetTipologiaFoglio = Constants_Excel.TipologiaFoglio_Format.FoglioLeghe;
                
                _fogliLetturaCorrente.Add(recognizedInfoOnSheet);
                
                return Constants_Excel.TipologiaFoglio_Format.FoglioLeghe;
            }


            // se mi trovo qui è perché non sono ancora riuscito a distinguere per il foglio corrente 
            Constants_Excel.EsitoRecuperoInformazioniFoglio riconoscimentoConcentrazioni = ExcelRecognizers.Recognize_Format1_InfoConcentrations(ref currentSheet, out recognizedInfoOnSheet);

            if(riconoscimentoConcentrazioni == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings ||
                riconoscimentoConcentrazioni == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto)
            {
                // attribuzione tipologia concentrazioni
                recognizedInfoOnSheet.GetTipologiaFoglio = Constants_Excel.TipologiaFoglio_Format.FoglioConcentrazioni;
                _fogliLetturaCorrente.Add(recognizedInfoOnSheet);

                return Constants_Excel.TipologiaFoglio_Format.FoglioConcentrazioni;
            }

            // attribuzione tipologia indefinita
            recognizedInfoOnSheet.GetTipologiaFoglio = Constants_Excel.TipologiaFoglio_Format.NotDefined;
            // nel caso in cui mi trovassi a non avere nessun riconoscimento per entrambi i formati il foglio non viene preso in considerazione ma solamente
            // i diversi errori / warnings di generazione durante l'analisi
            return Constants_Excel.TipologiaFoglio_Format.NotDefined;
        }


        /// <summary>
        /// Riconoscimento se il formato usato per la seconda tipologia per il foglio corrente è effettivamente valida per 
        /// il riconoscimento del foglio corrente come formato 2
        /// Viene quindi restituito l'array dei parametri che viene eventualmente riconosciuto dall'analisi
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="listaParametriRiconosciuti"></param>
        /// <returns></returns>
        private bool RecognizeTipoFoglio_Format2(ExcelWorksheet currentSheet)
        {
            // inizializzazione per il foglio corrente 
            Excel_AlloyInfo_Sheet RecognizedInfo;


            // analisi per il foglio corrente 
            Constants_Excel.EsitoRecuperoInformazioniFoglio esitoLetturaFoglioSecondoFormato = ExcelRecognizers.Recognize_Format2_InfoLegheConcentrazioni(
                ref currentSheet,
                out RecognizedInfo);
            
            if(esitoLetturaFoglioSecondoFormato == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings ||
                esitoLetturaFoglioSecondoFormato == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto)
            {
                // fisso la categoria per il foglio corrente e ritorno
                RecognizedInfo.GetTipologiaFoglio = Constants_Excel.TipologiaFoglio_Format.FoglioLegheConcentrazioni;

                // aggiunta del foglio corrente 
                _fogliLetturaCorrente.Add(RecognizedInfo);

                return true;
            }
            
            return false;
        }

        #endregion


        #region METODI PRIVATI - RECUPERO DI TUTTE LE INFORMAZIONI PRESENTI SUL FILE IN BASE ALLA TIPOLOGIA

        /// <summary>
        /// Recupero di tutte le informazioni dal file excel di FORMATO 1 per un foglio relativo alle leghe
        /// </summary>
        /// <returns></returns>
        private static bool ReadInfoFromExcelFormat_1_Leghe()
        {
            return false;
        }


        /// <summary>
        /// Recupero di tutte le informazioni del file excel di FORMATO1 per un foglio relativo alle concentrazioni
        /// </summary>
        /// <returns></returns>
        private static bool ReadInfoFromExcelFormat_1_Concentrations()
        {
            return false;
        }


        /// <summary>
        /// Recupero di tutte le informazioni del file excel di FORMATO2 per un foglio relativo sia a leghe che concentrazioni
        /// </summary>
        /// <returns></returns>
        private static bool ReadInfoFromExcelFormat_2()
        {
            return false;
        }

        #endregion
        

        #region SCRITTURA MESSAGGI ERRORE - WARNINGS SU APPOSITI LOG DI FOGLIO 

        /// <summary>
        /// Permette di verificare se le stringhe relative ai messaggi di errori / warnings per il foglio excel corrente e per lo step corrente sono 
        /// effettivamente valorizzate durante l'analisi e in base a questo le va a scrivere opportunamente sul file da ricreare o gia esistente e per 
        /// la folder relativa all'inserimento di tutti i messaggi per il foglio correntemente in analisi
        /// </summary>
        /// <param name="currentErrorsMessages"></param>
        /// <param name="currentWarningMessages"></param>
        /// <param name="currentSheetName"></param>
        private void CheckMessagesForCurrentSheet(string currentErrorsMessages, string currentWarningMessages, string currentSheetName, Constants_Excel.StepLetturaFoglio currentStep)
        {
            if (currentErrorsMessages != String.Empty)
                InsertErrorMessages(currentErrorsMessages, currentSheetName, currentStep);

            if (currentWarningMessages != String.Empty)
                InsertWarningsMessages(currentWarningMessages, currentSheetName, currentStep);
        } 


        /// <summary>
        /// Permette di inserire per messaggi di errore nell'apposito log
        /// </summary>
        /// <param name="currentErrorsMessages"></param>
        /// <param name="currentSheetName"></param>
        /// <param name="currentStep"></param>
        private void InsertErrorMessages(string currentErrorsMessages, string currentSheetName, Constants_Excel.StepLetturaFoglio currentStep)
        {
            // TODO : implementazione scrittura su log
        }


        /// <summary>
        /// Permette di inserire messaggi di warnings nell'apposito log
        /// </summary>
        /// <param name="currentWarningsMessages"></param>
        /// <param name="currentSheetName"></param>
        /// <param name="currentStep"></param>
        private void InsertWarningsMessages(string currentWarningsMessages, string currentSheetName, Constants_Excel.StepLetturaFoglio currentStep)
        {
            // TODO : implementazione scrittura su log
        }

        #endregion

    }


    /// <summary>
    /// Servizi di scrittura per il foglio excel corrente 
    /// </summary>
    public class ExcelWriters
    {
        #region ATTRIBUTI PRIVATI

        /// <summary>
        /// Riferimento all'excel aperto dal servizio principale
        /// </summary>
        private ExcelPackage _openedExcel;


        /// <summary>
        /// Formato excel in apertura corrente 
        /// </summary>
        private Constants.FormatFileExcel _formatoExcel;

        
        /// <summary>
        /// Lista di tutti i fogli excel che vengono eventualmente letti dalla seconda tipologia di formato 
        /// per il foglio excel
        /// </summary>
        private List<Excel_AlloyInfo_Sheet> _sheetsLetturaFormat_2;

        #endregion


        #region COSTRUTTORE

        /// <summary>
        /// Passaggio dello stream excel aperto con annesso anche il formato excel in scrittura corrente 
        /// </summary>
        /// <param name="openedExcel"></param>
        /// <param name="formatoExcel"></param>
        public ExcelWriters(ref ExcelPackage openedExcel, Constants.FormatFileExcel formatoExcel)
        {
            _openedExcel = openedExcel;
            _formatoExcel = formatoExcel;
        }

        #endregion


        #region SETTERS - INSERIMENTO DELLE LISTE DI INFORMAZIONI EVENTUALMENTE DA SCRIVERE PER UNA DELLE 2 TIPOLOGIE
        
        /// <summary>
        /// Carico tutte le informazioni prima di passare alla scrittura 
        /// </summary>
        public List<Excel_AlloyInfo_Sheet> SetExcelFormat_2Info { set { _sheetsLetturaFormat_2 = value; } }

        #endregion
    }
}
