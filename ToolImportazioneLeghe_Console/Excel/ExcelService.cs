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
        /// Lista di tutti i fogli excel che vengono eventualmente letti dalla prima tipologia di formato 
        /// per il foglio excel
        /// </summary>
        private List<Excel_Format1_Sheet> _sheetsLetturaFormat_1;


        /// <summary>
        /// Lista di tutti i fogli excel che vengono eventualmente letti dalla seconda tipologia di formato 
        /// per il foglio excel
        /// </summary>
        private List<Excel_Format2_Sheet> _sheetsLetturaFormat_2;


        /// <summary>
        /// Traccia delle concentrazioni in lettura corrente per il foglio delle concentrazioni e la prima tipologia
        /// di formato excel
        /// </summary>
        private List<Excel_Format1_Sheet2_ConcQuadrant> _quadrantiConcentrazioneLetturaCorrente;


        /// <summary>
        /// Inizio lettura leghe all'interno del primo foglio di leghe per il primo formato excel disponibile
        /// </summary>
        private int _startingPosLeghe_row_format1 = 0;


        /// <summary>
        /// Inizio lettura leghe all'interno del primo foglio di leghe per il primo formato excel disponibile
        /// </summary>
        private int _startingPosLeghe_col_format1 = 0;


        /// <summary>
        /// Fine lettura delle colonne per le proprieta di lega per il foglio riconosciuto come contenente informazioni di lega
        /// </summary>
        private int _endingPosLeghe_col_format1 = 0;


        /// <summary>
        /// Traccia della lettura delle concentrazioni per il secondo formato excel disponibile
        /// </summary>
        private List<Excel_Format2_Row_LegaProperties> _colonneConcentrazioniSecondoFormato;


        /// <summary>
        /// Inzio riga di lettura leghe per il secondo formato excel  
        /// </summary>
        private int _startingPosLeghe_row_format2 = 0;


        /// <summary>
        /// Inizio colonna lettura leghe per il secondo formato excel
        /// </summary>
        private int _startingPosLeghe_col_format2 = 0;


        /// <summary>
        /// Possibili messaggi di errori dati dalla lettura per il foglio excel corrente 
        /// questo messaggio di errore viene valorizzato ogni qual volta di analizza un foglio 1 per il primo formato 
        /// e il riconoscimento non avviene correttamente 
        /// </summary>
        private string _possibleErrors_Leghe_ReadCurrentExcel = String.Empty;


        /// <summary>
        /// Possibili messaggi di warnings dati dalla lettura per il foglio excel corrente 
        /// questo messaggio di warning viene valorizzato ogni qual volta si analizza un foglio 1 per il primo formato
        /// e il riconoscimento avviene correttamente ma con alcuni warnings
        /// </summary>
        private string _possibleWarnings_Leghe_ReadCurrentExcel = String.Empty;


        /// <summary>
        /// Possibili messaggi di errore dati durante la lettura per il foglio excel corrente 
        /// questo messaggio di errore viene valorizzato ogni qual volta si analizza un foglio 2 per il primo formato
        /// e il riconoscimento non avviene correttamente 
        /// </summary>
        private string _possibleErrors_Concentrations_ReadCurrentExcel = String.Empty;


        /// <summary>
        /// Possibili messaggi di warnings dati durante la lettura per il foglio excel corrente 
        /// questo messaggio di warning viene valorizzato ogni qual volta si analizza un foglio 2 per il primo formato
        /// e il riconoscimento non avviene correttamente 
        /// </summary>
        private string _possibleWarnings_Concentrations_ReadCurrentExcel = String.Empty;


        /// <summary>
        /// Possibili errori derivanti dall'analisi e lettura per i dati per il secondo formato excel disponibile 
        /// per le informazioni di lega 
        /// </summary>
        private string _possibleErrors_LetturaSecondoFormato = String.Empty;


        /// <summary>
        /// Possibili warnings derivanti dall'analisi e lettura per i dati per il secondo formato excel disponibile
        /// per le informazioni di lega
        /// </summary>
        private string _possibleWarnings_LetturaSecondoFormato = String.Empty;

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
            if (_openedExcel == null)
                throw new Exception(ExceptionMessages.EXCEL_FILENOTINMEMORY);
            
            // segnalazione della posizione per il file excel corrente
            int currentSheetPosition = 0;


            #region RICONOSCIMENTO TIPOLOGIA PER DATABASE LEGHE 

            // inizializzazione di una delle 2 liste in base al formato
            if (_formatoExcel == Constants.FormatFileExcel.DatabaseLeghe)
            {
                _sheetsLetturaFormat_1 = new List<Excel_Format1_Sheet>();
                

                foreach(ExcelWorksheet currentWorksheet in _openedExcel.Workbook.Worksheets)
                {
                    currentSheetPosition++;
                    
                    // riconoscimento della tipologia foglio per il primo formato
                    Constants_Excel.TipologiaFoglio_Format1 tipologiaRiconoscita = RecognizeTipoFoglio_Format1(currentWorksheet);


                    // verifica dei messaggi di errore / warnings per la lettura corrente e nuovo azzeramento 
                    CheckMessagesForCurrentSheet(_possibleErrors_Leghe_ReadCurrentExcel, _possibleWarnings_Leghe_ReadCurrentExcel, currentWorksheet.Name, Constants_Excel.StepLetturaFoglio.Riconoscimento);
                    CheckMessagesForCurrentSheet(_possibleErrors_Concentrations_ReadCurrentExcel, _possibleWarnings_Concentrations_ReadCurrentExcel, currentWorksheet.Name, Constants_Excel.StepLetturaFoglio.Riconoscimento);

                    _possibleErrors_Leghe_ReadCurrentExcel = String.Empty;
                    _possibleErrors_Concentrations_ReadCurrentExcel = String.Empty;
                    _possibleWarnings_Leghe_ReadCurrentExcel = String.Empty;
                    _possibleWarnings_Concentrations_ReadCurrentExcel = String.Empty;


                    if (!(tipologiaRiconoscita == Constants_Excel.TipologiaFoglio_Format1.NotDefined))
                    {
                        // segnalazione a console per la tipologia riconosciuta 
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format1(currentWorksheet.Name, currentSheetPosition, tipologiaRiconoscita.ToString());

                        Excel_Format1_Sheet foglioExcelCorrenteInfo = new Excel_Format1_Sheet(currentWorksheet.Name, tipologiaRiconoscita, currentSheetPosition);

                        // vedo se inserire per posizione iniziale per la lettura delle leghe oppure per i quadranti riconosciuti
                        if(tipologiaRiconoscita == Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe)
                        {
                            foglioExcelCorrenteInfo.StartingRow_letturaLeghe = _startingPosLeghe_row_format1;
                            foglioExcelCorrenteInfo.StartingCol_letturaLeghe = _startingPosLeghe_col_format1;
                            foglioExcelCorrenteInfo.EndingCol_letturaLeghe = _endingPosLeghe_col_format1;

                            // azzeramento proprieta per eventuale prossima lettura 
                            _startingPosLeghe_col_format1 = 0;
                            _startingPosLeghe_row_format1 = 0;
                            _endingPosLeghe_col_format1 = 0;
                        }
                        else if(tipologiaRiconoscita == Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni)
                        {
                            foglioExcelCorrenteInfo.GetConcQuadrants_Type2 = _quadrantiConcentrazioneLetturaCorrente;

                            // azzeramento proprieta per eventuale prossima lettura 
                            _quadrantiConcentrazioneLetturaCorrente = null;
                        }

                        // aggiunta del foglio corrente 
                        _sheetsLetturaFormat_1.Add(foglioExcelCorrenteInfo);
                    }
                    else
                    {
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_FoglioNonRiconosciuto(currentWorksheet.Name, currentSheetPosition);
                    }
                        
                }

                // non è presente nessun foglio sul quale eseguire la lettura delle informazioni
                if (_sheetsLetturaFormat_1.Count() == 0)
                    return false;

                // TODO: capire se discriminare anche a questo livello le informazioni 
                // (ad esempio aggiungendo una variante per la quale ci deve essere almeno un match per foglio concentrazioni / materiali)
                return true;

            }

            #endregion


            #region RICONOSCIMENTO TIPOLOGIA PER CLIENTE 

            else if(_formatoExcel == Constants.FormatFileExcel.Cliente)
            {
                _sheetsLetturaFormat_2 = new List<Excel_Format2_Sheet>();

                foreach (ExcelWorksheet currentWorksheet in _openedExcel.Workbook.Worksheets)
                {
                    currentSheetPosition++;

                    bool hoRiconosciutoSecondaTipologia = RecognizeTipoFoglio_Format2(currentWorksheet);


                    // verifica dei messaggi di errore / warnings per la lettura corrente e nuovo azzeramento 
                    CheckMessagesForCurrentSheet(_possibleErrors_LetturaSecondoFormato, _possibleWarnings_LetturaSecondoFormato, currentWorksheet.Name, Constants_Excel.StepLetturaFoglio.Riconoscimento);

                    _possibleErrors_LetturaSecondoFormato = String.Empty;
                    _possibleWarnings_LetturaSecondoFormato = String.Empty;



                    if (hoRiconosciutoSecondaTipologia)
                    {
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format2(currentWorksheet.Name, currentSheetPosition);

                        Excel_Format2_Sheet foglioExcelCorrenteInfo = new Excel_Format2_Sheet(currentWorksheet.Name, currentSheetPosition);

                        foglioExcelCorrenteInfo.StartingRow_Leghe = _startingPosLeghe_row_format2;
                        foglioExcelCorrenteInfo.StartingCol_Leghe = _startingPosLeghe_col_format2;
                        foglioExcelCorrenteInfo.AllInfoLeghe = _colonneConcentrazioniSecondoFormato;

                        _sheetsLetturaFormat_2.Add(foglioExcelCorrenteInfo);
                    }
                    else
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_FoglioNonRiconosciuto(currentWorksheet.Name, currentSheetPosition);
                }
            }

            #endregion

            // come il caso precedente 
            if (_sheetsLetturaFormat_2.Count() == 0)
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
            if (_openedExcel == null)
                throw new Exception(ExceptionMessages.EXCEL_FILENOTINMEMORY);


            #region RECUPERO INFORMAZIONI PER IL PRIMO FORMATO

            // caso in cui il file è di primo formato
            if (_formatoExcel == Constants.FormatFileExcel.DatabaseLeghe)
            {
                // eccezione nel caso in cui la lista relativa ai fogli riconosciuti per il primo formato sia NULL EMPTY
                if (_sheetsLetturaFormat_1 == null)
                    throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT1);
                if (_sheetsLetturaFormat_1.Count() == 0)
                    throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT1);

                // nuova variabile contenimento informazioni riempite per questo step
                List<Excel_Format1_Sheet> excelSheetsNewValues_Format1 = new List<Excel_Format1_Sheet>();


                foreach (Excel_Format1_Sheet currentFoglioExcel in _sheetsLetturaFormat_1)
                {
                    // istanza eventuali messaggi errore warnings per il foglio corrente in fase di recupero informazioni
                    string errorMessages = String.Empty;
                    string warningMessages = String.Empty;


                    // eccezione su validazione foglio corrente eventualmente mancata 
                    if (currentFoglioExcel.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format1.NotDefined)
                        throw new Exception(ExceptionMessages.EXCEL_READERINFO_TIPOLOGIANONDEFINITAFOGLIOCORRENTE);

                    if (currentFoglioExcel.GetPosSheet == 0)
                        throw new Exception(ExceptionMessages.EXCEL_READERINFO_NESSUNAPOSIZIONETROVATAPERFOGLIOCORRENTE);


                    // recupero del foglio corrente contenuto nel file di riferimento e dal quale continuare la lettura delle informazioni
                    ExcelWorksheet excelSheetReference = _openedExcel.Workbook.Worksheets[currentFoglioExcel.GetPosSheet];

                    // oggetto nel quale inserisco le informazioni recuperate per il foglio corrente
                    Excel_Format1_Sheet filledInfo;

                    // esito di recupero informazioni per il foglio corrente 
                    Constants_Excel.EsitoRecuperoInformazioniFoglio esitoRecuperoInformazioniFoglioFormato1Corrente = ExcelReaderInfo.ReadLegheInfo(
                            excelSheetReference,
                            currentFoglioExcel,
                            out filledInfo,
                            out warningMessages,
                            out errorMessages);

                    // scrittura eventuale il file di log per errori e wornings su recupero formato corrente 
                    CheckMessagesForCurrentSheet(errorMessages, warningMessages, excelSheetReference.Name, Constants_Excel.StepLetturaFoglio.RecuperoInformazioni_Validazione1);


                    // riconoscimenti per la lettura del foglio delle leghe 
                    if (currentFoglioExcel.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe)
                    {
                        if ((esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto) ||
                            (esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings))
                        {
                            // segnalazione di avvenuta lettura per il foglio corrente a console
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioTipoFormato1AvvenutaCorrettamente(excelSheetReference.Name, currentFoglioExcel.GetPosSheet, currentFoglioExcel.GetTipologiaFoglio);

                            // segnalazione di eventuale recupero con warnings 
                            if (esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
                                ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConWarnings(excelSheetReference.Name);

                            // aggiunta del foglio con le nuove informazioni all'interno della lista finale dei fogli informazioni
                            excelSheetsNewValues_Format1.Add(filledInfo);
                        }
                        // segnalazione che le informazioni di lega corrente sono state recuperate con degli errori
                        else
                        {
                            // oltre la segnalazione non c'è bisogno di aggiungere il foglio al nuovo set in quanto non piu valido
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConErrori(excelSheetReference.Name, currentFoglioExcel.GetPosSheet, currentFoglioExcel.GetTipologiaFoglio);
                        }
                    }
                    // riconoscimento per la lettura del foglio delle concentrazioni
                    else if (currentFoglioExcel.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni)
                    {
                        // dichiarazione della variabile per il contenimento delle informazioni per le concentrazioni lette dal foglio excel corrente 
                        Excel_Format1_Sheet foglioExcelConcentrations;

                        // nel caso in cui ho recuperato il foglio ma con degli errori allora non vado a leggerne ulteriormente le informazioni
                        if (!(ExcelReaderInfo.ReadConcentrationsInfo(
                            excelSheetReference,
                            currentFoglioExcel,
                            out foglioExcelConcentrations,
                            out warningMessages,
                            out errorMessages) == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConErrori))
                        {
                            // segnalazione di avvenuta lettura per il foglio corrente a console
                            ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioTipoFormato1AvvenutaCorrettamente(excelSheetReference.Name, currentFoglioExcel.GetPosSheet, currentFoglioExcel.GetTipologiaFoglio);

                            // segnalazione di eventuale recupero con warnings 
                            if (esitoRecuperoInformazioniFoglioFormato1Corrente == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
                                ConsoleService.ConsoleExcel.ExcelReaders_Message_LetturaFoglioFormato1AvvenutaConWarnings(excelSheetReference.Name);

                            // aggiunta del foglio con le nuove informazioni all'interno della lista finale dei fogli informazioni
                            excelSheetsNewValues_Format1.Add(filledInfo);
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
                if (excelSheetsNewValues_Format1.Count() == 0)
                    return false;

                // tra il set di fogli recuperati deve essercene almeno uno di materiali e uno di concentrazioni per poter continuare correttamente l'analisi
                // TODO: mettere questo controllo anche in fase di validazione per i fogli sul formato corrente 
                if (
                    excelSheetsNewValues_Format1.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe).Count() == 0 ||
                    excelSheetsNewValues_Format1.Where(x => x.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni).Count() == 0)
                {
                    // segnalazione di prossima interruzione del tool per mancanza di tutti i fogli necessari
                    ConsoleService.ConsoleExcel.ExcelReaders_Message_InterruzioneAnalisi_SetFogliRiconosciutiInsufficiente(ServiceLocator.GetUtilityFunctions.GetFileName(Constants.ExcelSourcePath));

                    return false;
                }

                // attribuzione del nuovo valore per i files recuperati per il formato corrente 
                _sheetsLetturaFormat_1 = excelSheetsNewValues_Format1;

                // per gli altri casi ritorno true in quanto il recupero si è concluso correttamente 
                return true;
            }

            #endregion


            #region RECUPERO INFORMAZIONI DAL SECONDO FORMATO

            // caso in cui il file è di secondo formato
            else if (_formatoExcel == Constants.FormatFileExcel.Cliente)
            {
                // eccezione nel caso in cui la lista relativa ai fogli riconosciuti per il primo formato sia NULL EMPTY
                if (_sheetsLetturaFormat_2 == null)
                    throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT2);
                if (_sheetsLetturaFormat_2.Count() == 0)
                    throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT2);

                // introduzione della variabile per inserire i fogli 2 con le informazioni correntemente recuperate per questo step
                List<Excel_Format2_Sheet> newValuesFormat2 = new List<Excel_Format2_Sheet>();

                foreach (Excel_Format2_Sheet currentFoglioExcel in _sheetsLetturaFormat_2)
                {

                    if (currentFoglioExcel.GetPosSheet == 0)
                        throw new Exception(ExceptionMessages.EXCEL_READERINFO_NESSUNAPOSIZIONETROVATAPERFOGLIOCORRENTE);

                    // dichiarazione delle possibili stringhe di errore / warnings per il caso corrente per la seconda tipologia in analisi
                    string errorsMessages_Format2 = String.Empty;
                    string warningMessages_Format2 = String.Empty;


                    // recupero del foglio corrente contenuto nel file di riferimento e dal quale continuare la lettura delle informazioni
                    ExcelWorksheet excelSheetReference = _openedExcel.Workbook.Worksheets[currentFoglioExcel.GetPosSheet];

                    // oggetto nel quale inserisco le informazioni recuperate per il foglio corrente
                    Excel_Format2_Sheet filledInfo;

                    // recupero informazioni per il foglio di formato 2
                    Constants_Excel.EsitoRecuperoInformazioniFoglio esitoRecuperoInformazioniFoglioFormato2Corrente = ExcelReaderInfo.ReadInfoFormat2(excelSheetReference, currentFoglioExcel, out filledInfo, out errorsMessages_Format2, out warningMessages_Format2
                        );

                    // scrittura eventuale il file di log per errori e wornings su recupero formato corrente 
                    CheckMessagesForCurrentSheet(errorsMessages_Format2, warningMessages_Format2, excelSheetReference.Name, Constants_Excel.StepLetturaFoglio.RecuperoInformazioni_Validazione1);


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
        private Constants_Excel.TipologiaFoglio_Format1 RecognizeTipoFoglio_Format1(ExcelWorksheet currentSheet)
        {
            int startingRow = 0;
            int startingCol = 0;
            int endingColIndex = 0;

            #region POSSIBILE RICONOSCIMENTO DI UN FOGLIO PER LE LEGHE 

            // inizializzazione dei possibili messaggi di errore / warnings emersi durante la lettura per il file excel corrente 
            string currentErrorsExcel = String.Empty;
            string currentWarningsExcel = String.Empty;


            Constants_Excel.EsitoRecuperoInformazioniFoglio esitoRiconoscimentoFoglio_Leghe = ExcelRecognizers.Recognize_Format1_InfoLeghe(
                ref currentSheet, 
                out startingRow, 
                out startingCol, 
                out endingColIndex,
                out currentErrorsExcel,
                out currentWarningsExcel);



            // inserisco la stringa con gli eventuali warnings emersi dalla lettura corrente 
            _possibleErrors_Leghe_ReadCurrentExcel = currentErrorsExcel;
            _possibleWarnings_Leghe_ReadCurrentExcel = currentWarningsExcel;


            // riconoscimento completo e senza nessuna segnalazione per il foglio leghe corrente 
            if (esitoRiconoscimentoFoglio_Leghe == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto ||
               esitoRiconoscimentoFoglio_Leghe == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
            {
                // attribuzione parametri privati di riga / colonna di lettura per le leghe
                _startingPosLeghe_row_format1 = startingRow;
                _startingPosLeghe_col_format1 = startingCol;
                _endingPosLeghe_col_format1 = endingColIndex;

                return Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe;
            }

            #endregion


            #region POSSIBILE RICONOSCIMENTO DI UN FOGLIO PER LE CONCENTRAZIONI

            // lista che conterrà i quadranti di concentrazioni finali per la lettura di un foglio di seconda tipologia 
            List<Excel_Format1_Sheet2_ConcQuadrant> concentrationsQuadrants;

            // azzeramento delle stringhe per i messaggi errori / warnings reincontrati
            currentErrorsExcel = String.Empty;
            currentWarningsExcel = String.Empty;


            // tentativo di riconoscimento foglio concentrazioni
            Constants_Excel.EsitoRecuperoInformazioniFoglio esitoRecuperoInformazioniFoglio_Concentrazioni = ExcelRecognizers.Recognize_Format1_InfoConcentrations(
                ref currentSheet, 
                out concentrationsQuadrants,
                out currentErrorsExcel,
                out currentWarningsExcel);


            // attribuzione dei messaggi di warnings e di errori finali durante il tentativo di riconoscimento di un foglio per le concentrazioni
            _possibleErrors_Concentrations_ReadCurrentExcel = currentErrorsExcel;
            _possibleWarnings_Concentrations_ReadCurrentExcel = currentWarningsExcel;


            // sono riuscito a riconoscere il foglio per le concentrazioni correnti
            if (esitoRecuperoInformazioniFoglio_Concentrazioni == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto ||
                esitoRecuperoInformazioniFoglio_Concentrazioni == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
            {
                _quadrantiConcentrazioneLetturaCorrente = concentrationsQuadrants;
                return Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni;
            }

            #endregion

            // nel caso in cui mi trovassi a non avere nessun riconoscimento per entrambi i formati il foglio non viene preso in considerazione ma solamente
            // i diversi errori / warnings di generazione durante l'analisi
            return Constants_Excel.TipologiaFoglio_Format1.NotDefined;
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
            int startingRow = 0;
            int startingCol = 0;

            // istanze messaggi di errore / warnings per iterazione corrente 
            string errorMessages_Format2 = String.Empty;
            string warningMessages_Format2 = String.Empty;


            List<Excel_Format2_Row_LegaProperties> listaConcentrations;

            // analisi per il foglio corrente 
            Constants_Excel.EsitoRecuperoInformazioniFoglio esitoLetturaFoglioSecondoFormato = ExcelRecognizers.Recognize_Format2_InfoLegheConcentrazioni(
                ref currentSheet,
                out startingRow,
                out startingCol,
                out listaConcentrations,
                out errorMessages_Format2,
                out warningMessages_Format2);

            // attribuzione dei messaggi globali errors e warnings per la lettura del foglio corrente 
            _possibleErrors_LetturaSecondoFormato = errorMessages_Format2;
            _possibleWarnings_LetturaSecondoFormato = warningMessages_Format2;


            if (esitoLetturaFoglioSecondoFormato == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto || 
                esitoLetturaFoglioSecondoFormato == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoConWarnings)
            {
                _startingPosLeghe_row_format2 = startingRow;
                _startingPosLeghe_col_format2 = startingCol;
                _colonneConcentrazioniSecondoFormato = listaConcentrations;

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


        #region GETTERS : INFORMAZIONI LETTE E VALIDATE PER L'EXCEL CORRENTEMENTE IN ANALISI

        /// <summary>
        /// Ritorno la lista di tutte le informazioni lette dal primo formato excel 
        /// </summary>
        public List<Excel_Format1_Sheet> GetExcelFormat_1Info { get { return _sheetsLetturaFormat_1; } }


        /// <summary>
        /// Ritorno la lista di tutte le informazioni lette dal secondo formato excel
        /// </summary>
        public List<Excel_Format2_Sheet> GetExcelFormat_2Info { get { return _sheetsLetturaFormat_2; } }

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
        /// Lista di tutti i fogli excel che vengono eventualmente letti dalla prima tipologia di formato 
        /// per il foglio excel
        /// </summary>
        private List<Excel_Format1_Sheet> _sheetsLetturaFormat_1;


        /// <summary>
        /// Lista di tutti i fogli excel che vengono eventualmente letti dalla seconda tipologia di formato 
        /// per il foglio excel
        /// </summary>
        private List<Excel_Format2_Sheet> _sheetsLetturaFormat_2;

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
        public List<Excel_Format1_Sheet> SetExcelFormat_1Info { set { _sheetsLetturaFormat_1 = value; } }


        /// <summary>
        /// Carico tutte le informazioni prima di passare alla scrittura 
        /// </summary>
        public List<Excel_Format2_Sheet> SetExcelFormat_2Info { set { _sheetsLetturaFormat_2 = value; } }

        #endregion
    }
}
