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
        private List<Excel_Format2_ConcColumns> _colonneConcentrazioniSecondoFormato;


        /// <summary>
        /// Inzio riga di lettura leghe per il secondo formato excel  
        /// </summary>
        private int _startingPosLeghe_row_format2 = 0;


        /// <summary>
        /// Inizio colonna lettura leghe per il secondo formato excel
        /// </summary>
        private int _startingPosLeghe_col_format2 = 0;

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


            // inizializzazione di una delle 2 liste in base al formato
            if (_formatoExcel == Constants.FormatFileExcel.DatabaseLeghe)
            {
                _sheetsLetturaFormat_1 = new List<Excel_Format1_Sheet>();

                foreach(ExcelWorksheet currentWorksheet in _openedExcel.Workbook.Worksheets)
                {
                    currentSheetPosition++;

                    // riconoscimento della tipologia foglio per il primo formato
                    Constants_Excel.TipologiaFoglio_Format1 tipologiaRiconoscita = RecognizeTipoFoglio_Format1(currentWorksheet);
                    if (!(tipologiaRiconoscita == Constants_Excel.TipologiaFoglio_Format1.NotDefined))
                    {
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format1(currentWorksheet.Name, currentSheetPosition, tipologiaRiconoscita.ToString());

                        Excel_Format1_Sheet foglioExcelCorrenteInfo = new Excel_Format1_Sheet(currentWorksheet.Name, tipologiaRiconoscita, currentSheetPosition);

                        // vedo se inserire per posizione iniziale per la lettura delle leghe oppure per i quadranti riconosciuti
                        if(tipologiaRiconoscita == Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe)
                        {
                            foglioExcelCorrenteInfo.StartingRow_letturaLeghe = _startingPosLeghe_row_format1;
                            foglioExcelCorrenteInfo.StartingCol_letturaLeghe = _startingPosLeghe_col_format1;
                            foglioExcelCorrenteInfo.EndingCol_letturaLeghe = _endingPosLeghe_col_format1;
                        }
                        else if(tipologiaRiconoscita == Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni)
                        {
                            foglioExcelCorrenteInfo.GetConcQuadrants_Type2 = _quadrantiConcentrazioneLetturaCorrente;
                        }

                        _sheetsLetturaFormat_1.Add(foglioExcelCorrenteInfo);
                    }
                    else
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_FoglioNonRiconosciuto(currentWorksheet.Name, currentSheetPosition);
                }

                // non è presente nessun foglio sul quale eseguire la lettura delle informazioni
                if (_sheetsLetturaFormat_1.Count() == 0)
                    return false;

                // TODO: capire se discriminare anche a questo livello le informazioni 
                // (ad esempio aggiungendo una variante per la quale ci deve essere almeno un match per foglio concentrazioni / materiali)
                return true;

            }
            else if(_formatoExcel == Constants.FormatFileExcel.Cliente)
            {
                _sheetsLetturaFormat_2 = new List<Excel_Format2_Sheet>();

                foreach (ExcelWorksheet currentWorksheet in _openedExcel.Workbook.Worksheets)
                {
                    currentSheetPosition++;

                    bool hoRiconosciutoSecondaTipologia = RecognizeTipoFoglio_Format2(currentWorksheet);
                    if(hoRiconosciutoSecondaTipologia)
                    {
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_RiconoscimentoSeguenteTipologia_Format2(currentWorksheet.Name, currentSheetPosition);

                        Excel_Format2_Sheet foglioExcelCorrenteInfo = new Excel_Format2_Sheet(currentWorksheet.Name, currentSheetPosition);

                        foglioExcelCorrenteInfo.StartingRow_Leghe = _startingPosLeghe_row_format2;
                        foglioExcelCorrenteInfo.StartingCol_Leghe = _startingPosLeghe_col_format2;
                        //foglioExcelCorrenteInfo.ColonneConcentrazioni = _colonneConcentrazioniSecondoFormato;

                        _sheetsLetturaFormat_2.Add(foglioExcelCorrenteInfo);
                    }
                    else
                        ConsoleService.ConsoleExcel.ExcelReaders_Message_FoglioNonRiconosciuto(currentWorksheet.Name, currentSheetPosition);
                }
            }

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
                if(_sheetsLetturaFormat_1.Count() == 0)
                    throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT1);

                foreach(Excel_Format1_Sheet currentFoglioExcel in _sheetsLetturaFormat_1)
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

                    // riconoscimenti per la lettura del foglio delle leghe 
                    if (currentFoglioExcel.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe)
                        if(!(ExcelReaderInfo.ReadLegheInfo(
                            excelSheetReference, 
                            currentFoglioExcel, 
                            out filledInfo, 
                            out warningMessages, 
                            out errorMessages) == Constants_Excel.EsitoRecuperoInformazioniFoglio.RecuperoCorretto))
                        {
                            // TODO: segnalazione + scrittura log errori warnings
                        }
                    // riconoscimento per la lettura del foglio delle concentrazioni
                    else if(currentFoglioExcel.GetTipologiaFoglio == Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni)
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
                                // TODO: segnalazione + scrittura log errori warnings
                            }
                        }
                        

                }
            }

            #endregion


            #region RECUPERO INFORMAZIONI DAL SECONDO FORMATO

            // caso in cui il file è di secondo formato
            else if(_formatoExcel == Constants.FormatFileExcel.Cliente)
            {
                // eccezione nel caso in cui la lista relativa ai fogli riconosciuti per il primo formato sia NULL EMPTY
                if (_sheetsLetturaFormat_2 == null)
                    throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT2);
                if (_sheetsLetturaFormat_2.Count() == 0)
                    throw new Exception(ExceptionMessages.EXCEL_READERINFO_FOGLINULLEMPTY_FORMAT2);

                foreach(Excel_Format2_Sheet currentFoglioExcel in _sheetsLetturaFormat_2)
                {

                    if (currentFoglioExcel.GetPosSheet == 0)
                        throw new Exception(ExceptionMessages.EXCEL_READERINFO_NESSUNAPOSIZIONETROVATAPERFOGLIOCORRENTE);


                    // recupero del foglio corrente contenuto nel file di riferimento e dal quale continuare la lettura delle informazioni
                    ExcelWorksheet excelSheetReference = _openedExcel.Workbook.Worksheets[currentFoglioExcel.GetPosSheet];

                    // oggetto nel quale inserisco le informazioni recuperate per il foglio corrente
                    Excel_Format2_Sheet filledInfo;

                    //if(ExcelReaderInfo.ReadInfoFormat2(excelSheetReference, currentFoglioExcel, out filledInfo))
                    //{

                    //}



                }
            }

            #endregion

            return false;


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


            bool riconoscimrentoCorrente = ExcelRecognizers.Recognize_Format1_InfoLeghe(ref currentSheet, out startingRow, out startingCol, out endingColIndex);

            // sono riuscito ad inviduare una prima congruenza per il riconoscimento del foglio delle leghe
            if (riconoscimrentoCorrente)
            {
                // attribuzione parametri privati di riga / colonna di lettura per le leghe
                _startingPosLeghe_row_format1 = startingRow;
                _startingPosLeghe_col_format1 = startingCol;
                _endingPosLeghe_col_format1 = endingColIndex;
                return Constants_Excel.TipologiaFoglio_Format1.FoglioLeghe;
            }

            List<Excel_Format1_Sheet2_ConcQuadrant> concentrationsQuadrants;

            // tentativo di riconoscimento foglio concentrazioni
            bool riconoscimentoFoglioConcentrazioni = ExcelRecognizers.Recognize_Format1_InfoConcentrations(ref currentSheet, out concentrationsQuadrants);

            // sono riuscito a riconoscere il foglio per le concentrazioni correnti
            if(riconoscimentoFoglioConcentrazioni)
            {
                _quadrantiConcentrazioneLetturaCorrente = concentrationsQuadrants;
                return Constants_Excel.TipologiaFoglio_Format1.FoglioConcentrazioni;
            }
            
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

            List<Excel_Format2_ConcColumns> listaConcentrations;


            if (ExcelRecognizers.Recognize_Format2_InfoLegheConcentrazioni(ref currentSheet, out startingRow, out startingCol, out listaConcentrations))
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
