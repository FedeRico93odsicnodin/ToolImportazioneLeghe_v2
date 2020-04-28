using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolImportazioneLeghe_Console.Excel.Model_Excel;
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
                    // ricreazione del file 
                    ServiceLocator.GetUtilityFunctions.BuildFilePath(excelPath);
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
            }
            catch (Exception e)
            {
                throw new Exception(String.Format(ExceptionMessages.EXCEL_PROBLEMAAPERTURAFILE, ServiceLocator.GetUtilityFunctions.GetFileName(excelPath), e.Message));
            }
            
            return false;
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

                }
            }
            else if(_formatoExcel == Constants.FormatFileExcel.Cliente)
            {
                _sheetsLetturaFormat_2 = new List<Excel_Format2_Sheet>();

                foreach (ExcelWorksheet currentWorksheet in _openedExcel.Workbook.Worksheets)
                {

                }
            }

            return false;
        }


        /// <summary>
        /// Lettura di tutte le informazioni lette con lo step precedente
        /// viene ritornato false se:
        /// 1) non è stata letta nessuna informazione dai fogli precedentemente letti come validi
        /// </summary>
        /// <returns></returns>
        public bool ReadExcelInfo()
        {
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
        /// <returns></returns>
        private Constants_Excel.TipologiaFoglio_Format1 RecognizeTipoFoglio_Format1(ref ExcelWorksheet currentSheet)
        {
            return Constants_Excel.TipologiaFoglio_Format1.NotDefined;
        }


        /// <summary>
        /// Riconoscimento se il formato usato per la seconda tipologia per il foglio corrente è effettivamente valida per 
        /// il riconoscimento del foglio corrente come formato 2
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <returns></returns>
        private bool RecognizeTipoFoglio_Format2(ref ExcelWorksheet currentSheet)
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
