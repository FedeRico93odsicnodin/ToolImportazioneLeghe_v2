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
        #region DIZIONARIO DI PROPRIETA PER L'ISTANZA DEL WRAPPER

        /// <summary>
        /// Contiene le proprieta in base all'istanza che si decide in fase di costruzione 
        /// del contenitore di proprieta (proprieta obbligatorie)
        /// </summary>
        private Dictionary<string, string> _propertiesSet_Mandatory;


        /// <summary>
        /// Contiene il set delle eventuali proprieta opzionali presenti per l'istanza corrente 
        /// </summary>
        private Dictionary<string, string> _propertiesSet_Optional;
        
        #endregion


        #region COSTRUTTORE

        /// <summary>
        /// Inizializzazione del dizionario delle proprieta in base alla tipologia di foglio corrente 
        /// </summary>
        /// <param name="mandatoryProperties"></param>
        /// <param name="optionalProperties"></param>
        /// <param name="sheetTipology"></param>
        public Excel_PropertyWrapper(string[] mandatoryProperties, string[] optionalProperties, Constants_Excel.TipologiaPropertiesFoglio sheetTipology)
        {

            if (mandatoryProperties == null)
                throw new Exception(ExceptionMessages.EXCEL_READINGPROPERTIES);

            if (mandatoryProperties.Count() == 0)
                throw new Exception(ExceptionMessages.EXCEL_READINGPROPERTIES);

            _propertiesSet_Mandatory = new Dictionary<string, string>();

            foreach (string mandatoryProperty in mandatoryProperties)
                _propertiesSet_Mandatory.Add(mandatoryProperty, String.Empty);

            if (optionalProperties == null)
                throw new Exception(ExceptionMessages.EXCEL_READINGPROPERTIES);

            if (optionalProperties.Count() == 0)
                throw new Exception(ExceptionMessages.EXCEL_READINGPROPERTIES);

            _propertiesSet_Optional = new Dictionary<string, string>();

            foreach (string optionalProperty in optionalProperties)
                _propertiesSet_Optional.Add(optionalProperty, String.Empty);


            TipologyPropertiesWrapper = sheetTipology;
        }

        #endregion


        #region GETTERS - SETTERS DI PROPRIETA E TIPOLOGIA WRAPPER

        /// <summary>
        /// Indica la tipologia di proprieta per il wrapper corrente tra quelle per cui il wrapper 
        /// puo essere effettivamente istanziato
        /// </summary>
        public Constants_Excel.TipologiaPropertiesFoglio TipologyPropertiesWrapper { get; }


        /// <summary>
        /// Permette di ritornare il counter per le proprieta obbligatorie correnti, questo valoreè eventualmente da confrontare con 
        /// la dimensionalita delle proprieta obbligatorie per capire se tutti gli inserimenti sono stati effettuati
        /// </summary>
        public int CounterMandatoryProperties
        {
            get
            {
                // ritorno il counter sulle proprieta che sono state effettivamente valorizzate durante un particolare inserimento per le proprieta obbligatorie
                return _propertiesSet_Mandatory.Where(x => x.Value != String.Empty).Select(x => x.Key).ToList().Count();
            }
        }


        /// <summary>
        /// Permette di ottenere il valore per una proprieta obbligatoria inserita 
        /// all'interno del wrapper corrente 
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public string GetMandatoryProperty(string property)
        {
            if (!_propertiesSet_Mandatory.ContainsKey(property))
                throw new Exception(String.Format(ExceptionMessages.EXCEL_PROPERTYNOTDEFINED, TipologyPropertiesWrapper));

            return _propertiesSet_Mandatory[property];
        }


        /// <summary>
        /// Permette di ottenere il valore per la proprieta opzionale inserita 
        /// all'interno del wrapper corrente 
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public string GetOptionalProperty(string property)
        {
            if(!_propertiesSet_Optional.ContainsKey(property))
                throw new Exception(String.Format(ExceptionMessages.EXCEL_PROPERTYNOTDEFINED, TipologyPropertiesWrapper));

            return _propertiesSet_Optional[property];
        }


        /// <summary>
        /// Permette di inserire un nuovo valore per la proprieta qualora questo sia nella definizione
        /// delle proprieta obbligatorie
        /// </summary>
        /// <param name="property"></param>
        /// <param name="value"></param>
        public void InsertMandatoryValue(string property, string value)
        {
            if(!_propertiesSet_Mandatory.ContainsKey(property))
                throw new Exception(String.Format(ExceptionMessages.EXCEL_PROPERTYNOTDEFINED, TipologyPropertiesWrapper));

            _propertiesSet_Mandatory[property] = value;
        }


        /// <summary>
        /// Permette di inserire un nuovo valore per la proprieta qualora questa sia nella definizione 
        /// delle proprieta opzionali
        /// </summary>
        /// <param name="property"></param>
        /// <param name="value"></param>
        public void InsertOptionalValue(string property, string value)
        {
            if(!_propertiesSet_Optional.ContainsKey(property))
                throw new Exception(String.Format(ExceptionMessages.EXCEL_PROPERTYNOTDEFINED, TipologyPropertiesWrapper));

            _propertiesSet_Optional[property] = value;
        }

        #endregion

    }
}
