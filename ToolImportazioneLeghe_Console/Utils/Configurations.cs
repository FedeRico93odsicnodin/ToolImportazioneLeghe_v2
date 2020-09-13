using Common.ExtendedElements;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolImportazioneLeghe_Console.Utils
{
    /// <summary>
    /// Classe nella quale si eseguira la lettura delle configurazioni aggiornando il file delle costanti e nella 
    /// quale verrà contenuta l'informazione relativa al tempo che sta scorrendo
    /// eventualmente questa classe dovra essere da aggiornare con le procedura per la scrittura e recupero dei valori 
    /// delle informazioni all'interno di un file ausiliario per una eventuale ripresa dell'import
    /// </summary>
    public class Configurations
    {
        /// <summary>
        /// Informazione relativa al tempo intercorrente sulla procedura 
        /// </summary>
        private ModifiedStopWatch _currentTimeOnProcedure;


        /// <summary>
        /// Inizializzazione parametri costruttre
        /// </summary>
        public Configurations()
        {
            _currentTimeOnProcedure = new ModifiedStopWatch();
        }


        #region TEMPO PROCEDURA 

        /// <summary>
        /// Permette di avviare il tempo su tutta la procedura corrente 
        /// </summary>
        public void StartTimeOnProcedure()
        {
            _currentTimeOnProcedure.Start();
        }


        /// <summary>
        /// Permette di fermare il tempo su tutta la procedura corrente 
        /// </summary>
        public void StopTimeOnProcedure()
        {
            _currentTimeOnProcedure.Stop();
        }


        /// <summary>
        /// Ottiene il tempo intercorso dall'inizio della procedura corrente 
        /// </summary>
        /// <returns></returns>
        public TimeSpan GetCurrentTimeOnProcedure()
        {
            return _currentTimeOnProcedure.Elapsed;
        }

        #endregion
    }
}
