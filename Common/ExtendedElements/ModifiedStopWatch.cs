using System;
using System.Collections.Generic;
using System.Text;

namespace Common.ExtendedElements
{
    /// <summary>
    /// Stopwatch modificato per poter eventualmente riprendere una esecuzione del tool e tenere traccia 
    /// di quanto tempo trascorre durante l'esecuzione
    /// </summary>
    public class ModifiedStopWatch : System.Diagnostics.Stopwatch
    {
        /// <summary>
        /// Tempo eventualmente da settare in base al recupero per la procedura corrente 
        /// </summary>
        TimeSpan _offset = new TimeSpan();


        /// <summary>
        /// Costruttore in situazione normale nel quale non c'è nessuna ripresa dell'esecuzione 
        /// </summary>
        public ModifiedStopWatch() { }


        /// <summary>
        /// Costruttore per il set del tempo di partenza
        /// </summary>
        /// <param name="offset"></param>
        public ModifiedStopWatch(TimeSpan offset)
        {
            _offset = offset;
        }


        /// <summary>
        /// Permette di settare il tempo di partenza per l'applicazione
        /// </summary>
        /// <param name="offsetElapsedTimeSpan"></param>
        public void SetOffset(TimeSpan offsetElapsedTimeSpan)
        {
            _offset = offsetElapsedTimeSpan;
        }


        /// <summary>
        /// Tempo trascorso dall'inizio della procedura 
        /// </summary>
        public TimeSpan Elapsed
        {
            get { return base.Elapsed + _offset; }
            set { _offset = value; }
        }


        /// <summary>
        /// Tempo trascorso in millisecondi
        /// </summary>
        public long ElapsedMilliseconds
        {
            get { return base.ElapsedMilliseconds + _offset.Milliseconds; }
        }


        /// <summary>
        /// Elapsed Ticks
        /// </summary>
        public long ElapsedTicks
        {
            get { return base.ElapsedTicks + _offset.Ticks; }
        }
    }
}
