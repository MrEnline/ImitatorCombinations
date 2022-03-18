using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImitComb.data
{
    class SignalState
    {
        public object Value;
        public object Quality;
        public object TimeStamp;

        public SignalState(object Value, object Quality, object TimeStamp)
        {
            this.Value = Value;
            this.Quality = Quality;
            this.TimeStamp = TimeStamp;
        }
    }
}
