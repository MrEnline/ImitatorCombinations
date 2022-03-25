using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImitComb.presentation
{
    interface IExeState
    {
        void GetStateExecute(string state, string combination = "", bool stopAutoImitation = false);
    }
}
