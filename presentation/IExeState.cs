using ImitComb.domain.Entity;
using System.Collections.Generic;

namespace ImitComb.presentation
{
    interface IExeState
    {
        void GetStateExecute(StatusOperation status);
        void CreateListBoxItems(Dictionary<string, List<string>> dictCombs);
    }
}
