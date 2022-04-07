using OPCAutomation;
using System.Collections.Generic;

namespace ImitComb.domain.Entity
{
    class OPCState
    {
        public bool ConnectOPC { get; set; }
        public OPCGroup OpcGroup { get; set; }
        public int StateServer { get; set; }
        public List<TagForViewModel> tags { set; get; }
    }
}
