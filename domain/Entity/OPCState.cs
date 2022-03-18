using OPCAutomation;

namespace ImitComb.domain.Entity
{
    class OPCState
    {
        public bool ConnectOPC { get; set; }
        public OPCGroup OpcGroup { get; set; }
        public int StateServer { get; set; }
    }
}
