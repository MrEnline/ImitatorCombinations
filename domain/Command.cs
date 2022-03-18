
namespace ImitComb.domain
{
    class Command
    {
        private const int OPENING = 4;
        private const int OPEN = 1;
        private const int CLOSING = 2;
        private const int CLOSE = 3;
        private const int MIDDLE = 5;

        public int SetStatusOpen()
        {
            return OPEN;
        }

        public int SetStatusOpening()
        {
            return OPENING;
        }

        public int SetStatusClosing()
        {
            return CLOSING;
        }

        public int SetStatusClose()
        {
            return CLOSE;
        }

        public int SetStatusMiddle()
        {
            return MIDDLE;
        }
    }
}
