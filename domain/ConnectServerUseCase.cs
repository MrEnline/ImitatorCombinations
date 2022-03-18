using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImitComb.domain
{
    class ConnectServerUseCase
    {
        IOperations repositoryImpl;
        public ConnectServerUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public void ConnectServer()
        {
            repositoryImpl.ConnectServer();
        }
    }
}
