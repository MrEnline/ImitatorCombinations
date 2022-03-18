using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImitComb.domain
{
    class AutoCheckUseCase
    {
        IOperations repositoryImpl;
        public AutoCheckUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public void AutoCheck()
        {
            repositoryImpl.AutoCheck();
        }
    }
}
