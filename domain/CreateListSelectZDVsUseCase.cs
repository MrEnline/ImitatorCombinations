using System;
using System.Collections.Generic;

namespace ImitComb.domain
{
    class CreateListSelectZDVsUseCase
    {
        IOperations repositoryImpl;
        public CreateListSelectZDVsUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public List<String> CreateListSelectZDVs(string nameZDV)
        {
            return repositoryImpl.CreateListSelectZDVs(nameZDV);
        }
    }
}
