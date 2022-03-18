using System.Collections.Generic;

namespace ImitComb.domain
{
    class ReadCombinationsUseCase
    {
        IOperations repositoryImpl;
        public ReadCombinationsUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public Dictionary<string, List<string>> ReadCombinations()
        {
            return repositoryImpl.ReadCombinations();
        }
    }
}
