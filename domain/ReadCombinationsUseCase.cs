using ImitComb.presentation;

namespace ImitComb.domain
{
    class ReadCombinationsUseCase
    {
        IOperations repositoryImpl;
        public ReadCombinationsUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public void ReadCombinations(IExeState exeState)
        {
            repositoryImpl.ReadCombinations(exeState);
        }
    }
}
