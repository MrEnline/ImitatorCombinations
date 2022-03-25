using ImitComb.presentation;


namespace ImitComb.domain
{
    class AutoCheckUseCase
    {
        IOperations repositoryImpl;
        public AutoCheckUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public void AutoCheck(IExeState exeState)
        {
            repositoryImpl.AutoCheck(exeState);
        }
    }
}
