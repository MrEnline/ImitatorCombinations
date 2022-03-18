
namespace ImitComb.domain
{
    class ClearListSelectZDVsUseCase
    {
        IOperations repositoryImpl;
        public ClearListSelectZDVsUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public void ClearListSelectZDVs()
        {
            repositoryImpl.ClearListSelectZDVs();
        }
    }
}
