
namespace ImitComb.domain
{
    class GetNameAreaUseCase
    {
        IOperations repositoryImpl;
        public GetNameAreaUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public string GetNameArea(string nameArea = "")
        {
            return repositoryImpl.GetNameArea(nameArea);
        }
    }
}
