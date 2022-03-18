
namespace ImitComb.domain
{
    class GetNameServerUseCase
    {
        IOperations repositoryImpl;
        public GetNameServerUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public string GetNameServer(string nameServer = "")
        {
            return repositoryImpl.GetNameServer(nameServer);
        }
    }
}
