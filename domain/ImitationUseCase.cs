
namespace ImitComb.domain
{
    class ImitationUseCase
    {
        IOperations repositoryImpl;
        public ImitationUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public string Imitation(int valueCommand, int keyOperations)
        {
            return repositoryImpl.Imitation(valueCommand, keyOperations);
        }
    }
}
