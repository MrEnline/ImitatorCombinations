
namespace ImitComb.domain
{
    class CheckExcelUseCase
    {
        IOperations repositoryImpl;
        public CheckExcelUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public bool CheckExcel(string pathCombFile)
        {
            return repositoryImpl.CheckExcel(pathCombFile);
        }
    }
}
