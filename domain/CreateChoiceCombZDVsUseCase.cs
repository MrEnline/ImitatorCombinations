using System.Collections.Generic;

namespace ImitComb.domain
{
    class CreateChoiceCombZDVsUseCase
    {
        IOperations repositoryImpl;
        public CreateChoiceCombZDVsUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public void CreateChoiceCombZDVs(List<string> listZDVs)
        {
            repositoryImpl.CreateChoiceCombZDVs(listZDVs);
        }
    }
}
