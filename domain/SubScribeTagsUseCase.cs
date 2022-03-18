using OPCAutomation;
using ImitComb.domain.Entity;

namespace ImitComb.domain
{
    class SubScribeTagsUseCase
    {
        IOperations repositoryImpl;
        public SubScribeTagsUseCase(IOperations repositoryImpl)
        {
            this.repositoryImpl = repositoryImpl;
        }

        public OPCState SubScribeTags()
        {
            return repositoryImpl.SubScribeTags();
        }
    }
}
