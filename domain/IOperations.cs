using System.Collections.Generic;
using ImitComb.domain.Entity;
using ImitComb.presentation;

namespace ImitComb.domain
{
    interface IOperations
    {
        void ReadCombinations(IExeState exeState);
        bool CheckExcel(string pathCombFile);
        string GetNameServer(string nameServer);
        string GetNameArea(string nameArea);
        void ConnectServer();
        string Imitation(int valueCommand, int keyOperations);
        void CreateChoiceCombZDVs(List<string> listZDVs);
        List<string> CreateListSelectZDVs(string nameZDV);
        void ClearListSelectZDVs();
        OPCState SubScribeTags();
        void AutoCheck(IExeState exeState);
    }
}
