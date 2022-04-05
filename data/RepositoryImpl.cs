using ClosedXML.Excel;
using ImitComb.domain;
using ImitComb.domain.Entity;
using ImitComb.presentation;
using OPCAutomation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;

namespace ImitComb.data
{

    class RepositoryImpl : IOperations
    {
        private XLWorkbook workbook;
        private IXLWorksheet worksheetCombs;
        private IXLWorksheet worksheetTags;
        private IXLWorksheet worksheetSubScribe;
        private IXLWorksheet worksheetSettings;

        private const string NAME_COMBS_LIST_EXCEL = "Combinations";
        private const string NAME_TAGS_LIST_EXCEL = "Tags";
        private const string NAME_SUBSCRIBE_LIST_EXCEL = "SubScribe";
        private const string NAME_SETTINGS_LIST_EXCEL = "Settings";

        private const string KEY_SERVER = "Server";
        private const string STATUS = ".Status";
        private const string GROUP_OPC_WRITE = "GroupOPCWrite";
        private const string GROUP_OPC_READ = "GroupOPCREAD";
        private const string GROUP_OPC_DATA_CHANGE = "GroupOPCDATACHANGE";
        
        private const int COLUMN_SETTINGS = 2;
        private const int ROW_SERVER_SETTINGS = 1;
        private const int ROW_PREFIX_SETTINGS = 2;
        private const int ROW_NAMETU_SETTINGS = 3;

        private const int OFFSET_BETWEEN_COMBINATIONS = 2;
        private const int DELAY = 1000;
        private const string ERROR_MESSAGE = "Ошибка";
        private const string SUCCESS_MESSAGE = "Успешно";
        private const string WARNING_MESSAGE = "Введите данные";
        private const string EXECUTE_AUTO_CHECK = "Выполняется";
        private const string DONE_AUTO_CHECK = "Операция выполнена";
        private const string ABORT_AUTO_CHECK = "Операция прервана";
        private string path = @"./ResultAutoImitations.txt";
        private Dictionary<string, List<String>> dictCombs;
        private Dictionary<string, string> dictTags;
        private Dictionary<string, List<String>> dictErrorCombs;
        private List<string> listZDVs;
        private List<string> listSelectZDVs;
        private List<string> listAutoCheckZDVs;
        private List<string> listDataChangeTags;
        private List<string> listSelectOneZDV;
        private Regex regexCombs;
        //private string pattern = @"[а-яA-Я]+\s+[№:]+\s+\d+";
        private string pattern = @"[а-яA-Я]+\s+[№:]+\s+\d+[^.]";
        private string nameServer;
        private string nameArea;
        private OPCServer server;
        private OPCGroups opcGroups;

        private OPCSettings opcWrite;
        private OPCSettings opcRead;
        private OPCSettings opcDataChange;

        private ArrayClass arrayWrite;
        private ArrayClass arrayRead;
        private ArrayClass arrayDataChange;

        private OPCState opcState;
        object value;
        object quality;
        object timeStamp;

        private Command command;

        private Thread backThread;
        private int backThreadId = 0;

        private IExeState exeState;

        private string nameTU;
        private int numberTU;
        private int numRowActiveTag = 0;
        private int prevIndexComb = -1;

        private Dictionary<int, int> setValue;

        public RepositoryImpl()
        {
            regexCombs = new Regex(pattern);

            dictCombs = new Dictionary<string, List<String>>();
            dictTags = new Dictionary<string, string>();
            dictErrorCombs = new Dictionary<string, List<String>>();

            listSelectZDVs = new List<string>();
            listDataChangeTags = new List<string>();
            listAutoCheckZDVs = new List<string>();
            listSelectOneZDV = new List<string>();

            server = new OPCServer();

            opcState = new OPCState();

            arrayWrite = new ArrayClass();
            arrayRead = new ArrayClass();
            arrayDataChange = new ArrayClass();

            opcRead = new OPCSettings();
            opcWrite = new OPCSettings();
            opcDataChange = new OPCSettings();

            command = new Command();

            setValue = new Dictionary<int, int>();

        }

        public bool CheckExcel(string pathCombFile)
        {
            if (String.IsNullOrEmpty(pathCombFile))
            {
                MessageBox.Show("Не введен путь до файла");
            }
            else
            {
                try
                {
                    workbook = new XLWorkbook(pathCombFile);
                    worksheetCombs = workbook.Worksheet(NAME_COMBS_LIST_EXCEL);
                    worksheetTags = workbook.Worksheet(NAME_TAGS_LIST_EXCEL);
                    worksheetSubScribe = workbook.Worksheet(NAME_SUBSCRIBE_LIST_EXCEL);
                    worksheetSettings = workbook.Worksheet(NAME_SETTINGS_LIST_EXCEL);
                    GetNameTU();
                    return true;
                }
                catch (Exception ex)
                {
                        MessageBox.Show(ex.Message);
                        workbook = null;
                }
            }
            return false;
        }

        private void GetNameTU()
		{
            if (worksheetSettings != null)
                nameTU = worksheetSettings.Cell(ROW_NAMETU_SETTINGS, COLUMN_SETTINGS).Value.ToString();
            if (!string.IsNullOrEmpty(nameTU))
                if (Regex.IsMatch(nameTU, @"\d"))
                    numberTU = Convert.ToInt16(Regex.Match(nameTU, @"\d").Value);
        }

        public string GetNameServer(string nameServer = "")
        {
            this.nameServer = GetInputData(nameServer, worksheetSettings, ROW_SERVER_SETTINGS);
            return this.nameServer;
        }

        public string GetNameArea(string nameArea = "")
        {
            this.nameArea = GetInputData(nameArea, worksheetSettings, ROW_PREFIX_SETTINGS);
            if (String.IsNullOrEmpty(this.nameArea)) return WARNING_MESSAGE;
            return this.nameArea;
        }

        private string GetInputData(string textInput, IXLWorksheet worksheet, int row)
        {
            if (!String.IsNullOrEmpty(textInput))
            {
                return textInput;
            }
            else if (worksheet != null && !String.IsNullOrEmpty(worksheet.Cell(row, COLUMN_SETTINGS).Value.ToString()))
            {
                return worksheet.Cell(row, COLUMN_SETTINGS).Value.ToString();
            }
            return "";
        }

        public Dictionary<string, List<string>> ReadCombinations()
        {
            if (workbook == null) return null;
            if (dictTags.Count == 0)
                ReadTags();
            dictCombs.Clear();
            try
            {
                int count = 0, row = 1, col = 1;
                string keyComb = "";
                while (count < OFFSET_BETWEEN_COMBINATIONS)
                {
                        string valueCell = worksheetCombs.Cell(row, col).Value as String;
                        count = String.IsNullOrEmpty(valueCell) ? count + 1 : 0;
                        CreateDictionaryCombinations(valueCell, ref keyComb);
                        row++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return dictCombs;
        }

        private void ReadTags()
        {
            int row = 1, colDesc = 1, colTags = 2;
            dictTags.Clear();
            while (!String.IsNullOrEmpty(worksheetTags.Cell(row, colDesc).Value.ToString()))
            {
                string key = (worksheetTags.Cell(row, colDesc).Value as String).ToLower();
                string value = worksheetTags.Cell(row, colTags).Value as String;
                dictTags.Add(key, value);
                row++;
            }
            //if (server != null)
            //    MessageBox.Show("Состояние сервера: " + server.ServerState.ToString());
        }

        private void CreateDictionaryCombinations(string valueCell, ref string keyComb)
        {
            if (regexCombs.IsMatch(valueCell))
            {
                keyComb = valueCell;
                dictCombs.Add(keyComb, new List<String>());
            }
            else if (!String.IsNullOrEmpty(valueCell))
            {
                dictCombs[keyComb].Add(valueCell);
            };
        }

        public List<string> CreateListSelectZDVs(string nameZDV)
        {
            if (!listSelectZDVs.Contains(nameZDV))
                listSelectZDVs.Add(nameZDV);
            return listSelectZDVs;
        }

        private List<string> GetListZDVs(int keyOperation)
        {
            switch (keyOperation)
            {
                case 0:
                        return listZDVs;
                case 1:
                case 2:
                case 3:
                case 4:
                case 5:
                        return listSelectZDVs;
                case 6:
                        return listAutoCheckZDVs;
                case 7:
                        return listSelectOneZDV;
            }
            MessageBox.Show("Список с задвижками пустой");
            return null;
        }


        public void ClearListSelectZDVs()
        {
            if (listSelectZDVs.Count > 0)
                listSelectZDVs.Clear();
        }

        public void CreateChoiceCombZDVs(List<string> listZDVs)
        {
            this.listZDVs = listZDVs;
        }

        public void ConnectServer()
        {
            //if (server != null)
            //    MessageBox.Show("Состояние сервера: " + server.ServerState.ToString());
            try
            {
                server.Connect(nameServer);
                opcGroups = server.OPCGroups;
                opcState.ConnectOPC = true;
                opcState.StateServer = server.ServerState;
            }
            catch (Exception ex)
            {
                opcState.ConnectOPC = false;
                MessageBox.Show(ex.Message);
            }
        }

        public string Imitation(int valueCommand, int keyOperations)
        {
            List<string> tags = GetListZDVs(keyOperations);
            InitArrays(arrayWrite, tags.Count);
            if (tags == null) return ERROR_MESSAGE;
            FormationArrays(arrayWrite, GetListZDVs(keyOperations), "dictionary");
            if (!InitOPC(arrayWrite, GROUP_OPC_WRITE, tags.Count, opcWrite)) return ERROR_MESSAGE;
            //MessageBox.Show("Imitation");
            WriteValues(valueCommand, opcWrite);
            opcGroups.Remove(GROUP_OPC_WRITE);
            opcWrite.opcGroup = null;
            return SUCCESS_MESSAGE;
        }

        private bool InitOPC(ArrayClass arrayClass, string nameOPCGroups, int count, OPCSettings opcData)
        {
            if (server != null)
            {
                try
                {
                        opcData.opcGroup = opcGroups.Add(nameOPCGroups);
                        opcData.opcItems = opcData.opcGroup.OPCItems;
                        opcData.opcItems.AddItems(count, ref arrayClass.ItemIDs, ref arrayClass.ClientHandles, out arrayClass.ServerHandles, out arrayClass.Errors);
                }
                catch (Exception)
                {

                        MessageBox.Show("Список с сигналами скорее всего пустой");
                        opcRead.opcGroup = null;
                        opcWrite.opcGroup = null;
                        return false;
                }
            }
            else
            {
                MessageBox.Show("Подключение к OPC-серверу не было произведено");
                return false;
            }
            return true;
        }

        private void FormationArrays(ArrayClass arrayClass, List<string> currListData, string data)
        {
            for (int i = 1; i <= currListData.Count; i++)
            {
                string tag = GetTag(i, currListData, data);
                arrayClass.ItemIDs.SetValue(nameArea + tag, i);
                arrayClass.ClientHandles.SetValue(i, i);
            }
        }

        private string GetTag(int i, List<string> currListTags, string data)
        {
            switch (data)
            {
                case "dictionary":
                        return dictTags[(currListTags[i - 1]).ToLower()] + STATUS;
                case "list":
                        return currListTags[i - 1];
                default:
                        return "";
            }
        }

        private void InitArrays(ArrayClass arrayClass, int count)
        {
            arrayClass.ItemIDs = new string[count + 1];
            arrayClass.ClientHandles = new int[count + 1];
            arrayClass.ServerHandles = new int[count + 1];
            arrayClass.Errors = new int[count + 1];
        }

        private void WriteValues(int valueCommand, OPCSettings opcData)
        {
            //MessageBox.Show(opcData.opcItems.Count.ToString());
            //MessageBox.Show(Thread.CurrentThread.ManagedThreadId.ToString());

            //OPCItem item1 = opcItems.GetOPCItem((int)arrayWrite.ServerHandles.GetValue(1));
            //int count = opcData.opcItems.Count;
            //item1.Write(valueCommand);

            //foreach (OPCItem item in opcItems)
            //{
            //    item.Write(valueCommand);
            //}
            int count = opcData.opcItems.Count;
            OPCItems opcItems = opcData.opcItems;
            for (int i = 1; i <= count; i++)
            {
                OPCItem item = opcItems.GetOPCItem((int)arrayWrite.ServerHandles.GetValue(i));
                item.Write(valueCommand);
            }
        }

        public List<SignalState> ReadValues(OPCSettings opcData, ArrayClass arrayClass)
        {
            List<SignalState> listValue = new List<SignalState>();
            //foreach (OPCItem item in opcData.opcItems)
            //{
            //    item.Read(1, out value, out quality, out timeStamp);
            //    listValue.Add(new SignalState(value, quality, timeStamp));
            //}
            int count = opcData.opcItems.Count;
            OPCItems opcItems = opcData.opcItems;
            for (int i = 1; i <= count; i++)
            {
                OPCItem item = opcItems.GetOPCItem((int)arrayClass.ServerHandles.GetValue(i));
                item.Read(1, out value, out quality, out timeStamp);
                listValue.Add(new SignalState(value, quality, timeStamp));
            }
            return listValue;
        }

        public OPCState SubScribeTags()
        {
            if (!GetTagsSubscribe() && !opcState.ConnectOPC) return opcState;
            ReadTagsValues(listDataChangeTags, arrayDataChange, opcDataChange, GROUP_OPC_DATA_CHANGE, "list");
            opcState.OpcGroup = opcDataChange.opcGroup;
            opcDataChange.opcGroup.DataChange += ObjOPCGroup1_DataChange;
            return opcState;
        }

        private List<SignalState> ReadTagsValues(List<string> listTags, ArrayClass arrayClass, OPCSettings opcData, string nameGroup, string data = "dictionary")
        {
            InitArrays(arrayClass, listTags.Count);

            FormationArrays(arrayClass, listTags, data);

            InitOPC(arrayClass, nameGroup, listTags.Count, opcData);
            List<SignalState> listValues = ReadValues(opcData, arrayClass);
            if (nameGroup != GROUP_OPC_DATA_CHANGE)
            {
                opcGroups.Remove(GROUP_OPC_READ);
                opcData.opcGroup = null;
            }
            return listValues;
        }

        private bool GetTagsSubscribe()
        {
            if (worksheetSubScribe != null)
            {
                int i = 1;
                while (worksheetSubScribe.Cell(i, 1).Value.ToString() != "")
                {
                        listDataChangeTags.Add(worksheetSubScribe.Cell(i, 1).Value.ToString());
                        i++;
                }
            }
            return listDataChangeTags.Count > 0;
        }

        public void AutoCheck(IExeState exeState)
        {
            this.exeState = exeState;
            if (backThreadId == 0)
                InitBackThread();
            else
                AbortBackThread();
        }

        private void InitBackThread()
        {
            backThread = new Thread(ExecuteAutoCheck);
            backThread.Start();
            backThreadId = backThread.ManagedThreadId;
        }

        private void AbortBackThread()
        {
            exeState.GetStateExecute(ABORT_AUTO_CHECK, nameTU, stopAutoImitation: true);
            backThread.Abort();
            backThread.Join();
            backThread = null;
            backThreadId = 0;
            opcRead.opcGroup = null;
            opcWrite.opcGroup = null;
            //opcGroups.Remove(GROUP_OPC_READ);
            //opcGroups.Remove(GROUP_OPC_WRITE);
            //opcGroups.RemoveAll();
            //opcGroups.Add(GROUP_OPC_DATA_CHANGE);
        }

        private void ExecuteAutoCheck()
        {
            try
            {
                if (dictCombs.Count == 0 || dictTags.Count == 0) return;
                foreach (KeyValuePair<string, List<string>> keyValuePair in dictCombs)
                {
                    string keyCurrentComb = keyValuePair.Key;

                    int indexComb = Convert.ToInt32(Regex.Match(Regex.Matches(keyCurrentComb, @"[а-яA-Я]+\s+[№:]+\s+\d+")[1].Value, @"\d+").Value);
                    SetCheckValue(indexComb);

                    listAutoCheckZDVs.Clear();
                    foreach (var value in keyValuePair.Value)
                    {
                        listAutoCheckZDVs.Add(value);
                    }
                    exeState.GetStateExecute(EXECUTE_AUTO_CHECK, nameTU);
                    ExecuteStep(command.SetStatusClose(), 6, setValue[1], keyCurrentComb);
                    foreach (var item in listAutoCheckZDVs)
                    {
                        exeState.GetStateExecute(EXECUTE_AUTO_CHECK, nameTU, keyCurrentComb + "%" + item);
                        listSelectOneZDV.Clear();
                        listSelectOneZDV.Add(item);
                        ExecuteStep(command.SetStatusOpen(), 7, setValue[2], keyCurrentComb, item);
                        ExecuteStep(command.SetStatusClose(), 7, setValue[1], keyCurrentComb, item);
                    }
                    ExecuteStep(command.SetStatusOpen(), 6, setValue[2], keyCurrentComb);
                }
                OutputErrorData();
                exeState.GetStateExecute(DONE_AUTO_CHECK, nameTU);
                MessageBox.Show(DONE_AUTO_CHECK);
            }
            catch (ThreadAbortException)
            {}
        }

        //установить значения с которыми будут проверяться комбинации
        private void SetCheckValue(int indexComb)
		{
            if (prevIndexComb == indexComb && setValue.Count > 0)
                return;
            setValue.Clear();
            switch (indexComb)
			{
                case 4:
                    setValue.Add(1, 0);
                    setValue.Add(2, 1);
                    numRowActiveTag = 1;
                    break;
                case 2010:
                    setValue.Add(1, 1);
                    setValue.Add(2, 2);
                    numRowActiveTag = 0;
                    break;
            }
            prevIndexComb = indexComb;
		}

        private void ExecuteStep(int commandValue, int keyOperation, int currValueTag, string keyCurrentComb, string nameZDV = "")
        {
            Imitation(commandValue, keyOperation);                     //выполним операции с задвижками из комбинации согласно команде(закрыть, открыть)
            Thread.Sleep(DELAY);
            int valueTag = ReadTagsValues(listDataChangeTags, arrayRead, opcRead, GROUP_OPC_READ, "list")[numRowActiveTag].Value != null ?
                                            (Int32)ReadTagsValues(listDataChangeTags, arrayRead, opcRead, GROUP_OPC_READ, "list")[numRowActiveTag].Value : 0;
            SetDictErrorCombs(currValueTag, keyCurrentComb, valueTag, nameZDV);
        }

        private void OutputErrorData()
        {
            using (FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (StreamWriter writer = new StreamWriter(fileStream))
                {
                        if (dictErrorCombs.Count != 0)
                        {
                            foreach (KeyValuePair<string, List<string>> keyValue in dictErrorCombs)
                            {
                                    if (keyValue.Value.Count > 0)
                                        writer.WriteLine(keyValue.Key + " :" + GetListErrorZDV(keyValue.Value));
                                    else
                                        writer.WriteLine(keyValue.Key);
                            }
                            writer.WriteLine(DONE_AUTO_CHECK);
                        }
                }
            }
        }

        private string GetListErrorZDV(List<string> errorZDV)
        {
            StringBuilder result = new StringBuilder();
            foreach (var item in errorZDV)
            {
                result.Append(item + " ");
            }
            return result.ToString().Trim();
        }

        private void SetDictErrorCombs(int currBlockWay, string keyCurrentComb, int valueTag, string nameZDV)
        {
            //if (currBlockWay != (Int32)arrayDataChange.ItemValues.GetValue(valueTag))
            //    if (!dictErrorCombs.ContainsKey(keyCurrentComb))
            //        dictErrorCombs.Add(keyCurrentComb, new List<string>());
            //    else
            //        dictErrorCombs[keyCurrentComb].Add(nameZDV);
            if (currBlockWay != valueTag)
            {
                if (!dictErrorCombs.ContainsKey(keyCurrentComb))
                {
                        dictErrorCombs.Add(keyCurrentComb, new List<string>());
                        if (nameZDV != "")
                            dictErrorCombs[keyCurrentComb].Add(nameZDV);
                }
                else
                {
                        if (!dictErrorCombs[keyCurrentComb].Contains(nameZDV))
                            dictErrorCombs[keyCurrentComb].Add(nameZDV);
                }
            }
        }

        //обработчик события
        public void ObjOPCGroup1_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {

            //if (ItemValues.GetValue(1) != null && 2 == (Int32)ItemValues.GetValue(1))
            //{

            //}
            //else
            //{

            //}
            arrayDataChange.ItemValues = ItemValues;
            arrayDataChange.Qualities = Qualities;
            arrayDataChange.TimeStamps = TimeStamps;

            //TEST_VALUE = (Int32)ItemValues.GetValue(1);
            //if ((currBlockWay) != (Int32)ItemValues.GetValue(1) && !String.IsNullOrEmpty(keyCurrentComb))
            //{
            //    if (!dictErrorCombs.ContainsKey(keyCurrentComb))
            //        dictErrorCombs.Add(keyCurrentComb, new List<string>());
            //    else
            //        dictErrorCombs[keyCurrentComb].Add(listSelectOneZDV[0]);
            //}
        }
    }
}
