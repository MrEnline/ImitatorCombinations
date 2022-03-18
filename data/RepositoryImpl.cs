using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using OPCAutomation;
using ImitComb.domain;
using ImitComb.domain.Entity;
using System.Threading;
using System.IO;

namespace ImitComb.data
{

    class RepositoryImpl : IOperations
    {
        private XLWorkbook workbook;
        private IXLWorksheet worksheetCombs;
        private IXLWorksheet worksheetTags;
        private IXLWorksheet worksheetSubScribe;
        private const string NAME_COMBS_LIST_EXCEL = "Combinations";
        private const string NAME_TAGS_LIST_EXCEL = "Tags";
        private const string NAME_SUBSCRIBE_LIST_EXCEL = "SubScribe";
        private const string KEY_SERVER = "Server";
        private const string STATUS = ".Status";
        private const string GROUP_OPC_WRITE = "GroupOPCWrite";
        private const string GROUP_OPC_READ = "GroupOPCREAD";
        private const string GROUP_OPC_DATA_CHANGE = "GroupOPCDATACHANGE";
        private const int COLUMN_SETTINGS = 4;
        private const int ROW_SERVER_SETTINGS = 1;
        private const int ROW_PREFIX_SETTINGS = 2;
        private const int OFFSET_BETWEEN_COMBINATIONS = 2;
        private const int DELAY = 1000;
        private const string ERROR_MESSAGE = "Ошибка";
        private const string SUCCESS_MESSAGE = "Успешно";
        private const string WARNING_MESSAGE = "Введите данные";
        private Dictionary<string, List<String>> dictCombs;
        private Dictionary<string, string> dictTags;
        private Dictionary<string, List<String>> dictErrorCombs;
        private List<string> listZDVs;
        private List<string> listSelectZDVs;
        private List<string> listAutoCheckZDVs;
        private List<string> listDataChangeTags;
        private List<string> listSelectOneZDV;
        private Regex regexCombs;
        private string pattern = @"[а-яA-Я]+\s+[№:]+\s+\d+";
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

        private SignalState signalState;

        private OPCState opcState;
        object value;
        object quality;
        object timeStamp;

        private Command command;
        private int currBlockWay;
        private string keyCurrentComb;

        private Thread threadTwo;

        private int TEST_VALUE = 0;

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

        public string GetNameServer(string nameServer = "")
        {
            this.nameServer = GetInputData(nameServer, worksheetCombs, ROW_SERVER_SETTINGS);
            return this.nameServer;
        }

        public string GetNameArea(string nameArea = "")
        {
            this.nameArea = GetInputData(nameArea, worksheetCombs, ROW_PREFIX_SETTINGS);
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
                return worksheetCombs.Cell(row, COLUMN_SETTINGS).Value.ToString();
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

        private string GetTag(int i, List<string> currListTags,  string data)
        {
            switch (data)
            {
                case "dictionary":
                    return dictTags[(currListTags[i - 1]).ToLower()] + STATUS;
                case "list":
                    return currListTags[i - 1];
            }
            return "";
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
            foreach (OPCItem item in opcData.opcItems)
            {
                item.Write(valueCommand);
            }
        }

        public List<SignalState> ReadValues(OPCSettings opcData)
        {
            List<SignalState> listValue = new List<SignalState>();
            int i = 0;
            foreach (OPCItem item in opcData.opcItems)
            {
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
            List<SignalState> listValues = ReadValues(opcData);
            if (nameGroup != GROUP_OPC_DATA_CHANGE)
                opcGroups.Remove(GROUP_OPC_READ);
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

        public void AutoCheck()
        {
            //threadTwo = new Thread(ExecuteAutoCheck);
            //threadTwo.Start();
            if (dictCombs.Count == 0 || dictTags.Count == 0) return;
            foreach (KeyValuePair<string, List<string>> keyValuePair in dictCombs)
            {
                int valueTag = 0;
                keyCurrentComb = keyValuePair.Key;
                //MessageBox.Show(keyCurrentComb);
                listAutoCheckZDVs.Clear();
                foreach (var value in keyValuePair.Value)
                {
                    listAutoCheckZDVs.Add(value);
                }
                ExecuteStep(command.SetStatusClose(), 6, 1, keyCurrentComb);
                foreach (var item in listAutoCheckZDVs)
                {
                    listSelectOneZDV.Clear();
                    listSelectOneZDV.Add(item);
                    ExecuteStep(command.SetStatusOpen(), 7, 2, keyCurrentComb, item);
                    ExecuteStep(command.SetStatusClose(), 7, 1, keyCurrentComb, item);
                }
            }
            OutputErrorData();
            MessageBox.Show("Операция выполнена");
        }

        private void OutputErrorData()
        {
            string path = @"./ResultAutoImitations.txt";
            using (FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (StreamWriter writer = new StreamWriter(fileStream))
                {
                    if (dictErrorCombs.Count != 0)
                    {
                        foreach (KeyValuePair<string, List<string>> keyValue in dictErrorCombs)
                        {
                            if (keyValue.Value.Count > 0)
                                writer.WriteLine(keyValue.Key + " :" + keyValue.Value);
                            else
                                writer.WriteLine(keyValue.Key);
                        }
                        writer.WriteLine("Операция выполнена");
                    }
                }
            }
        }

        private void ExecuteStep(int commandValue, int keyOperation, int currBlockWay, string keyCurrentComb, string nameZDV = "")
        {
            Imitation(commandValue, keyOperation);                     //закроем все задвижки из комбинации
            Thread.Sleep(DELAY);
            int valueTag = (Int32)ReadTagsValues(listDataChangeTags, arrayRead, opcRead, GROUP_OPC_READ, "list")[0].Value;
            SetDictErrorCombs(currBlockWay, keyCurrentComb, valueTag, nameZDV);
        }

        private void ExecuteAutoCheck()
        {

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
            TEST_VALUE = (Int32)ItemValues.GetValue(1);
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
