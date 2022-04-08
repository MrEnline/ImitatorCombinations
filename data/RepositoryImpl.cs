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
using System.Diagnostics;
using System.Threading.Tasks;

namespace ImitComb.data
{
    //TODO Сделать начало работы с определенной комбинации, а следовательно и поиск данной комбинации

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
        private string nameServer;
        private string nameArea;

        private Dictionary<string, List<String>> dictCombs;
        private Dictionary<string, string> dictTags;
        private Dictionary<string, List<String>> dictErrorCombs;
        private List<string> listZDVs;
        private List<string> listSelectZDVs;
        private List<string> listAutoCheckZDVs;
        private List<string> listDataChangeTags;
        private Dictionary<int, List<string>> dictDataChangeTags;
        private Dictionary<string, int> _dictDataChangeTags;
        private List<string> listSelectOneZDV;
        private Dictionary<int, int> setValue;

        private Regex regexCombs;
        //private string pattern = @"[а-яA-Я]+\s+[№:]+\s+\d+";
        private string pattern = @"[а-яA-Я]+\s+[№:]+\s+\d*[^.]";
       
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
        private StatusOperation statusOperation;

        private string nameTU;
        private int numberTU;
        private int prevIndexComb = -1;
        private int indexComb = -1;
        private string lastCheckNumbComb = "";

        private bool isConnectExcel = false;

        public RepositoryImpl()
        {
            regexCombs = new Regex(pattern);

            dictCombs = new Dictionary<string, List<String>>();
            dictTags = new Dictionary<string, string>();
            dictErrorCombs = new Dictionary<string, List<String>>();

            listSelectZDVs = new List<string>();
            
            listDataChangeTags = new List<string>();
            dictDataChangeTags = new Dictionary<int, List<string>>();
            _dictDataChangeTags = new Dictionary<string, int>();

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

            statusOperation = new StatusOperation();
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
                    if (isConnectExcel)
                        ClearAllData();
                    SetRefExcelElement(pathCombFile);
                    GetNameTU();
                    isConnectExcel = true;
                    return isConnectExcel;
                }
                catch (Exception ex)
                {
                        MessageBox.Show(ex.Message);
                        workbook = null;
                }
            }
            isConnectExcel = false;
            return isConnectExcel;
        }

        private void SetRefExcelElement(string pathCombFile)
		{
            workbook = new XLWorkbook(pathCombFile);
            worksheetCombs = workbook.Worksheet(NAME_COMBS_LIST_EXCEL);
            worksheetTags = workbook.Worksheet(NAME_TAGS_LIST_EXCEL);
            worksheetSubScribe = workbook.Worksheet(NAME_SUBSCRIBE_LIST_EXCEL);
            worksheetSettings = workbook.Worksheet(NAME_SETTINGS_LIST_EXCEL);
        }

        private void ClearAllData()
		{
            workbook = null;
            worksheetCombs = null;
            worksheetTags = null;
            worksheetSubScribe = null;
            worksheetSettings = null;
            dictCombs.Clear();
            dictTags.Clear();
            dictErrorCombs.Clear();
            //listZDVs.Clear();
            listSelectZDVs.Clear();
            listAutoCheckZDVs.Clear();
            listDataChangeTags.Clear();
            dictDataChangeTags.Clear();
            _dictDataChangeTags.Clear();
            listSelectOneZDV.Clear();
            setValue.Clear();
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

        public void ReadCombinations(IExeState exeState)
        {
            this.exeState = exeState;
            new Thread(() =>
            {
                if (workbook == null) return;
                if (dictTags.Count == 0)
                    ReadTagsFromSource();
                dictCombs.Clear();
                string keyComb = "";
                int count = 0, row = 1, col = 1;
                try
                {
                    while (count < OFFSET_BETWEEN_COMBINATIONS)
                    {
                        string valueCell = worksheetCombs.Cell(row, col).Value.ToString();
                        count = String.IsNullOrEmpty(valueCell) ? count + 1 : 0;
                        CreateDictionaryCombinations(valueCell, ref keyComb);
                        row++;
                    }
                    exeState.CreateListBoxItems(dictCombs);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }).Start();
        }

        private void ReadTagsFromSource()
        {
            int row = 1, colDesc = 1, colTags = 2;
            dictTags.Clear();
            while (!String.IsNullOrEmpty(worksheetTags.Cell(row, colDesc).Value.ToString()))
            {
                string key = worksheetTags.Cell(row, colDesc).Value.ToString().ToLower();
                string value = worksheetTags.Cell(row, colTags).Value.ToString();
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

                        //MessageBox.Show("Список с сигналами скорее всего пустой");
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
                OPCItem item = opcItems.GetOPCItem((Int32)arrayClass.ServerHandles.GetValue(i));
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
            //opcDataChange.opcGroup.DataChange += ObjOPCGroup1_DataChange;
            opcState.tags = FormTagsVM();
            return opcState;
        }

        //для ViewModel формируем список тэгов, хэндлов для них и значений установки и снятия
        private List<TagForViewModel> FormTagsVM()
		{
            int handle = 1;
            List<TagForViewModel> tagVMs = new List<TagForViewModel>();
			foreach (string item in _dictDataChangeTags.Keys)
			{
                TagForViewModel tagVM = new TagForViewModel();
                tagVM.Handle = handle++;
                tagVM.Tag = item;
                indexComb = _dictDataChangeTags[item];
                SetCheckValue();
                tagVM.DrawDown = setValue[1];
                tagVM.Reset = setValue[2];
                tagVMs.Add(tagVM);
            }
            _dictDataChangeTags.Clear();
            indexComb = -1;
            return tagVMs;
        }

        private List<SignalState> ReadTagsValues(List<string> listTags, ArrayClass arrayClass, OPCSettings opcData, string nameGroup, string data = "dictionary")
        {
            InitArrays(arrayClass, listTags.Count);

            FormationArrays(arrayClass, listTags, data);

            InitOPC(arrayClass, nameGroup, listTags.Count, opcData);

            List<SignalState> listValues = ReadValues(opcData, arrayClass);
            RemoveGroupOPCRead(opcData, nameGroup);
            return listValues;
        }

        private void RemoveGroupOPCRead(OPCSettings opcData, string nameGroup)
		{
            if (nameGroup != GROUP_OPC_DATA_CHANGE)
            {
                opcGroups.Remove(GROUP_OPC_READ);
                opcData.opcGroup = null;
            }
		}

        private bool GetTagsSubscribe()
        {
            if (worksheetSubScribe != null)
            {
                int i = 2;
                while (worksheetSubScribe.Cell(i, 1).Value.ToString() != "")
                {
                    string tag = worksheetSubScribe.Cell(i, 1).Value.ToString();
                    int index = Convert.ToInt32(worksheetSubScribe.Cell(i, 2).Value);
                    listDataChangeTags.Add(tag);    //лист с тэгами для подписки во viewModel
                    List<string> listTag = new List<string>();
                    listTag.Add(tag);
                    dictDataChangeTags.Add(index, listTag); //формируем словарь с индексами и тэгами
                    _dictDataChangeTags.Add(tag, index); //формируем словарь с индексами и тэгами
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
            GetStatusOperation(ABORT_AUTO_CHECK, nameTU, isStopAutoImitation: true);
            exeState.GetStateExecute(statusOperation);
            backThread.Abort();
            backThread.Join();
            ResetSettingsBackThread();
            ResetOPCGroup();
            //opcGroups.Remove(GROUP_OPC_READ);
            //opcGroups.Remove(GROUP_OPC_WRITE);
            //opcGroups.RemoveAll();
            //opcGroups.Add(GROUP_OPC_DATA_CHANGE);
        }

        private void ResetSettingsBackThread()
		{
            backThread = null;
            backThreadId = 0;
        }

        private void ResetOPCGroup()
		{
            opcRead.opcGroup = null;
            opcWrite.opcGroup = null;
        }

        private void ExecuteAutoCheck()
        {
            var startTime = Stopwatch.StartNew();
            string keyCurrentComb;
            statusOperation.CountCombinations = 0;
            statusOperation.EllapsedTime = "00:00:00:00";
            try
            {
                if (dictCombs.Count == 0 || dictTags.Count == 0) return;
                foreach (KeyValuePair<string, List<string>> keyValuePair in dictCombs)
                {
                    keyCurrentComb = keyValuePair.Key;
                    indexComb = Convert.ToInt32(Regex.Match(Regex.Matches(keyCurrentComb, @"[а-яA-Я]+\s+[№:]+\s+\d+")[1].Value, @"\d+").Value);
                    SetCheckValue();
                    FormationListAutoCheckZDVs(keyValuePair.Value);
                    GetStatusOperation(EXECUTE_AUTO_CHECK, nameTU);
                    exeState.GetStateExecute(statusOperation);
                    ExecuteStep(command.SetStatusClose(), 6, setValue[1], keyCurrentComb);
                    foreach (var item in listAutoCheckZDVs)
                    {
                        GetStatusOperation(EXECUTE_AUTO_CHECK, nameTU, keyCurrentComb + "%" + item);
                        exeState.GetStateExecute(statusOperation);
                        listSelectOneZDV.Clear();
                        listSelectOneZDV.Add(item);
                        ExecuteStep(command.SetStatusOpen(), 7, setValue[2], keyCurrentComb, item);
                        ExecuteStep(command.SetStatusClose(), 7, setValue[1], keyCurrentComb, item);
                    }
                    ExecuteStep(command.SetStatusOpen(), 6, setValue[2], keyCurrentComb);
                    lastCheckNumbComb = Regex.Match(Regex.Matches(keyCurrentComb, @"[а-яA-Я]+\s+[№:]+\s+\d+")[0].Value, @"\d+").Value;
                    statusOperation.CountCombinations++;
                    statusOperation.EllapsedTime = GetTime(startTime);
                }
                startTime.Stop();
                GetStatusOperation(DONE_AUTO_CHECK, nameTU, isStopAutoImitation: true);
                exeState.GetStateExecute(statusOperation);
                ResetSettingsBackThread();
                ResetOPCGroup();
                MessageBox.Show(DONE_AUTO_CHECK);
            }
            catch (ThreadAbortException)
            {
                Imitation(command.SetStatusOpen(), 6);  //откроем все задвижки в комбинации, чтобы можно было начать проверку занаво 
                if (startTime.IsRunning)
                    startTime.Stop();
                MessageBox.Show(ABORT_AUTO_CHECK);
            }
			finally
			{
                OutputErrorData(GetTime(startTime));
			}
        }

        private void FormationListAutoCheckZDVs(List<string> listZDVs)
		{
            listAutoCheckZDVs.Clear();
            foreach (var value in listZDVs)
                listAutoCheckZDVs.Add(value);
        }

        private void GetStatusOperation(string state, string nameTU, string combination = null, bool isStopAutoImitation = false)
		{
            statusOperation.State = state;
            statusOperation.NameTU = nameTU;
            statusOperation.Combination = combination;
            statusOperation.IsStopAutoImitation = isStopAutoImitation;
		}

        private string GetTime(Stopwatch startTime)
		{
            var resultTime = startTime.Elapsed;
            // elapsedTime - строка, которая будет содержать значение затраченного времени
            return String.Format("{0:00}:{1:00}:{2:00}:{3:00}",
                resultTime.Days,
                resultTime.Hours,
                resultTime.Minutes,
                resultTime.Seconds);
        }

        private void SetValues(int drapdown, int reset)
		{
            setValue.Add(1, drapdown);  //сработка
            setValue.Add(2, reset);     //сброс сработки
        }

        //установить значения с которыми будут проверяться комбинации
        private void SetCheckValue()
		{
            if (prevIndexComb == indexComb && setValue.Count > 0)
                return;
            setValue.Clear();
            //перекрытие
            if (indexComb >= 0 && indexComb <= 4)
                SetValues(1, 0);
            //отсечение
            if (indexComb >= 1000 && indexComb <= 1100)
                SetValues(1, 0);
            //блокировка
            if (indexComb >= 2000 && indexComb <= 2012)
                SetValues(1, 2);
            //путь течения
            if (indexComb >= 4003 && indexComb <= 4008)
                SetValues(1, 2);
            //луппинг
            if (indexComb >= 5000 && indexComb <= 5005)
                SetValues(0, 1);

            prevIndexComb = indexComb;
		}

        private void ExecuteStep(int commandValue, int keyOperation, int currValueTag, string keyCurrentComb, string nameZDV = "")
        {
            Imitation(commandValue, keyOperation);                     //выполним операции с задвижками из комбинации согласно команде(закрыть, открыть)
            Thread.Sleep(DELAY);
			try
			{
                int valueTag = Convert.ToInt32(ReadTagsValues(dictDataChangeTags[indexComb], arrayRead, opcRead, GROUP_OPC_READ, "list")[0].Value);
                SetDictErrorCombs(currValueTag, keyCurrentComb, valueTag, nameZDV);
            }
            catch
			{
                MessageBox.Show("Заполните список сигналов для подписки во вкладке Subscribe");
			}
        }

        private void OutputErrorData(string elapsedTime)
        {
            if (File.Exists(path))
                File.Delete(path);

			try
			{
                using (FileStream fileStream = new FileStream(path, FileMode.Create))
                {
                    using (StreamWriter writer = new StreamWriter(fileStream))
                    {
                        WriteDataInFile(writer, elapsedTime);
                    }
                }
            }
			catch
			{
                MessageBox.Show("Закройте файл с результатами имитации");
			}
        }

        private void WriteDataInFile(StreamWriter writer, string elapsedTime)
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
            }
            writer.WriteLine($"Последняя проверенная комбинация: {lastCheckNumbComb}");
            writer.WriteLine($"Потраченное время на проверку: {elapsedTime}");
            writer.WriteLine(DONE_AUTO_CHECK);
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

        private void SetDictErrorCombs(int currValueTag, string keyCurrentComb, int valueTag, string nameZDV)
        {
            if (currValueTag != valueTag)
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
        //public void ObjOPCGroup1_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        //{
        //    arrayDataChange.ItemValues = ItemValues;
        //    arrayDataChange.Qualities = Qualities;
        //    arrayDataChange.TimeStamps = TimeStamps;
        //}
    }
}
