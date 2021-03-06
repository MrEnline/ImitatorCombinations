using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Shapes;
using System.Windows.Controls;
using OPCAutomation;
using ImitComb.domain;
using ImitComb.data;
using ImitComb.domain.Entity;
using ImitComb.presentation;
using System.Windows.Media;

namespace ImitComb
{
    class ViewModel: IExeState
    {
        private const string DONE_AUTO_CHECK = "Операция выполнена";
        private const string ABORT_AUTO_CHECK = "Операция прервана";

        private const string ALARM = "alarm";
        private const string CUT_OFF = "cutoff";
        private const string BLOCK_WAY = "blockway";
        private const string WAY_TO_RP_NPS_LAST = "waytorpnpslast";
        private const string LOOPING = "looping";

        private const string XLSX = ".xlsx";

        private MainWindow mainWindow;
        private Dictionary<string, List<string>> dictCombs;
        private ListBox listBoxCombs;
        private ListBox listBoxZDVs;
        private ListBox listBoxSelectZDVs;
        private TextBox textBoxNameServer;
        private TextBox textBoxPathCombFile;
        private TextBox textBoxArea;
        private CheckBox checkBoxClosing;
        private CheckBox checkBoxClosed;
        private CheckBox checkBoxOpen;
        private Label labelResultImitation;
        private Label labelCombs;
        private Label labelStateAutoImitation;
        private Label labelCombAutoImitation;
        private Label labelZDVAutoImitation;
        private Label labelCountCheckComb;
        private Label labelEllapsedTime;
        private Button buttonAutoCheck;
        private Button buttonAutoCheckBlock;
        private Button buttonImitation;
        private Button buttonOpen;
        private Button buttonClose;
        private Button buttonOpening;
        private Button buttonClosing;
        private Button buttonMiddle;
        private Button buttonClearForm;
        private Rectangle blinkerBlockWay;
        private Rectangle blinkerAlarm;
        private Rectangle blinkerCutOff;
        private Rectangle blinkerLooping;
        private Rectangle blinkerFlowPath;
        private RepositoryImpl repository;
        private ReadCombinationsUseCase readCombinations;
        private CheckExcelUseCase checkExcel;
        private GetNameServerUseCase getNameServer;
        private GetNameAreaUseCase getNameArea;
        private ImitationUseCase imitation;
        private AutoCheckUseCase autoCheck;
        private CreateChoiceCombZDVsUseCase createChoiceCombZDVs;
        private CreateListSelectZDVsUseCase createListSelectZDVs;
        private ClearListSelectZDVsUseCase clearListSelectZDVs;
        private SubScribeTagsUseCase subScribeTags;
        private ConnectServerUseCase connectServer;
        private Command command;
        private int countCombinations;
        private OPCState opcState;
        private OPCGroup opcGroupDataChange;
        private bool isStopAutoImitation;

        public ViewModel(MainWindow mainWindow)
        {
            this.mainWindow = mainWindow;
            listBoxCombs = mainWindow.listBoxComb;
            listBoxZDVs = mainWindow.listBoxZDVs;
            listBoxSelectZDVs = mainWindow.listBoxSelectZDV;
            textBoxNameServer = mainWindow.textBoxNameServer;
            textBoxPathCombFile = mainWindow.textBoxPathCombFile;
            checkBoxClosing = mainWindow.checkBoxClosing;
            checkBoxClosed = mainWindow.checkBoxClosed;
            checkBoxOpen = mainWindow.checkBoxOpen;
            textBoxArea = mainWindow.textBoxArea;
            labelResultImitation = mainWindow.labelResultImitation;
            labelCombs = mainWindow.labelCombs;
            labelStateAutoImitation = mainWindow.labelStateAutoImitation;
            labelCombAutoImitation = mainWindow.labelCombAutoImitation;
            labelZDVAutoImitation = mainWindow.labelZDVAutoImitation;
            labelCountCheckComb = mainWindow.labelCountCheckComb;
            labelEllapsedTime = mainWindow.labelEllapsedTime;
            buttonAutoCheck = mainWindow.buttonAutoCheck;
            buttonImitation = mainWindow.buttonImitation;
            buttonOpen = mainWindow.buttonOpen;
            buttonClose = mainWindow.buttonClose;
            buttonOpening = mainWindow.buttonOpening;
            buttonClosing = mainWindow.buttonClosing;
            buttonMiddle = mainWindow.buttonMiddle;
            buttonClearForm = mainWindow.buttonClearForm;
            blinkerBlockWay = mainWindow.blinkerBlockWay;
            blinkerAlarm = mainWindow.blinkerAlarm;
            blinkerCutOff = mainWindow.blinkerCutOff;
            blinkerFlowPath = mainWindow.blinkerFlowPath;
            blinkerLooping = mainWindow.blinkerLooping;
            checkBoxClosing.IsChecked = true;
            repository = new RepositoryImpl();
            readCombinations = new ReadCombinationsUseCase(repository);
            checkExcel = new CheckExcelUseCase(repository);
            getNameServer = new GetNameServerUseCase(repository);
            getNameArea = new GetNameAreaUseCase(repository);
            imitation = new ImitationUseCase(repository);
            autoCheck = new AutoCheckUseCase(repository);
            createChoiceCombZDVs = new CreateChoiceCombZDVsUseCase(repository);
            createListSelectZDVs = new CreateListSelectZDVsUseCase(repository);
            clearListSelectZDVs = new ClearListSelectZDVsUseCase(repository);
            connectServer = new ConnectServerUseCase(repository);
            subScribeTags = new SubScribeTagsUseCase(repository);
            command = new Command();
            countCombinations = 0;
        }

        private string ParsePathFile(string pathCombFile = "")
        {
            if (!pathCombFile.Contains(XLSX))
                return pathCombFile + XLSX;
            return pathCombFile;
        }

        public void CheckExcel(string pathCombFile)
        {
            bool isCheckExcel = checkExcel.CheckExcel(ParsePathFile(pathCombFile));
            if (isCheckExcel)
            {
                GetNameServer();
                GetNameArea();
                SubScribeTags();
                if (opcState != null)
                    SetParamOPCDataChange();
                readCombinations.ReadCombinations(this);
            }
            else
            {
                listBoxCombs.Items.Clear();
                textBoxArea.Clear();
                textBoxNameServer.Clear();
            }
        }

        private void SetParamOPCDataChange()
        {
            opcGroupDataChange = null;
            if (opcState.ConnectOPC)
            {
                opcGroupDataChange = opcState.OpcGroup;
                opcGroupDataChange.DataChange += ObjOPCGroup_DataChange;
                opcGroupDataChange.UpdateRate = 100;
                opcGroupDataChange.IsActive = true;
                opcGroupDataChange.IsSubscribed = true;
            }
        }

        public void GetNameServer(string nameServer="")
        {
            textBoxNameServer.Text = getNameServer.GetNameServer(nameServer);
            connectServer.ConnectServer();
        }

        public void GetNameArea(string nameArea = "")
        {
            textBoxArea.Text = getNameArea.GetNameArea(nameArea);
        }

        //public void ReadCombinations()
        //{
        //    //dictCombs = readCombinations.ReadCombinations();
        //    //CreateListBoxItems();
        //    readCombinations.ReadCombinations(this);
        //}

        public void CreateListBoxItems(Dictionary<string, List<string>> dictCombs)
        {
            this.dictCombs = dictCombs;
            if (dictCombs != null)
            {
                mainWindow.Dispatcher.Invoke(() =>
                {
                    listBoxCombs.ItemsSource = null;
                    BlockElementsForm();
                    listBoxCombs.ItemsSource = dictCombs.Keys;
                    labelCombs.Content = "Кол-во комбинаций: " + dictCombs.Keys.Count;
                    UnBlockElementsForm();
                });
            }

            //listBoxCombs.Items.Clear();
            //foreach (var key in dictCombs.Keys)
            //    listBoxCombs.Items.Add(key);
        }

        public void CreateListBoxZDVs(string keyCombination)
        {
            try
            {
                List<string> listZDVs = dictCombs[keyCombination];
                listBoxZDVs.Items.Clear();
                foreach (var zdv in listZDVs)
                    listBoxZDVs.Items.Add(zdv);
                createChoiceCombZDVs.CreateChoiceCombZDVs(listZDVs);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void CreateListBoxSelectZDV(string nameZDV)
        {
            try
            {
                List<string> listSelectZDVs = createListSelectZDVs.CreateListSelectZDVs(nameZDV); ;
                listBoxSelectZDVs.Items.Clear();
                if (listSelectZDVs.Count > 0)
                    foreach (var zdv in listSelectZDVs)
                        listBoxSelectZDVs.Items.Add(zdv);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ClearListSelectZDVs()
        {
            clearListSelectZDVs.ClearListSelectZDVs();
            listBoxSelectZDVs.Items.Clear();
        }

        public void Imitation()
        {
            labelResultImitation.Content = "";
            if (listBoxZDVs.Items.Count == 0)
            {
                MessageBox.Show("Поле с комбинациями пустое");
                labelResultImitation.Content = "Ошибка";
                return;
            }
            int valueCommand = GetCommand(); 
            labelResultImitation.Content = imitation.Imitation(valueCommand, 0);
        }

        private int GetCommand()
        {
            if (checkBoxClosing.IsChecked.Value) return command.SetStatusClosing();
            if (checkBoxClosed.IsChecked.Value) return command.SetStatusClose();
            if (checkBoxOpen.IsChecked.Value) return command.SetStatusOpen();
            return command.SetStatusOpen();
        }

        public void OpenZDVs()
        {
            labelResultImitation.Content = imitation.Imitation(command.SetStatusOpen(), 1);
        }

        public void ClosingZDVs()
        {
            labelResultImitation.Content = imitation.Imitation(command.SetStatusClosing(), 2);
        }

        public void CloseZDVs()
        {
            labelResultImitation.Content = imitation.Imitation(command.SetStatusClose(), 3);
        }

        public void OpeningZDVs()
        {
            labelResultImitation.Content = imitation.Imitation(command.SetStatusOpening(), 4);
        }

        public void MiddleZDVs()
        {
            labelResultImitation.Content = imitation.Imitation(command.SetStatusMiddle(), 1);
        }

        public void SubScribeTags()
        {
            opcState = subScribeTags.SubScribeTags();
        }

        public void AutoCheck()
        {
            autoCheck.AutoCheck(this);
            listBoxSelectZDVs.Items.Clear();
            listBoxZDVs.Items.Clear();
        }

        //обработчик события
        public void ObjOPCGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
			for (int i = 1; i <= NumItems; i++)
			{
                int itemHandle = (Int32)ClientHandles.GetValue(i);
                Rectangle blinker = GetBlinker(opcState.tags[itemHandle - 1].Tag);
                if (ItemValues.GetValue(i) != null && opcState.tags[itemHandle - 1].DrawDown == Convert.ToInt32(ItemValues.GetValue(i)))
                    SetColorBlinker(blinker, Colors.Red);
                if (ItemValues.GetValue(i) != null && opcState.tags[itemHandle - 1].Reset == Convert.ToInt32(ItemValues.GetValue(i)))
                    SetColorBlinker(blinker, Colors.Green);
            }
        }

        private void SetColorBlinker(Rectangle blinker, Color color)
		{
            SolidColorBrush solidColor = new SolidColorBrush(color);
            blinker.Fill = solidColor;
        }

        private Rectangle GetBlinker(string tag)
		{
            Rectangle blinker = null;
            if (tag.ToLower().Contains(ALARM))
                blinker = blinkerAlarm;
            if (tag.ToLower().Contains(CUT_OFF))
                blinker = blinkerCutOff;
            if (tag.ToLower().Contains(BLOCK_WAY))
                blinker = blinkerBlockWay;
            if (tag.ToLower().Contains(WAY_TO_RP_NPS_LAST))
                blinker = blinkerFlowPath;
            if (tag.ToLower().Contains(LOOPING))
                blinker = blinkerLooping;
            return blinker;
        }

        public void GetStateExecute(StatusOperation status)
        {
            isStopAutoImitation = status.IsStopAutoImitation;

            //if (!WorkWithButtons(nameTU)) return;
            mainWindow.Dispatcher.Invoke(() =>
            {
                buttonAutoCheck.Content = isStopAutoImitation ? $"Автопроверка {status.NameTU}" : $"     Остановить\nавтопроверку {status.NameTU}";
                labelStateAutoImitation.Content = status.State;
                labelCombAutoImitation.Content = String.IsNullOrEmpty(status.Combination) ? status.Combination : status.Combination.Split('%')[0];
                labelZDVAutoImitation.Content = String.IsNullOrEmpty(status.Combination) ? status.Combination : status.Combination.Split('%')[1];
                labelCountCheckComb.Content = $"Количество проверенных комбинаций: {status.CountCombinations}";
                labelEllapsedTime.Content = $"Затраченное время: {status.EllapsedTime}";

                if (status.State == DONE_AUTO_CHECK || status.State == ABORT_AUTO_CHECK)
                    UnBlockElementsForm();
                else
                    BlockElementsForm();
            });
            
        }

  //      private bool WorkWithButtons(string nameTU)
		//{
  //          if (!String.IsNullOrEmpty(nameTU))
  //          {
  //              int numberTU = Convert.ToInt16(Regex.Match(nameTU, @"\d").Value);
  //              switch (numberTU)
  //              {
  //                  case 2:
  //                      SetWorkButtons(buttonAutoCheckAG2, buttonAutoCheckAG3);
  //                      return true;
  //                  case 3:
  //                      SetWorkButtons(buttonAutoCheckAG3, buttonAutoCheckAG2);
  //                      return true;
  //              }
  //          }
  //          return false;
  //      }

  //      private void SetWorkButtons(Button buttonCurrent, Button buttonBlock)
		//{
  //          buttonAutoCheckCurrent = buttonCurrent;
  //          buttonAutoCheckBlock = buttonBlock;
  //      }

        private void BlockElementsForm()
        {
            listBoxCombs.IsEnabled = false;
            textBoxNameServer.IsEnabled = false;
            textBoxPathCombFile.IsEnabled = false;
            textBoxArea.IsEnabled = false;
            buttonImitation.IsEnabled = false;
            buttonOpen.IsEnabled = false;
            buttonClose.IsEnabled = false;
            buttonOpening.IsEnabled = false;
            buttonClosing.IsEnabled = false;
            buttonMiddle.IsEnabled = false;
            buttonClearForm.IsEnabled = false;
    }

        private void UnBlockElementsForm()
        {
            listBoxCombs.IsEnabled = true;
            textBoxNameServer.IsEnabled = true;
            textBoxPathCombFile.IsEnabled = true;
            textBoxArea.IsEnabled = true;
            buttonImitation.IsEnabled = true;
            buttonOpen.IsEnabled = true;
            buttonClose.IsEnabled = true;
            buttonOpening.IsEnabled = true;
            buttonClosing.IsEnabled = true;
            buttonMiddle.IsEnabled = true;
            buttonClearForm.IsEnabled = true;
        }
    }
}
