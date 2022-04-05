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
using System.Text.RegularExpressions;

namespace ImitComb
{
    class ViewModel: IExeState
    {
        private const string DONE_AUTO_CHECK = "Операция выполнена";
        private const string ABORT_AUTO_CHECK = "Операция прервана";
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
        private Button buttonAutoCheckAG2;
        private Button buttonAutoCheckAG3;
        private Button buttonAutoCheckCurrent;
        private Button buttonAutoCheckBlock;
        private Button buttonImitation;
        private Button buttonOpen;
        private Button buttonClose;
        private Button buttonOpening;
        private Button buttonClosing;
        private Button buttonMiddle;
        private Button buttonClearForm;
        private Rectangle blinkerBlockWay11;
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
        private bool stopAutoImitation;

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
            buttonAutoCheckAG2 = mainWindow.buttonAutoCheckAG2;
            buttonAutoCheckAG3 = mainWindow.buttonAutoCheckAG3;
            buttonImitation = mainWindow.buttonImitation;
            buttonOpen = mainWindow.buttonOpen;
            buttonClose = mainWindow.buttonClose;
            buttonOpening = mainWindow.buttonOpening;
            buttonClosing = mainWindow.buttonClosing;
            buttonMiddle = mainWindow.buttonMiddle;
            buttonClearForm = mainWindow.buttonClearForm;
            blinkerBlockWay11 = mainWindow.blinkerBlockWay11;
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
            if (!pathCombFile.Contains(".xlsx"))
                return pathCombFile = pathCombFile + ".xlsx";
            return pathCombFile;
        }

        public void CheckExcel(string pathCombFile)
        {
            if (checkExcel.CheckExcel(ParsePathFile(pathCombFile)))
            {
                GetNameServer();
                GetNameArea();
                SubScribeTags();
                if (opcState != null)
                    SetParamOPCDataChange();
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

        public void ReadCombinations()
        {
            dictCombs = readCombinations.ReadCombinations();
            if (dictCombs != null)
            {
                CreateListBoxItems();
                countCombinations = dictCombs.Keys.Count;
            }
            else
            {
                MessageBox.Show("Скорее всего введенные данные не верны");
            }
            labelCombs.Content = "Кол-во комбинаций: " + countCombinations;
        }

        private void CreateListBoxItems()
        {
            listBoxCombs.Items.Clear();
            foreach (var key in dictCombs.Keys)
                listBoxCombs.Items.Add(key);
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

            if (ItemValues.GetValue(1) != null && 2 == (Int32)ItemValues.GetValue(1))
            {
                SolidColorBrush solidColor = new SolidColorBrush(Colors.Green);
                blinkerBlockWay11.Fill = solidColor;
            } else
            {
                SolidColorBrush solidColor = new SolidColorBrush(Colors.Red);
                blinkerBlockWay11.Fill = solidColor;
            }
        }

        public void GetStateExecute(string state, string nameTU, string combination = null, bool stopAutoImitation = false)
        {
            this.stopAutoImitation = stopAutoImitation;
            
            if (!WorkWithButtons(nameTU)) return;
            
            mainWindow.Dispatcher.Invoke(() =>
            {
                if (stopAutoImitation)
                {
                    buttonAutoCheckCurrent.Content = "Автопроверка " + nameTU;
                }
                else
                {
                    buttonAutoCheckCurrent.Content = "Остановить\nавтопроверку " + nameTU;
                }
                labelStateAutoImitation.Content = state;
                if (!String.IsNullOrEmpty(combination))
                {
                    labelCombAutoImitation.Content = combination.Split('%')[0];
                    labelZDVAutoImitation.Content = combination.Split('%')[1];
                }
                else
                {
                    labelCombAutoImitation.Content = "";
                    labelZDVAutoImitation.Content = "";
                }
                if (state == DONE_AUTO_CHECK || state == ABORT_AUTO_CHECK)
                    UnBlockElementsForm();
                else
                    BlockElementsForm();
            });
            
        }

        private bool WorkWithButtons(string nameTU)
		{
            if (!String.IsNullOrEmpty(nameTU))
            {
                int numberTU = Convert.ToInt16(Regex.Match(nameTU, @"\d").Value);
                switch (numberTU)
                {
                    case 2:
                        SetWorkButtons(buttonAutoCheckAG2, buttonAutoCheckAG3);
                        return true;
                    case 3:
                        SetWorkButtons(buttonAutoCheckAG3, buttonAutoCheckAG2);
                        return true;
                }
            }
            return false;
        }

        private void SetWorkButtons(Button buttonCurrent, Button buttonBlock)
		{
            buttonAutoCheckCurrent = buttonCurrent;
            buttonAutoCheckBlock = buttonBlock;
        }

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
            buttonAutoCheckBlock.IsEnabled = false;
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
            buttonAutoCheckBlock.IsEnabled = true;
        }
    }
}
