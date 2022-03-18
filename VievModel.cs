using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Shapes;
using System.Windows.Controls;
using OPCAutomation;
using ImitComb.domain;
using ImitComb.data;
using ImitComb.domain.Entity;
using System.Windows.Media;

namespace ImitComb
{
    class ViewModel
    {
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
            autoCheck.AutoCheck();
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
    }
}
