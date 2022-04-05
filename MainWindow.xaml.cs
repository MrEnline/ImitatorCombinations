using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ImitComb
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ViewModel viewModel;
        public MainWindow()
        {
            InitializeComponent();
            viewModel = new ViewModel(this);
            //viewModel.ReadCombinations(textBoxPathCombFile.Text);
            listBoxComb.SelectionChanged += ListBoxComb_SelectionChanged;
            listBoxZDVs.SelectionChanged += ListBoxZDVs_SelectionChanged;
            textBoxPathCombFile.KeyDown += TextBoxPathCombFile_KeyDown;
            checkBoxClosing.Checked += CheckBoxClosing_Checked;
            checkBoxClosed.Checked += CheckBoxClosed_Checked;
            checkBoxOpen.Checked += CheckBoxOpen_Checked;
            textBoxNameServer.KeyDown += TextBoxNameServer_KeyDown;
            textBoxArea.KeyDown += TextBoxArea_KeyDown;
            buttonImitation.Click += ButtonImitation_Click;
            buttonImitation.Content = "Имитировать для всей\n       комбинации";
            buttonClearForm.Click += ButtonClearForm_Click;
            buttonOpen.Click += ButtonOpen_Click;
            buttonClose.Click += ButtonClose_Click;
            buttonOpening.Click += ButtonOpening_Click;
            buttonClosing.Click += ButtonClosing_Click;
            buttonMiddle.Click += ButtonMiddle_Click;
            buttonAutoCheckAG2.Click += ButtonAutoCheck_Click;
            buttonAutoCheckAG3.Click += ButtonAutoCheck_Click;
        }

        private void ListBoxComb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ListBox).SelectedItem != null)
            {
                string curItem = (sender as ListBox).SelectedItem.ToString();
                viewModel.ClearListSelectZDVs();
                viewModel.CreateListBoxZDVs(curItem);
            }
        }

        private void ListBoxZDVs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ListBox).SelectedItem != null)
            {
                string curItem = (sender as ListBox).SelectedItem.ToString();
                viewModel.CreateListBoxSelectZDV(curItem);
            }
        }

        private void TextBoxPathCombFile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                viewModel.CheckExcel((sender as TextBox).Text);
                viewModel.ReadCombinations();
            }
        }

        private void TextBoxNameServer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                viewModel.GetNameServer((sender as TextBox).Text);
        }

        private void TextBoxArea_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                viewModel.GetNameArea((sender as TextBox).Text);
        }

        private void ButtonImitation_Click(object sender, RoutedEventArgs e)
        {
            viewModel.Imitation();
        }

        private void ButtonAutoCheck_Click(object sender, RoutedEventArgs e)
        {
            viewModel.AutoCheck();
        }

        private void CheckBoxClosed_Checked(object sender, RoutedEventArgs e)
        {
            checkBoxClosing.IsChecked = false;
            checkBoxOpen.IsChecked = false;
        }

        private void CheckBoxClosing_Checked(object sender, RoutedEventArgs e)
        {
            checkBoxClosed.IsChecked = false;
            checkBoxOpen.IsChecked = false;
        }

        private void CheckBoxOpen_Checked(object sender, RoutedEventArgs e)
        {
            checkBoxClosing.IsChecked = false;
            checkBoxClosed.IsChecked = false;
        }

        private void ButtonClearForm_Click(object sender, RoutedEventArgs e)
        {
            viewModel.ClearListSelectZDVs();
        }

        private void ButtonOpen_Click(object sender, RoutedEventArgs e)
        {
            viewModel.OpenZDVs();
        }

        private void ButtonOpening_Click(object sender, RoutedEventArgs e)
        {
            viewModel.OpeningZDVs();
        }

        private void ButtonClosing_Click(object sender, RoutedEventArgs e)
        {
            viewModel.ClosingZDVs();
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            viewModel.CloseZDVs();
        }

        private void ButtonMiddle_Click(object sender, RoutedEventArgs e)
        {
            viewModel.MiddleZDVs();
        }
    }
}
