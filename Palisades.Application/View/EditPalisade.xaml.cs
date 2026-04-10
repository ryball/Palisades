using Palisades.ViewModel;
using System.ComponentModel;
using System.Windows;

namespace Palisades.View
{
    /// <summary>
    /// Logique d'interaction pour EditPalisade.xaml
    /// </summary>
    public partial class EditPalisade : Window
    {
        private bool settingsSessionStarted;
        private bool settingsSaved;

        public EditPalisade()
        {
            InitializeComponent();
            Loaded += EditPalisade_Loaded;
        }

        private void EditPalisade_Loaded(object sender, RoutedEventArgs e)
        {
            if (settingsSessionStarted)
            {
                return;
            }

            if (DataContext is PalisadeViewModel viewModel)
            {
                viewModel.BeginSettingsEditSession();
                settingsSessionStarted = true;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is PalisadeViewModel viewModel)
            {
                viewModel.CommitSettingsEditSession();
            }

            settingsSaved = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (!settingsSaved && settingsSessionStarted && DataContext is PalisadeViewModel viewModel)
            {
                viewModel.CancelSettingsEditSession();
            }

            base.OnClosing(e);
        }
    }
}
