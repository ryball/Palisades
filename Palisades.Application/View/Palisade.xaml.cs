using Palisades.ViewModel;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace Palisades.View
{
    public partial class Palisade : Window
    {
        private readonly PalisadeViewModel viewModel;
        public Palisade(PalisadeViewModel defaultModel)
        {
            InitializeComponent();
            DataContext = defaultModel;
            viewModel = defaultModel;
            TrySetWindowIcon();
            Show();
        }

        private void TrySetWindowIcon()
        {
            try
            {
                string iconPath = Path.Combine(AppContext.BaseDirectory, "Ressources", "icon.ico");
                if (!File.Exists(iconPath))
                {
                    return;
                }

                Icon = BitmapFrame.Create(new Uri(iconPath, UriKind.Absolute));
            }
            catch
            {
                // Ignore icon loading errors so fence creation never fails.
            }
        }

        private void Header_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            DragMove();
        }

        private void ShortcutGroup_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is not Expander expander)
            {
                return;
            }

            string? groupName = (expander.DataContext as CollectionViewGroup)?.Name?.ToString();
            expander.IsExpanded = viewModel.IsGroupExpanded(groupName);
        }

        private void ShortcutGroup_Expanded(object sender, RoutedEventArgs e)
        {
            if (!ReferenceEquals(sender, e.OriginalSource) || sender is not Expander expander)
            {
                return;
            }

            string? groupName = (expander.DataContext as CollectionViewGroup)?.Name?.ToString();
            viewModel.SetGroupExpanded(groupName, true);
        }

        private void ShortcutGroup_Collapsed(object sender, RoutedEventArgs e)
        {
            if (!ReferenceEquals(sender, e.OriginalSource) || sender is not Expander expander)
            {
                return;
            }

            string? groupName = (expander.DataContext as CollectionViewGroup)?.Name?.ToString();
            viewModel.SetGroupExpanded(groupName, false);
        }
    }
}
