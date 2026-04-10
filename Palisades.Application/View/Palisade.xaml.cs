using Palisades.Model;
using Palisades.ViewModel;
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Palisades.View
{
    public partial class Palisade : Window
    {
        private const string TabDragDataFormat = "Palisades.TabIdentifier";
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_APPWINDOW = 0x00040000;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;

        private readonly PalisadeViewModel viewModel;
        private Point? tabDragStartPoint;
        private string? draggedTabIdentifier;

        public Palisade(PalisadeViewModel defaultModel)
        {
            InitializeComponent();
            DataContext = defaultModel;
            viewModel = defaultModel;
            viewModel.PropertyChanged += ViewModel_PropertyChanged;
            SourceInitialized += Palisade_SourceInitialized;
            Closed += Palisade_Closed;
            TrySetWindowIcon();
            Show();
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        private void Palisade_SourceInitialized(object? sender, EventArgs e)
        {
            ApplyAltTabVisibility();
        }

        private void Palisade_Closed(object? sender, EventArgs e)
        {
            viewModel.PropertyChanged -= ViewModel_PropertyChanged;
            SourceInitialized -= Palisade_SourceInitialized;
            Closed -= Palisade_Closed;
        }

        private void ViewModel_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(e.PropertyName) || e.PropertyName == nameof(PalisadeViewModel.ShowInAltTab))
            {
                ApplyAltTabVisibility();
            }
        }

        private void ApplyAltTabVisibility()
        {
            IntPtr handle = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero)
            {
                return;
            }

            int extendedStyle = GetWindowLong(handle, GWL_EXSTYLE);
            extendedStyle = viewModel.ShowInAltTab
                ? (extendedStyle | WS_EX_APPWINDOW) & ~WS_EX_TOOLWINDOW
                : (extendedStyle | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW;

            SetWindowLong(handle, GWL_EXSTYLE, extendedStyle);
            SetWindowPos(handle, IntPtr.Zero, 0, 0, 0, 0,
                SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
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
            if (viewModel.IsLayoutLocked)
            {
                return;
            }

            base.OnMouseLeftButtonDown(e);

            try
            {
                DragMove();
                PalisadesManager.TryJoinPalisadeByOverlap(viewModel.Identifier);
            }
            catch (InvalidOperationException)
            {
                // Ignore drag cancellation if the mouse was released before a move started.
            }
        }

        private void TabButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Button button)
            {
                return;
            }

            draggedTabIdentifier = ResolveTabIdentifier(button.DataContext);
            tabDragStartPoint = e.GetPosition(this);
        }

        private void TabButton_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (sender is not Button button
                || e.LeftButton != MouseButtonState.Pressed
                || tabDragStartPoint == null
                || string.IsNullOrWhiteSpace(draggedTabIdentifier))
            {
                return;
            }

            Point currentPosition = e.GetPosition(this);
            if (Math.Abs(currentPosition.X - tabDragStartPoint.Value.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(currentPosition.Y - tabDragStartPoint.Value.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            string dragIdentifier = draggedTabIdentifier;
            DataObject dataObject = new();
            dataObject.SetData(TabDragDataFormat, dragIdentifier);

            DragDropEffects result = DragDrop.DoDragDrop(button, dataObject, DragDropEffects.Move);
            if (result == DragDropEffects.None && !string.IsNullOrWhiteSpace(dragIdentifier))
            {
                PalisadesManager.SplitPalisadeFromTabs(dragIdentifier, GetCurrentMouseScreenPosition());
            }

            tabDragStartPoint = null;
            draggedTabIdentifier = null;
            e.Handled = true;
        }

        private void TabButton_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(TabDragDataFormat) ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        }

        private void TabButton_Drop(object sender, DragEventArgs e)
        {
            if (sender is not Button button || !e.Data.GetDataPresent(TabDragDataFormat))
            {
                return;
            }

            string? draggedIdentifier = e.Data.GetData(TabDragDataFormat) as string;
            string targetIdentifier = ResolveTabIdentifier(button.DataContext);
            if (string.IsNullOrWhiteSpace(draggedIdentifier) || string.IsNullOrWhiteSpace(targetIdentifier))
            {
                return;
            }

            Point dropPosition = e.GetPosition(button);
            bool placeAfter = dropPosition.X >= button.ActualWidth / 2d;
            PalisadesManager.ReorderTabbedFence(draggedIdentifier, targetIdentifier, placeAfter);

            tabDragStartPoint = null;
            draggedTabIdentifier = null;
            e.Handled = true;
        }

        private Point GetCurrentMouseScreenPosition()
        {
            System.Drawing.Point cursorPosition = System.Windows.Forms.Control.MousePosition;
            Point screenPoint = new(cursorPosition.X, cursorPosition.Y);

            PresentationSource? source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget != null)
            {
                screenPoint = source.CompositionTarget.TransformFromDevice.Transform(screenPoint);
            }

            return screenPoint;
        }

        private static string ResolveTabIdentifier(object? dataContext)
        {
            return dataContext is PalisadeTabInfo tabInfo ? tabInfo.Identifier : string.Empty;
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

        private void ShortcutRenameTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is not TextBox textBox)
            {
                return;
            }

            textBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                textBox.Focus();
                textBox.SelectAll();
            }), DispatcherPriority.Input);
        }

        private void ShortcutRenameTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.DataContext is Shortcut shortcut)
            {
                viewModel.CommitRenameShortcut(shortcut);
            }
        }

        private void ShortcutRenameTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (sender is not TextBox textBox || textBox.DataContext is not Shortcut shortcut)
            {
                return;
            }

            Key pressedKey = e.Key == Key.System ? e.SystemKey : e.Key;
            if (pressedKey == Key.Enter)
            {
                viewModel.CommitRenameShortcut(shortcut);
                e.Handled = true;
                return;
            }

            if (pressedKey == Key.Escape)
            {
                viewModel.CancelRenameShortcut(shortcut);
                e.Handled = true;
            }
        }
    }
}
