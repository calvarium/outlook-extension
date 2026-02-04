using System.Windows;
using System.Windows.Controls;

namespace outlook_extension.UI.Wpf.Controls.Repeater
{
    public class ItemsRepeater : ItemsControl
    {
        public static readonly DependencyProperty OrientationProperty = DependencyProperty.Register(
            nameof(Orientation),
            typeof(Orientation),
            typeof(ItemsRepeater),
            new PropertyMetadata(Orientation.Vertical, OnOrientationChanged));

        public ItemsRepeater()
        {
            VirtualizingPanel.SetIsVirtualizing(this, true);
            VirtualizingPanel.SetVirtualizationMode(this, VirtualizationMode.Recycling);
            ScrollViewer.SetCanContentScroll(this, true);
            UpdateItemsPanel();
        }

        public Orientation Orientation
        {
            get => (Orientation)GetValue(OrientationProperty);
            set => SetValue(OrientationProperty, value);
        }

        private static void OnOrientationChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ItemsRepeater repeater)
            {
                repeater.UpdateItemsPanel();
            }
        }

        private void UpdateItemsPanel()
        {
            var factory = new FrameworkElementFactory(typeof(VirtualizingStackPanel));
            factory.SetValue(VirtualizingStackPanel.OrientationProperty, Orientation);
            ItemsPanel = new ItemsPanelTemplate(factory);
        }
    }
}
