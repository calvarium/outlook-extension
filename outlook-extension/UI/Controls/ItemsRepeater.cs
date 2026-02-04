using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace outlook_extension.UI.Controls
{
    public class ItemsRepeater : ListBox
    {
        public ItemsRepeater()
        {
            SetValue(VirtualizingPanel.IsVirtualizingProperty, true);
            SetValue(VirtualizingPanel.VirtualizationModeProperty, VirtualizationMode.Recycling);
            SetValue(ScrollViewer.CanContentScrollProperty, true);
        }
    }
}
