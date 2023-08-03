using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;

namespace LibSrd
{
    public static class WpfHelpers
    {
        /// <summary>
        /// [OBSELETE] -> Use grid row and column definitions
        /// Fixes the awful offset of the wpf menus. 
        /// https://www.red-gate.com/simple-talk/blogs/wpf-menu-displays-to-the-left-of-the-window/
        /// </summary>
        public static void DropDownMenuFix()
        {
            var menuDropAlignmentField = typeof(SystemParameters).GetField("_menuDropAlignment", BindingFlags.NonPublic | BindingFlags.Static);
            Action setAlignmentValue = () => { if (SystemParameters.MenuDropAlignment && menuDropAlignmentField != null) menuDropAlignmentField.SetValue(null, false); };
            setAlignmentValue();
            SystemParameters.StaticPropertyChanged += (sender, e) => { setAlignmentValue(); };
        }

        /// <summary>
        /// [OBSELETE] -> Use content control rather than frame!
        ///  Removes the backspace key shortcut to navigate back.
        /// </summary>
        /// <param name="frame">Input frame to be fixed.</param>
        public static void BackspaceNavigationFix(Frame frame)
        {
            frame.Navigated += (sender, e) =>
            {
                if (((Frame)sender).CanGoBack)
                    ((Frame)sender).RemoveBackEntry();
            };
        }
    }
}