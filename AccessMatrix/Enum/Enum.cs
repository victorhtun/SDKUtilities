using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessMatrix.Enum
{
    public enum MsgBoxType
    {
        WinMsgBox = 0,
        B1MsgBox = 1,
        B1StatusBar = 2
    };

    public enum DocType
    {
        Item = 0,
        Service = 1,
        Other = 2
    };

    public enum IsSingle
    {
        No = 0,
        Yes = 1
    };

    public enum ControlType
    {
        TextBox = 1,
        CheckBox = 2,
        DropdowwList = 3

    };

    public enum EventType
    {
        Blur = 0,
        KeyPress = 1,
        ButtonClick = 2,
        SelectedValueChange = 3
    };

    public enum RequiredCustomized
    {
        No = 0,
        Yes = 1
    };
}
