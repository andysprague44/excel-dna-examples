using System;
using System.Windows.Forms;
using MsgBox = System.Windows.Forms.MessageBox;

namespace MessageBoxAddin.Forms
{
    public class ExcelWinFormsUtil : IExcelWinFormsUtil
    {
        public DialogResult MessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return ShowModal(parentWindow => MsgBox.Show(parentWindow, text, caption, buttons, icon));
        }

        private static DialogResult ShowModal(Func<IWin32Window, DialogResult> dialogFunc)
        {
            var parentWindow = new NativeWindow();
            parentWindow.AssignHandle(ExcelDna.Integration.ExcelDnaUtil.WindowHandle);

            try
            {
                return dialogFunc(parentWindow);
            }
            finally
            {
                parentWindow.ReleaseHandle();
            }
        }
    }
}