using System.Windows.Forms;

namespace MessageBoxAddin.Forms
{
    public interface IExcelWinFormsUtil
    {
        DialogResult ShowForm(Form form);
        DialogResult MessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon);
    }
}
