using System.Windows.Forms;

namespace MessageBoxAddin.Forms
{
    public interface IExcelWinFormsUtil
    {
        DialogResult MessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon);
    }
}
