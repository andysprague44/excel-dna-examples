using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using MessageBoxAddin.Forms;

namespace MessageBoxAddin
{
    public class ExcelController : IDisposable
    {
        private readonly IRibbonUI _modelingRibbon;
        private readonly IExcelWinFormsUtil _excelWinFormsUtil;
        private readonly Application _excel;

        public ExcelController(Application excel, IRibbonUI modelingRibbon, IExcelWinFormsUtil excelWinFormsUtil)
        {
            _modelingRibbon = modelingRibbon;
            _excel = excel;
            _excelWinFormsUtil = excelWinFormsUtil;
        }

        public void PressMe()
        {
            var dialogResult = _excelWinFormsUtil.MessageBox(
                "This is a message box asking for your input - write something?",
                "Choose Option",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            switch (dialogResult)
            {
                case DialogResult.Yes:
                    _excel.Range("A1").Value = "Yes chosen";
                    break;
                case DialogResult.Cancel:
                    _excel.Range("A1").Value = "Canceled";
                    break;
                case DialogResult.No:
                    _excel.Range("A1").Value = null;
                    break;
            }
        }

        public void OnPressMeBackgroundThread(int delay)
        {
            Task.Factory.StartNew(
                () => RunBackgroundThread(delay),
                CancellationToken.None,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Current
            );
        }

        public void RunBackgroundThread(int delay)
        {
            Thread.Sleep(delay*1000);
            ExcelAsyncUtil.QueueAsMacro(() =>
                _excelWinFormsUtil.MessageBox(
                    "Message box called from background thread",
                    "Long Running Thread",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            );
        }

        public void Dispose()
        {
        }
    }
}