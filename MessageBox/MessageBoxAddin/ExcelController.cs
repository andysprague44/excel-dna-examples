using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using MessageBoxAddin.Forms;
using static MessageBoxAddin.Extensions.ExcelDnaExtensions;

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
        public async Task RunBackgroundThread(int delay)
        {
            Thread.Sleep(delay*1000);
            
            //get user input as part of a background thread
            var dialogResult = await _excel.QueueAsMacroAsync(xl =>
                _excelWinFormsUtil.MessageBox(
                    "Message box called from background thread",
                    "Long Running Thread",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information)
            );

            //do stuff depending on dialog result in the background

            //finally, call back to excel to write some result
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                _excel.Range("A1").Value = dialogResult.ToString();
            });
        }

        public void Dispose()
        {
        }
    }
}