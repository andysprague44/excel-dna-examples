using System;
using System.Drawing;
using System.IO;
using System.Resources;
using System.Reflection;
using System.Runtime.InteropServices;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using NetOffice.VBIDEApi;

namespace MyExcelAddin
{
    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        private Application _excel;
        private IRibbonUI _thisRibbon;

        public override string GetCustomUI(string ribbonId)
        {
            _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            string ribbonXml = GetCustomRibbonXML();
            return ribbonXml;
        }

        private string GetCustomRibbonXML()
        {
            string ribbonXml;
            var thisAssembly = typeof(CustomRibbon).Assembly;
            var resourceName = typeof(CustomRibbon).Namespace + ".CustomRibbon.xml";

            using (Stream stream = thisAssembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                ribbonXml = reader.ReadToEnd();
            }

            if (ribbonXml == null)
            {
                throw new MissingManifestResourceException(resourceName);
            }
            return ribbonXml;
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));
            }

            _thisRibbon = ribbon;

            _excel.WorkbookActivateEvent += OnInvalidateRibbon;
            _excel.WorkbookDeactivateEvent += OnInvalidateRibbon;
            _excel.SheetActivateEvent += OnInvalidateRibbon;
            _excel.SheetDeactivateEvent += OnInvalidateRibbon;

            if (_excel.ActiveWorkbook == null)
            {
                _excel.Workbooks.Add();
            }
        }

        private void OnInvalidateRibbon(object obj)
        {
            _thisRibbon.Invalidate();
        }

        public void OnPressMe(IRibbonControl control)
        {
            using (var controller = new ExcelController(_excel, _thisRibbon))
            {
                controller.PressMe();
            }
        }

//        public Bitmap GetImage(IRibbonControl control)
//        {
//            switch (control.Id)
//            {
//                case "GetCrackerTemplateButton":
//                    return new Bitmap(Properties.Resources.cracker);
//                case "GetGenericTemplateButton":
//                    return new Bitmap(Properties.Resources.generic);
//                case "GetAllAnalysisButton":
//                    return new Bitmap(Properties.Resources.download);
//                case "GetPreviousAnalysisButton":
//                    return new Bitmap(Properties.Resources.copy);
//                case "RunRollupButton":
//                    return new Bitmap(Properties.Resources.rollup);
//                default:
//                    return null;
//            }
//        }
//        public void OnGetRugbyImage
//        //<button id="CustomImage" label="RugbyImage"  getImage="OnGetRugbyImage" size="large" onAction="OnPressMe" />
    }
}