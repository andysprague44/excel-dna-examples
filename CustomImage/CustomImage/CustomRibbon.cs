using System;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using System.Drawing;

namespace CustomImage
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
                switch (control.Id)
                {
                    case "RugbyImageButton":
                        controller.PressMe("Rugby");
                        break;
                    case "FootballImageButton":
                        controller.PressMe("Football");
                        break;
                    case "GolfImageButton":
                        controller.PressMe("Golf");
                        break;
                    default:
                        controller.PressMe("Sports");
                        break;
                }
            }
        }
        public Bitmap GetImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "RugbyImageButton": return new Bitmap(Properties.Resources.rugby);
                case "FootballImageButton": return new Bitmap(Properties.Resources.football);
                case "GolfImageButton": return new Bitmap(Properties.Resources.golf);
                default: return null;
            }
        }
    }
}