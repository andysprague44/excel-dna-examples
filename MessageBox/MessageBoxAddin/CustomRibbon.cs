﻿using System;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using MessageBoxAddin.Forms;

namespace MessageBoxAddin
{
    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        private Application _excel;
        private IExcelWinFormsUtil _excelWinFormsUtil;
        private IRibbonUI _thisRibbon;

        public override string GetCustomUI(string ribbonId)
        {
            _excel = new Application(null, ExcelDna.Integration.ExcelDnaUtil.Application);
            _excelWinFormsUtil = new ExcelWinFormsUtil();

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
            using (var controller = new ExcelController(_excel, _thisRibbon, _excelWinFormsUtil))
            {
                controller.PressMe();
            }
        }

        private static int _delaySeconds = 5;

        public void OnDelayEditBoxChanged(IRibbonControl control, string text)
        {
            int delay;
            var success = int.TryParse(text, out delay);
            if (success)
                _delaySeconds = delay;
            else
                _delaySeconds = 5;
        }

        public void OnPressMeBackgroundThread(IRibbonControl control)
        {
            using (var controller = new ExcelController(_excel, _thisRibbon, _excelWinFormsUtil))
            {
                controller.OnPressMeBackgroundThread(_delaySeconds);
            }
        }
    }
}