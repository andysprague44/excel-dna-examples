using System;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration.CustomUI;
using NetOffice.ExcelApi;

namespace CustomImage
{
    class ExcelController : IDisposable
    {
        private readonly IRibbonUI _modelingRibbon;
        protected readonly Application _excel;

        public ExcelController(Application excel, IRibbonUI modelingRibbon)
        {
            _modelingRibbon = modelingRibbon;
            _excel = excel;
        }

        public void PressMe(string id)
        {
            var activeSheet = _excel.ActiveSheet as Worksheet;
            activeSheet.Range("A1").Value = $"{id}, {id.ToLower()}, {id.ToLower()}";
        }

        public void Dispose()
        {
        }
    }
}
