using System;
using System.Threading.Tasks;
using Application = NetOffice.ExcelApi.Application;
using ExcelDna.Integration;

namespace MessageBoxAddin.Extensions
{
    public static class ExcelDnaExtensions
    {
        /// <summary>
        /// Run a function using ExcelAsyncUtil.QueueAsMacro and allow waiting for the result.
        /// Waits until excel resources are free, runs the func, then waits for the func to complete.
        /// </summary>
        /// <example>
        /// var dialogResult = await excel.QueueAsMacroAsync(e =>
        ///     _excelWinFormsUtil.MessageBox("Message", "Caption", MessageBoxButtons.YesNo, MessageBoxIcon.Question) );
        /// </example>
        public static async Task<T> QueueAsMacroAsync<T>(this Application excel, Func<Application, T> func)
        {
            try
            {
                var tcs = new TaskCompletionSource<T>();
                ExcelAsyncUtil.QueueAsMacro((x) =>
                {
                    var tcsState = (TaskCompletionSource<T>)((object[])x)[0];
                    var f = (Func<Application, T>)((object[])x)[1];
                    var xl = (Application)((object[])x)[2];
                    try
                    {
                        var result = f(xl);
                        tcsState.SetResult(result);
                    }
                    catch (Exception ex)
                    {
                        tcsState.SetException(ex);
                    }
                }, new object[] { tcs, func, excel });
                var t = await tcs.Task;
                return t;
            }
            catch (AggregateException aex)
            {
                var flattened = aex.Flatten();
                throw new Exception(flattened.Message, flattened);
            }
        }
    }
}

