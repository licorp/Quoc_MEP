using System;
using System.Threading.Tasks;
using System.Windows.Threading;
using Autodesk.Revit.UI;

namespace Quoc_MEP.Lib
{
    /// <summary>
    /// Mini async helper cho Revit API - inspired by Revit.Async
    /// Cho phép chạy Revit API code từ bất kỳ context nào (modeless window, background thread)
    /// </summary>
    public static class RevitAsyncHelper
    {
        private static UIApplication _uiApp;
        private static ExternalEvent _externalEvent;
        private static GenericExternalEventHandler _handler;

        /// <summary>
        /// Initialize trong Revit API context (IExternalCommand.Execute hoặc IExternalApplication.OnStartup)
        /// </summary>
        public static void Initialize(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _handler = new GenericExternalEventHandler();
            _externalEvent = ExternalEvent.Create(_handler);
        }

        /// <summary>
        /// Run Revit API code async - SAFE từ bất kỳ thread nào
        /// </summary>
        public static Task RunAsync(Action<UIApplication> action)
        {
            if (_uiApp == null || _externalEvent == null)
                throw new InvalidOperationException("RevitAsyncHelper not initialized. Call Initialize() first!");

            var tcs = new TaskCompletionSource<bool>();
            
            _handler.SetAction(() =>
            {
                try
                {
                    action(_uiApp);
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            _externalEvent.Raise();
            return tcs.Task;
        }

        /// <summary>
        /// Run Revit API code async với return value
        /// </summary>
        public static Task<T> RunAsync<T>(Func<UIApplication, T> func)
        {
            if (_uiApp == null || _externalEvent == null)
                throw new InvalidOperationException("RevitAsyncHelper not initialized. Call Initialize() first!");

            var tcs = new TaskCompletionSource<T>();

            _handler.SetAction(() =>
            {
                try
                {
                    var result = func(_uiApp);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            _externalEvent.Raise();
            return tcs.Task;
        }

        /// <summary>
        /// Internal handler for ExternalEvent
        /// </summary>
        private class GenericExternalEventHandler : IExternalEventHandler
        {
            private Action _action;

            public void SetAction(Action action)
            {
                _action = action;
            }

            public void Execute(UIApplication app)
            {
                _action?.Invoke();
                _action = null; // Clear after execution
            }

            public string GetName()
            {
                return "RevitAsyncHelper";
            }
        }
    }
}
