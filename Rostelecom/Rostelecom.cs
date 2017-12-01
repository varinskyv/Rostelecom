using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using DetailedOperatorServicesCore;

namespace Rostelecom
{
    public class Rostelecom : IDisposable
    {
        public static string Name = "Rostelecom";

        public delegate void CallBackHandler(bool result);
        private CallBackHandler resultListener;

        private Thread classThread;

        //Асинхронный метод, передающий результат в вызывающий класс
        private bool CallBack(bool result)
        {
            return result;
        }

        public void Parse(List<List<object>> data, CallBackHandler resultListener)
        {
            this.resultListener = resultListener;

            classThread = new Thread(new ParameterizedThreadStart(_Parse));
            classThread.Start(data);
        }

        private void _Parse(object obj)
        {
            bool result = false;

            //TODO:

            resultListener?.Invoke(result);
        }

        public void Dispose()
        {
            classThread.Abort();
            classThread = null;

            resultListener = null;

            GC.GetTotalMemory(true);
        }
    }
}
