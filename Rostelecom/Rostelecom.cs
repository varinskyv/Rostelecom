using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using DetailedOperatorServicesCore;

namespace Rostelecom
{
    public class CallBackResult
    {
        public bool Result;
        public DateTime Period;
    }

    public class Rostelecom : IDisposable
    {
        public static string Name = "Rostelecom";

        public delegate void CallBackHandler(CallBackResult result);
        private CallBackHandler resultListener;

        private Thread classThread;

        //Асинхронный метод, передающий результат в вызывающий класс
        private CallBackResult CallBack(CallBackResult result)
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
            CallBackResult result = new CallBackResult();

            try
            {
                List<List<object>> data = (List<List<object>>)obj;


            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }

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
