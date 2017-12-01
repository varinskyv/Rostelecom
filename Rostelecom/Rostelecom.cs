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
        public DateTime StartPeriodDate;
    }

    public class Rostelecom : IDisposable
    {
        public static string Name = "Rostelecom";

        public delegate void CallBackHandler(CallBackResult result);
        private CallBackHandler resultListener;

        private Thread classThread;

        private static string Period_Tag = "за период с";
        private static string Subscriber_Start_Tag = "абонент Тел";
        private static string Subscriber_End_Tag = "Всего по абоненту";
        private static string Internet_Tag = "Передача данных";
        private static string End_Unit_Tag = "Передача данных";

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

                result.StartPeriodDate = GetStartPeriodDate(data);
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }

            resultListener?.Invoke(result);
        }

        private DateTime GetStartPeriodDate(List<List<object>> data)
        {
            data.Find(item => item.Contains(cel => ((string)cel).ToUpper.Equals(Period_Tag.ToUpper())));

            return new DateTime(); //TODO:
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
