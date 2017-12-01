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

        private LocalBase lBase;

        private static string Period_Tag = "за период с";
        private static string Subscriber_Start_Tag = "абонент Тел";
        private static string Subscriber_End_Tag = "Всего по абоненту";
        private static string Internet_Tag = "Передача данных";
        private static string End_Unit_Tag = "Передача данных";

        public Rostelecom()
        {
            lBase = LocalBase.getInstance();
        }

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

            //try
            {
                List<List<object>> data = (List<List<object>>)obj;

                result.StartPeriodDate = GetStartPeriodDate(data);

                int i = 0;
                while(!((string)data[i][0]).Equals("Sheet2") && i < data.Count)
                {
                    if (data[i].Find(cell => Convert.ToString(cell).Contains(Subscriber_Start_Tag)) != null)
                    {
                        Subscriber subscriber = new Subscriber();
                        subscriber.Number = GetSubscriberNumber(data[i]);

                        int Id = lBase.AddSubscriber(subscriber);
                    }

                    i++;
                }

                result.Result = true;
            }
            //catch(Exception e)
            //{
            //    Console.WriteLine(e);
            //}

            resultListener?.Invoke(result);
        }

        private DateTime GetStartPeriodDate(List<List<object>> data)
        {
            List<object> periodRow = data.Find(item => item.Find(row => Convert.ToString(row).Contains(Period_Tag)) != null);

            string periodStr = GetString(periodRow, Period_Tag);

            string startDate = "";
            int i = periodStr.IndexOf(Period_Tag) + Period_Tag.Length + 1;
            while(!periodStr[i].Equals(' '))
            {
                startDate += periodStr[i];

                i++;
            }

            return DateTime.Parse(startDate);
        }

        private string GetString(List<object> list, string tag)
        {
            return (string)list.Find(item => Convert.ToString(item).Contains(tag));
        }

        private string GetSubscriberNumber(List<object> row)
        {
            string result = "";

            string str = GetString(row, Subscriber_Start_Tag);
            int i = str.IndexOf(Subscriber_Start_Tag) + Subscriber_Start_Tag.Length + 1;
            while (i < str.Length)
            {
                if (Char.IsDigit(str[i]))
                {
                    result += str[i];
                }

                i++;
            }

            return result; 
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
