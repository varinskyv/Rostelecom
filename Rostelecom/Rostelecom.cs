using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using DetailedOperatorServicesCore;
using System.Globalization;

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
        private static string[] Internet_Headder = { "Тип вызова", "Интернет адрес", "Время начала", "Объем(байт)", "Стоимость (руб)" };
        private static string Internet_GPRS_Tag = "GPRS-Интернет";
        private static string Internet_WAP_Tag = "WAP-Интернет";
        private static string SMS_Tag = "Прием SMS";
        private static string Phone_Tag = " Телефония исходящая";
        private static string End_Unit_Tag = "Итого";

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
                        Subscriber subscriber = new Subscriber()
                        {
                            Number = GetSubscriberNumber(data[i])
                        };

                        if (!String.IsNullOrEmpty(subscriber.Number))
                        {
                            subscriber.Id = lBase.GetSubscriberIdByNumber(subscriber.Number);
                            if (subscriber.Id == 0)
                                subscriber.Id = lBase.AddSubscriber(subscriber);

                            if (subscriber.Id == 0)
                                continue;

                            i++;
                            while (data[i].Find(cell => Convert.ToString(cell).Contains(Subscriber_End_Tag)) == null && i < data.Count)
                            {
                                if (data[i].Find(cell => Convert.ToString(cell).Contains(Internet_Tag)) != null)
                                {
                                    i++;
                                    while (FindHeadder(data[i], Internet_Headder) != true && i < data.Count)
                                        i++;

                                    i++;
                                    while (data[i].Find(cell => Convert.ToString(cell).Contains(End_Unit_Tag)) == null && i < data.Count)
                                    {
                                        Connection connection = GetInternetConnection(data[i], Internet_Headder.Length);

                                        if (connection != null)
                                            lBase.AddConnection(subscriber.Id, connection);

                                        i++;
                                    }
                                }

                                //

                                i++;
                            }
                        }
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

        private bool FindHeadder(List<object> list, string[] headder)
        {
            bool result = false;

            if (headder.Length > list.Count)
                return result;

            int i = 0;
            while (i < headder.Length)
            {
                if (headder[i].Equals((string)list[i]))
                    result = true;
                else
                    result = false;

                i++;
            }

            return result;
        }

        private Connection GetInternetConnection(List<object> list, int length)
        {
            Connection result = null;

            if (list.Count < length)
                return result;

            try
            {
                result = new Connection();

                if (Convert.ToString(list[0]).Replace(" ","").Equals(Internet_GPRS_Tag))
                    result.Type = ConnectionType.GPRS;
                else
                    result.Type = ConnectionType.WAP;

                result.IOTarget = Convert.ToString(list[1]).Replace(" ", "");
                result.Date = DateTime.Parse(Convert.ToString(list[2]));
                result.Value = Convert.ToInt32(list[3]);
                result.Cost = Decimal.Parse(Convert.ToString(list[4]), NumberStyles.Currency, CultureInfo.InvariantCulture); ;
            }
            catch(Exception e)
            {
                result = null;

                Console.WriteLine(e);
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
