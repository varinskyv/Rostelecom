using System;
using System.Collections.Generic;
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

        private enum DataType
        {
            Intertet,
            SMS,
            MMS,
            Phone
        }

        private static string Stop_Sheet_Tag = "Sheet2";

        private static string Period_Tag = "за период с";

        private static string Subscriber_Start_Tag = "абонент Тел";
        private static string Subscriber_End_Tag = "Всего по абоненту";

        private static string Internet_Tag = "Передача данных";
        private static string[] Internet_Headder = { "Тип вызова", "Интернет адрес", "Время начала", "Объем(байт)", "Стоимость (руб)" };
        private static string[] Internet_Fields = { "0", "2", "1", "4", "3" }; 
        private static string[] Internet_GPRS_Tags = { "GPRS-Интернет" };
        private static string[] Internet_WAP_Tags = { "WAP-Интернет" };

        private static string SMS_Tag = "SMS";
        private static string[] SMS_Headder = { "Дата", "Время", "Услуга", "Телефон", "Количество", "Стоимость (руб)" };
        private static string[] SMS_Fields = { "2", "0 1", "3", "5", "4" };
        private static string[] SMS_Incoming_Tags = { "Входящее СМС" };
        private static string[] SMS_Outgoing_Tags = { "Исх.СМС" };

        private static string MMS_Tag = "MMS";
        private static string[] MMS_Headder = { "Дата", "Время", "Услуга", "Телефон", "Количество", "Стоимость (руб)" };
        private static string[] MMS_Fields = { "2", "0 1", "3", "5", "4" };
        private static string[] MMS_Incoming_Tags = { "Входящее ММС" };
        private static string[] MMS_Outgoing_Tags = { "Исх.ММС" };

        private static string Phone_Tag = " Телефония";
        private static string[] Phone_Headder = { "Дата", "Время", "Услуга", "Телефон", "Длительность (сек)", "Стоимость (руб)" };
        private static string[] Phone_Fields = { "2", "0 1", "3", "5", "4" };
        private static string[] Phone_Incoming_Tags = { "Вход" };
        private static string[] Phone_Outgoing_Tags = { "Исх.", "Экстренные оперативные службы" };

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

            try
            {
                List<List<object>> data = (List<List<object>>)obj;

                result.StartPeriodDate = GetStartPeriodDate(data);

                lBase.Transaction();

                int i = 0;
                while(i < data.Count)
                {
                    if (((string)data[i][0]).Equals(Stop_Sheet_Tag))
                        break;

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
                            while (i < data.Count)
                            {
                                if (data[i].Find(cell => Convert.ToString(cell).Contains(Subscriber_End_Tag)) != null)
                                    break;

                                if (data[i].Find(cell => Convert.ToString(cell).Contains(Internet_Tag)) != null)
                                {
                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (FindHeadder(data[i], Internet_Headder) == true)
                                            break;

                                        i++;
                                    }

                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (data[i].Find(cell => Convert.ToString(cell).Contains(End_Unit_Tag)) != null)
                                            break;

                                        Connection connection = GetConnection(data[i], DataType.Intertet);

                                        if (connection != null)
                                            lBase.AddConnection(subscriber.Id, connection);

                                        i++;
                                    }
                                }

                                if (data[i].Find(cell => Convert.ToString(cell).Contains(SMS_Tag)) != null)
                                {
                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (FindHeadder(data[i], SMS_Headder) == true)
                                            break;

                                        i++;
                                    }

                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (data[i].Find(cell => Convert.ToString(cell).Contains(End_Unit_Tag)) != null)
                                            break;

                                        Connection connection = GetConnection(data[i], DataType.SMS);

                                        if (connection != null)
                                            lBase.AddConnection(subscriber.Id, connection);

                                        i++;
                                    }
                                }

                                if (data[i].Find(cell => Convert.ToString(cell).Contains(MMS_Tag)) != null)
                                {
                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (FindHeadder(data[i], MMS_Headder) == true)
                                            break;

                                        i++;
                                    }

                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (data[i].Find(cell => Convert.ToString(cell).Contains(End_Unit_Tag)) != null)
                                            break;

                                        Connection connection = GetConnection(data[i], DataType.MMS);

                                        if (connection != null)
                                            lBase.AddConnection(subscriber.Id, connection);

                                        i++;
                                    }
                                }

                                if (data[i].Find(cell => Convert.ToString(cell).Contains(Phone_Tag)) != null)
                                {
                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (FindHeadder(data[i], Phone_Headder) == true)
                                            break;

                                        i++;
                                    }

                                    i++;
                                    while (i < data.Count)
                                    {
                                        if (data[i].Find(cell => Convert.ToString(cell).Contains(End_Unit_Tag)) != null)
                                            break;

                                        Connection connection = GetConnection(data[i], DataType.Phone);

                                        if (connection != null)
                                            lBase.AddConnection(subscriber.Id, connection);

                                        i++;
                                    }
                                }
                                
                                i++;
                            }
                        }
                    }

                    i++;
                }

                result.Result = true;
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }

            lBase.Commit();

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

        private string GetString(List<object> row, string tag)
        {
            return (string)row.Find(item => Convert.ToString(item).Contains(tag));
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

        private bool FindHeadder(List<object> row, string[] headder)
        {
            bool result = false;

            if (headder.Length > row.Count)
                return result;

            int i = 0;
            while (i < headder.Length)
            {
                if (((string)row[i]).Contains(headder[i]))
                    result = true;
                else
                    result = false;

                i++;
            }

            return result;
        }

        private Connection GetConnection(List<object> row, DataType dataType)
        {
            Connection result = null;

            string[] fields = null;
            switch (dataType)
            {
                case DataType.Intertet:
                    fields = Internet_Fields;
                    break;

                case DataType.SMS:
                    fields = SMS_Fields;
                    break;

                case DataType.MMS:
                    fields = MMS_Fields;
                    break;

                case DataType.Phone:
                    fields = Phone_Fields;
                    break;
            }

            if (row.Count < fields.Length)
                return result;

            try
            {
                result = new Connection();

                string cell = GetCell(row, fields[0]);

                switch (dataType)
                {
                    case DataType.Intertet:
                        {
                            if (SetConnectionType(cell, Internet_GPRS_Tags))
                            {
                                result.Type = ConnectionType.GPRS;
                                break;
                            }

                            if (SetConnectionType(cell, Internet_WAP_Tags))
                            {
                                result.Type = ConnectionType.WAP;
                                break;
                            }

                            result.Type = ConnectionType.OtherInternet;
                        }
                        break;

                    case DataType.SMS:
                        {
                            if (SetConnectionType(cell, SMS_Incoming_Tags))
                            {
                                result.Type = ConnectionType.IncomingSMS;
                                break;
                            }

                            if (SetConnectionType(cell, SMS_Outgoing_Tags))
                            {
                                result.Type = ConnectionType.OutgoingSMS;
                                break;
                            }
                        }
                        break;

                    case DataType.MMS:
                        {
                            if (SetConnectionType(cell, MMS_Incoming_Tags))
                            {
                                result.Type = ConnectionType.IncomingMMS;
                                break;
                            }

                            if (SetConnectionType(cell, MMS_Outgoing_Tags))
                            {
                                result.Type = ConnectionType.OutgoingMMS;
                                break;
                            }
                        }
                        break;

                    case DataType.Phone:
                        {
                            if (SetConnectionType(cell, Phone_Incoming_Tags))
                            {
                                result.Type = ConnectionType.IncomingCall;
                                break;
                            }

                            if (SetConnectionType(cell, Phone_Outgoing_Tags))
                            {
                                result.Type = ConnectionType.OutgoingCall;
                                break;
                            }
                        }
                        break;
                }

                cell = GetCell(row, fields[1]);
                result.Date = DateTime.Parse(cell);

                cell = GetCell(row, fields[2]);
                result.IOTarget = cell.Replace(" ", "");

                cell = GetCell(row, fields[3]);
                result.Cost = Decimal.Parse(cell, NumberStyles.Currency, CultureInfo.InvariantCulture);

                cell = GetCell(row, fields[4]);
                result.Value = Convert.ToInt32(cell);
            }
            catch (Exception e)
            {
                result = null;

                Console.WriteLine(e);
            }

            return result;
        }

        private bool SetConnectionType(string cell, string[] tags)
        {
            int i = 0;
            while (i < tags.Length)
            {
                if (cell.Contains(tags[i]))
                    return true;

                i++;
            }

            return false;
        }

        private string GetCell(List<object> row, string field)
        {
            int index = 0;
            if (int.TryParse(field, out index))
            {
                return Convert.ToString(row[index]);
            }
            else
            {
                int[] indices = null;
                string[] splits = null;
                ParseField(field, out indices, out splits);

                string result = "";
                int i = 0;
                if (indices.Length >= splits.Length)
                {
                    while (i < indices.Length)
                    {
                        result += Convert.ToString(row[indices[i]]);

                        if (i < splits.Length)
                            result += splits[i];

                        i++;
                    }
                }
                else
                {
                    while (i < splits.Length)
                    {
                        result += splits[i];

                        if (i < indices.Length)
                            result += Convert.ToString(row[indices[i]]);

                        i++;
                    }
                }

                return result;
            }
        }

        private void ParseField(string field, out int[] indices, out string[] splits)
        {
            int[] i = new int[0];
            string[] s = new string[0];

            bool isNotIndexFirst;
            if (Char.IsDigit(field[0]))
                isNotIndexFirst = false;
            else
                isNotIndexFirst = true;

            int n = 0;
            string index = "";
            string split = "";
            while(n < field.Length)
            {
                if (Char.IsDigit(field[n]))
                {
                    index += field[n];

                    if (!String.IsNullOrEmpty(split))
                    {
                        Array.Resize(ref s, s.Length + 1);
                        s[s.Length - 1] = split;

                        split = "";
                    }
                }
                else
                {
                    split += field[n];

                    if (!String.IsNullOrEmpty(index))
                    {
                        Array.Resize(ref i, i.Length + 1);
                        i[i.Length - 1] = Convert.ToInt32(index);

                        index = "";
                    }
                }

                n++;
            }

            if (!String.IsNullOrEmpty(index))
            {
                Array.Resize(ref i, i.Length + 1);
                i[i.Length - 1] = Convert.ToInt32(index);
            }

            if (!String.IsNullOrEmpty(split))
            {
                Array.Resize(ref s, s.Length + 1);
                s[s.Length - 1] = split;
            }

            if (isNotIndexFirst)
            {
                Array.Resize(ref s, s.Length + 1);
                s[s.Length - 1] = "";
            }

            indices = i;
            splits = s;
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
