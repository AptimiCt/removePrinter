using System;
using System.Collections.Generic;
using System.Management;
using System.Reflection;

namespace remove_printer
{
    class ListOfWin32Printer
    {
        public class Win32_Printer
        {
            public string DriverName;
            public string Location;
            public string Name;
            public bool Network;
            public string PortName;
            public string ServerName;
            public bool Shared;
            //public PrinterStatus Status;
            public bool WorkOffline;

            public static List<Win32_Printer> GetList()
            {
                string query = "Select * From Win32_Printer";

                ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);

                ManagementObjectCollection results = searcher.Get();

                List<Win32_Printer> list = new List<Win32_Printer>(results.Count);

                foreach (ManagementObject obj in results)
                {
                    Win32_Printer entry = new Win32_Printer();

                    foreach (FieldInfo field in typeof(Win32_Printer).GetFields())
                        field.SetValue(entry, ConvertValue(obj[field.Name], field.FieldType));

                    list.Add(entry);
                }
                return list;
            }

            private static object ConvertValue(object value, Type type)
            {
                if (value != null)
                {
                    if (type == typeof(DateTime))
                    {
                        string time = value.ToString();
                        time = time.Substring(0, time.IndexOf("."));
                        return DateTime.ParseExact(time, "yyyyMMddHHmmss", null);
                    }
                    else if (type == typeof(long))
                        return Convert.ToInt64(value);
                    else if (type == typeof(int))
                        return Convert.ToInt32(value);
                    else if (type == typeof(short))
                        return Convert.ToInt16(value);
                    else if (type == typeof(string))
                        return value.ToString();
                    //else if (type == typeof(PrinterStatus))
                    //    return (PrinterStatus)Enum.Parse(typeof(PrinterStatus), value.ToString());
                }
                return null;
            }
        }
    }
}

