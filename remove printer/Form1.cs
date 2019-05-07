using Microsoft.Win32;
using System;
using System.IO;
using System.Management;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;

namespace remove_printer
{
    public partial class Form1 : Form
    {
        //Расположение дефолтного принтера пользователя в реестре
        private readonly string dfpath = @"Software\Microsoft\Windows NT\CurrentVersion\Windows\";
        //Расположение всех принтеров пользователя в реестре
        private readonly string dvpath = @"Software\Microsoft\Windows NT\CurrentVersion\Devices\";

        //string user_key = @"Volatile Environment\";

        //Определение каталога запуска приложения
        private static string pathcsv = Application.StartupPath;

        //Путь до файла с исключениями
        private string excludeCSV = $@"{pathcsv}\exclude.csv";
        //Пути до файлов с логами
        private string err_def = $@"{pathcsv}\err_def.txt";
        private string err_del_pr = $@"{pathcsv}\err_del_pr.txt";
        private string err_ex = $@"{pathcsv}\err_ex.txt";
        private string def_prn = $@"{pathcsv}\default.txt";

        //string ps = Environment.UserName;
        private string defaulPrinter = "";        
        private ManagementScope managementScope = null;
        private ManagementObjectCollection managementObjectCollection = null;
        private string[] listPrn;

        public Form1()
        {
            InitializeComponent();
        }
        

        private void Logger(Exception ex, string path)
        {
            File.AppendAllText(path, ex.StackTrace, Encoding.Unicode);
        }

        //string ReadReg (string path, string param)
        //{
        //    RegistryKey regKey = null;
           
        //        using (regKey = Registry.CurrentUser.OpenSubKey(path))
        //        {

        //            return regKey.GetValue(param).ToString();
        //        }
             

           
        //}
        //Принтер по-умолчанию
        string GetDefaultPrinter()
        {
            try
            {
                string str = "";
                using (RegistryKey regKey = Registry.CurrentUser.OpenSubKey(dfpath))
                {
                    str = regKey.GetValue("Device").ToString();
                }
                //ReadReg(dfpath, "Device");
                str = str.Substring(0, str.IndexOf(','));
                int lastIndOf = str.LastIndexOf("\\"); 
                return lastIndOf == -1 ? str : str.Substring(lastIndOf + 1); 
            }
            catch (Exception ex)
            {
                Logger(ex, err_def);
                return "";
            }
        }
        //Удаление принтера
        public bool RemovePrinter(string name)
        {
            string err = "";
            try
            {
                //Использование ManagementScope class для подключения к арм
                ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath);
                scope.Connect();
                //Запрос Win32_Printer
                SelectQuery query = new SelectQuery("select * from Win32_Printer");
                ManagementObjectSearcher search = new ManagementObjectSearcher(scope, query);
                //Получение всех принтеров на арм
                ManagementObjectCollection printers = search.Get();
                foreach (ManagementObject printer in printers)
                {
                    //получение имени конкретного принтера            
                    //проверить совпадение, если найдено, удалить или продолжить                    
                    if (printer["Name"].ToString().ToLower().Equals(name.ToLower()))
                    {
                        err = printer["Name"].ToString();
                        printer.Delete();
                        break;
                    }
                    else
                        continue;
                }
                return true;
            }
            catch (Exception ex)
            {
                Logger(ex, err_del_pr);
                File.AppendAllText(err_del_pr, $"Не вышло удалить {err}", Encoding.Unicode);
                return false;
            }

        }
        //Получение коллекции объектов
        private ManagementObjectCollection GetManagementObject(string printerName)
        {
            managementScope = new ManagementScope(ManagementPath.DefaultPath);
            managementScope.Connect();
            SelectQuery selectQuery = new SelectQuery();
            selectQuery.QueryString = @"SELECT * FROM Win32_Printer 
	            WHERE Name = '" + printerName.Replace("\\", "\\\\") + "'";
            ManagementObjectSearcher managementObjectSearcher =
               new ManagementObjectSearcher(managementScope, @selectQuery);
            return managementObjectSearcher.Get();
        }
        public string GetPrinterPort(string printerName)
        {
            managementObjectCollection = GetManagementObject(printerName);
            string portName = "";
            foreach (ManagementObject managementItem in managementObjectCollection)
            {              
                portName = managementItem["PortName"].ToString();
            }
            return portName;
        }
        public void SetDefaultPrinter(string printerName)
        {
            managementObjectCollection = GetManagementObject(printerName);
            if (managementObjectCollection.Count != 0)
            {
                foreach (ManagementObject managementItem in managementObjectCollection)
                {
                    managementItem.InvokeMethod("SetDefaultPrinter", new object[] { printerName });
                    return;
                }
            }
        }

        void MainTask()
        {
                defaulPrinter = GetDefaultPrinter();
                File.AppendAllText(def_prn, defaulPrinter, Encoding.Unicode);

                string[] excludePorts = null;
                try
                {
                    excludePorts = File.ReadAllLines(excludeCSV);
                }
                catch(Exception ex)
                {
                    Logger(ex, err_ex);
                    File.AppendAllText(err_ex, "Нет файла exclude.csv", Encoding.Unicode);
                    Close();
                }
                listPrn = Registry.CurrentUser.OpenSubKey(dvpath).GetValueNames();
                foreach (string printer in listPrn)
                {
                    string portOfPrinter = GetPrinterPort(printer);
                    byte count = 0;
                    foreach (string port in excludePorts)
                    {
                        if (portOfPrinter.ToLower().Contains(port.ToLower()))
                        {
                            count++;
                        }
                    }
                    if (count == 0)
                    {
                        RemovePrinter(printer);
                    }
                }
                //Установка принтера поумолчанию
                byte countDef = 0;
                string[] listPrns = Registry.CurrentUser.OpenSubKey(dvpath).GetValueNames();
                foreach (string listPrn in listPrns)
                {
                    if (defaulPrinter.ToLower().Contains(listPrn.ToLower()))
                    {
                        defaulPrinter = listPrn;
                        break;
                    }
                    else
                        countDef++;
                }
                if (countDef == listPrns.Length)
                    SetDefaultPrinter("SafeQ");
                else
                    SetDefaultPrinter(defaulPrinter);

                Close();            
        }
        private void ReadFileAndRemove_Click(object sender, EventArgs e)
        {
            MainTask();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //label1.Text = ps;
            //label2.Text = ReadReg(user_key, "USERNAME");
            MainTask();
        }
    }
}
