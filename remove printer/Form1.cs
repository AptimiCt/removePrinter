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
        //Расположение всех принтеров пользователя в реестре
        //private readonly string volatileEnvironment = @"Volatile Environment";
        static string logonUI = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI";
       
        //static string lastLoggedOnSAMUser = "LastLoggedOnSAMUser";
        //string logonUser = LogonUSER(logonUI, "LastLoggedOnSAMUser");
        static string lastLoggedOnUserSID = LogonUSER(logonUI, "LastLoggedOnUserSID");//Registry.LocalMachine.OpenSubKey(logonUI).GetValue("LastLoggedOnUserSID").ToString();
        //string logonUSER = Registry.LocalMachine.OpenSubKey(logonUI).GetValue("LastLoggedOnSAMUser").ToString();
        private readonly string dfpath = $@"{lastLoggedOnUserSID}\Software\Microsoft\Windows NT\CurrentVersion\Windows";
        private readonly string dvpath = $@"{lastLoggedOnUserSID}\Software\Microsoft\Windows NT\CurrentVersion\Devices\";
        
        
        //Определение каталога запуска приложения
        private static string pathcsv = Application.StartupPath;

        //Путь до файла с исключениями
        private string excludeCSV = $@"{pathcsv}\exclude.csv";
        //Пути до файлов с логами
        private string err_def = $@"{pathcsv}\err_def.txt";
        private string err_del_pr = $@"{pathcsv}\err_del_pr.txt";
        private string err_ex = $@"{pathcsv}\err_ex.txt";
        private string def_prn = $@"{pathcsv}\default.txt";
        private string del = $@"{pathcsv}\del.txt";

        string ps = Environment.UserName;
        private string defaulPrinter = "";        
        private ManagementScope managementScope = null;
        private ManagementObjectCollection managementObjectCollection = null;
        private string[] listPrn;

        public Form1()
        {
            InitializeComponent();
        }
        static string LogonUSER(string logonUI, string value)
        {
            string str = "";
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(logonUI)) {
                str = key.GetValue(value).ToString();
            } ;
            return str.StartsWith(@".\") ? str.Substring(2) : str; 
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

                using (RegistryKey regKey = Registry.Users.OpenSubKey(dfpath))
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
                listPrn = Registry.Users.OpenSubKey(dvpath).GetValueNames();
                foreach (string printer in listPrn)
                {
                    string portOfPrinter = GetPrinterPort(printer);
                    byte count = 0;
                    foreach (string port in excludePorts)
                    {
                        if (portOfPrinter.ToLower().Contains(port.ToLower()))
                        {
                            count++;
                            break;
                        }
                    }
                    if (count == 0)
                    {
                        File.AppendAllText(del, printer);
                        File.AppendAllText(del, "\n");
                    RemovePrinter(printer);
                    }
                }
                //Установка принтера поумолчанию
                byte countDef = 0;
                string[] listPrns = Registry.Users.OpenSubKey(dvpath).GetValueNames();
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

            label1.Text = ps;
            //label2.Text = Registry.CurrentUser.OpenSubKey(volatileEnvironment).GetValue("USERNAME").ToString();   //ReadReg(user_key, "USERNAME");
            label2.Text = LogonUSER(logonUI, "LastLoggedOnSAMUser");
            label3.Text = lastLoggedOnUserSID;
            //MainTask();
        }
    }
}
