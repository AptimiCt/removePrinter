using Microsoft.Win32;
using System;
using System.IO;
using System.Management;
using System.Text;
using System.Windows.Forms;

namespace remove_printer
{
    public partial class Form1 : Form
    {
        //Расположение дефолтного принтера пользователя в реестре
        readonly string dfpath = @"Software\Microsoft\Windows NT\CurrentVersion\Windows\";
        //Расположение всех принтеров пользователя в реестре
        readonly string dvpath = @"Software\Microsoft\Windows NT\CurrentVersion\Devices\";
        //Путь до файла с исключениями
        string excludeCSV = @"C:\test\exclude.csv";
        string defaulPrinter = "";        
        private ManagementScope managementScope = null;
        private ManagementObjectCollection managementObjectCollection = null;  
        string[] listPrn;
        public Form1()
        {
            InitializeComponent();
        }        
        //Принтер по-умолчанию
        string GetDefaultPrinter()
        {
            try
            {
                string str = Registry.CurrentUser.OpenSubKey(dfpath).GetValue("Device").ToString();
                str = str.Substring(0, str.IndexOf(','));
                int lastIndOf = str.LastIndexOf("\\"); 
                return lastIndOf == -1 ? str : str.Substring(lastIndOf + 1); 
            }
            catch (Exception ex)
            {
                File.WriteAllText(@"C:\test\err.txt", ex.StackTrace.ToString(), Encoding.ASCII);
                return "";
            }
        }
        //Удаление принтера
        public bool RemovePrinter(string name)
        {
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
                File.WriteAllText(@"C:\test\err_del_pr.txt", ex.StackTrace.ToString(), Encoding.ASCII);
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
        private void ReadFileAndRemove_Click(object sender, EventArgs e)
        {
            defaulPrinter = GetDefaultPrinter();
            File.WriteAllText(@"C:\test\exclude.txt", defaulPrinter, Encoding.ASCII);         
            string[] excludePorts = File.ReadAllLines(excludeCSV);
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

        private void Label2_Click(object sender, EventArgs e)
        {

        }
    }
}
