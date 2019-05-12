using Microsoft.Win32;
using NLog;
using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Text;
using System.Windows.Forms;


namespace remove_printer
{
    public partial class Form1 : Form
    {
        //Переменные
        static string logonUI = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI";
        static string profileList = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList";
        //static string lastLoggedOnUserSID = SidThisUser(LogonUSER(logonUI, "LastLoggedOnSAMUser"));

        static string lastLoggedOnUserSID = SidThisUser(LogonUSER(logonUI, "LastLoggedOnSAMUser"));
        private readonly string dfpath = $@"{lastLoggedOnUserSID}\Software\Microsoft\Windows NT\CurrentVersion\Windows";
        private readonly string dvpath = $@"{lastLoggedOnUserSID}\Software\Microsoft\Windows NT\CurrentVersion\Devices\";

        private static Logger logger = LogManager.GetCurrentClassLogger();
        //Определение каталога запуска приложения
        private static string startupPath = Application.StartupPath;

        //Путь до файла с исключениями
        private string excludeCSV = $@"{startupPath}\exclude_test.csv";
        //Пути до файлов с логами
        //private string err_def = $@"{startupPath}\err_def.txt";
        //private string err_del_pr = $@"{startupPath}\err_del_pr.txt";
        private string err_ex = $@"{startupPath}\err_ex.txt";
        private string def_prn = $@"{startupPath}\default.txt";
        private string del = $@"{startupPath}\del.txt";
        private string pspath = $@"{startupPath}\PSTool\";

        string ps = Environment.UserName;
        
        private string defaulPrinter = "";
        private ManagementScope managementScope = null;
        private ManagementObjectCollection managementObjectCollection = null;
        private string[] listPrn;
        //Методы
        public Form1()
        {
            InitializeComponent();
            EulaAcceptedPsGetsid();
        }
        private void EulaAcceptedPsGetsid()
        {
            RegistryKey currentUserKey = Registry.CurrentUser;
            RegistryKey psGetSid = currentUserKey.CreateSubKey(@"SOFTWARE\Sysinternals\PsGetSid");
            psGetSid.SetValue("EulaAccepted", 1);
            psGetSid.Close();
        }
        //Проверка от кого запущено приложение и кто вошел в систему
        private bool IsStartIsLogon()
        {
            logger.Info($"Приложение запущено от {ps}, залогиненный пользователь {lastLoggedOnUserSID}");
            return ps == lastLoggedOnUserSID;
        }
        public static string ShellProcessCommandLine(string cmdLineArgs, string path)
        {
            
            StringBuilder sb = new StringBuilder();
            Process process = new Process
            {
                StartInfo =
            {
                //WorkingDirectory = @"c:\WINDOWS\System32\",
                FileName = "cmd.exe",
                CreateNoWindow = true,
                Arguments = $"/C PsGetsid.exe {cmdLineArgs} -nobanner",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                UseShellExecute = false
            }
            };
            process.OutputDataReceived += delegate (object sender, DataReceivedEventArgs args)
            {
                sb.AppendLine(args.Data);
            };
            process.Start();
            process.BeginOutputReadLine();
            process.WaitForExit();
            return sb.ToString();
        }
        //Получение sid пользователя
        private static string SidThisUser(string logonUser)
        {
            string value = "";
            string sid ="";
            string userInSid = "";

            int lastIndOf = logonUser.LastIndexOf("\\");
            string user = lastIndOf == -1 ? logonUser : logonUser.Substring(lastIndOf + 1);
            
            //string str = user + "\n";
            //string user = logonUser.LastIndexOf("\\") == -1 ? logonUser : logonUser.Substring(lastIndOf + 1);

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(profileList))
            {
                foreach (string val in key.GetSubKeyNames())
                {
                    using (RegistryKey subKey = Registry.LocalMachine.OpenSubKey(profileList +"\\" + val))
                    {
                        value = subKey.GetValue("ProfileImagePath").ToString();
                    }
                    userInSid = value.StartsWith(@"C:\Users\") ? value.Substring(@"C:\Users\".Length) : value;
                    //str += $"user {userInSid} sid {val}\n";                    
                    //str += (value.StartsWith(@"C:\Users\") ? value.Substring(@"C:\Users\".Length) : value) + "\n";
                    if (user.Equals(userInSid))
                    {
                        sid = val;
                        break;
                    }
                }
            };

            return sid;
        }
        
        //Получение коллекции объектов
        private ManagementObjectCollection GetManagementObject(string printerName)
        {
            ConnectionOptions options = new ConnectionOptions();
            options.EnablePrivileges = true;
            //options.Username = "lastLoggedOnUserSID";
            options.Impersonation = ImpersonationLevel.Impersonate;
            managementScope = new ManagementScope(ManagementPath.DefaultPath, options);
            managementScope.Connect();
            SelectQuery selectQuery = new SelectQuery();
            selectQuery.QueryString = @"SELECT * FROM Win32_Printer 
	            WHERE Name = '" + printerName.Replace("\\", "\\\\") + "'";
            ManagementObjectSearcher managementObjectSearcher =
               new ManagementObjectSearcher(managementScope, @selectQuery);
            return managementObjectSearcher.Get();
        }
        //установка принтера по умолчанию через реестр
        private void SetDefaultPrinterOnReg(string name)
        {
            //Дописать
            string str = "";
            using (RegistryKey regKey = Registry.Users.OpenSubKey(dfpath))
            {
                str = regKey.GetValue("Device").ToString();
            }
        }
        
        //Получение залогиненного пользователя
        static string LogonUSER(string logonUI, string value)
        {
            string str = "";
            string user = "";

            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(logonUI))
                {
                    str = key.GetValue(value).ToString();
                };
                user = str.StartsWith(".\\") ? str.Substring(2) : str;
                //user = str;
                //logger.Info($"Logged in {user}");
                return user;
            }
            catch (Exception ex)
            {
                //logger.Info(ex, "Could not get user who is logged in pc");
                return "";
            }
        }

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
                logger.Error(ex, "Не вышло получить принтер по умолчанию");
                return "";
            }
        }
        
        //Удаление принтера Старый
        public bool RemovePrinter(string name)
        {
            string err = "";
            try
            {
                //Использование ManagementScope class для подключения к арм

                //ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath);
                //scope.Connect();
                ////Запрос Win32_Printer
                //SelectQuery query = new SelectQuery("select * from Win32_Printer");
                //ManagementObjectSearcher search = new ManagementObjectSearcher(scope, query);
                ////Получение всех принтеров на арм
                //ManagementObjectCollection printers = search.Get();   printers
                foreach (ManagementObject printer in GetManagementObject(name))
                {
                    //получение имени конкретного принтера            
                    //проверить совпадение, если найдено, удалить или продолжить                    
                    if (printer["Name"].ToString().ToLower().Equals(name.ToLower()))
                    {
                        err = printer["Name"].ToString();
                        printer.Delete();
                        //logger.Info($"Принтер {err} удален");
                        break;
                    }
                    else
                        continue;
                }
                return true;
            }
            catch (Exception ex)
            {
                //logger.Error(ex, "Something bad happened");
                //logger.Error($"Не удалось удалить {err}");
                //File.AppendAllText(err_del_pr, $"Не вышло удалить {err}", Encoding.Unicode);
                return false;
            }

        }

        public bool RemovePrinter1(string name)
        {

            ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath);
            scope.Connect();
            SelectQuery query = new SelectQuery("select * from Win32_Printer");
            ManagementObjectSearcher search = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection printers = search.Get();
            foreach (ManagementObject printer in printers)
            {
                string printerName = printer["Name"].ToString().ToLower();
                logger.Info($"Принтеры на ПК {printerName}");
                if (printerName.Equals(name.ToLower()))
                {
                    printer.Delete();
                    break;
                }
            }

            return true;
        }
        
        public string GetPrinterPort(string printerName)
        {
            //managementObjectCollection = GetManagementObject(printerName);
            string portName = "";
            foreach (ManagementObject managementItem in GetManagementObject(printerName))
            {
                portName = managementItem["PortName"].ToString();
                //logger.Info($"Принтер {printerName} порт {portName}");
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
            catch (Exception ex)
            {
                
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
            label3.Text = SidThisUser(LogonUSER(logonUI, "LastLoggedOnSAMUser"));
            //Environment.UserName
            label4.Text = ShellProcessCommandLine(ps, startupPath);
            //ShellProcessCommandLine(LogonUSER(logonUI, "LastLoggedOnSAMUser"), startupPath);
            //MainTask();
        }
    }
}
