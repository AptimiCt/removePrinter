using Microsoft.Win32;
using NLog;
using System;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Management;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;


namespace remove_printer
{
    public partial class Form1 : Form
    {
        //Переменные
        static string logonUI = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI";
        //static string profileList = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList";
        string hklmPrinters = @"SYSTEM\CurrentControlSet\Control\Print\Printers";
        private static Logger logger = LogManager.GetCurrentClassLogger();
        //Определение каталога запуска приложения
        private static string startupPath = Application.StartupPath;
        //Путь до файла с исключениями
        private string excludeCSV = $@"{startupPath}\exclude.csv";
        //Пути до файлов с логами
        private string err_ex = $@"{startupPath}\err_ex.txt";
        private string def_prn = $@"{startupPath}\default.txt";
        private string def_prn_sys = $@"{startupPath}\default_sys.txt";
        //private string del = $@"{startuppath}\del.txt";
        //private string pspath = $@"{startupPath}\PSTool\";
        string[] listPrinters;
        //string ps = Environment.UserName;
        static string psWI = WindowsIdentity.GetCurrent().Name;
        private string defaulPrinter = "";
        private ManagementScope managementScope = null;
        private ManagementObjectCollection managementObjectCollection = null;
        //private string[] listPrn;
        string[] exclude;
        bool isDebug = false;
        //Методы
        public Form1()
        {
            InitializeComponent();
            ReadArgs();
            EulaAcceptedPsGetsid();
        }
        void ReadArgs()
        {
            string[] commandLineArgs = Environment.GetCommandLineArgs();
            foreach (string args in commandLineArgs)
            {
                if (commandLineArgs.Length == 2)
                {
                    if (args.ToLowerInvariant() == "-debug")
                    {
                        isDebug = true;
                    }
                }
            }
            
        }

        string DVPath()
        {
            return $@"{SidThisUser(psWI)}\Software\Microsoft\Windows NT\CurrentVersion\Devices\";
        }

        string DFPath()
        {
            return @"Software\Microsoft\Windows NT\CurrentVersion\Windows"; ;
        }

        private string HKLMPrinters()
        {
            return @"SYSTEM\CurrentControlSet\Control\Print\Printers";
        }
        private void EulaAcceptedPsGetsid()
        {            
            RegistryKey psGetSid = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Sysinternals\PsGetSid");
            psGetSid.SetValue("EulaAccepted", 1);
            psGetSid.Close();
            logger.Info("psGetSid EulaAccepted");
        }

        //Проверка от кого запущено приложение и кто вошел в систему
        private bool IsStartIsLogon()
        {

            bool check = psWI.Contains(LogonUSER(logonUI, "LastLoggedOnSAMUser"));
                logger.Info($"Checking IsStartIsLogon is {check}");
            return check;
        }

        private string LocAppData()
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return folderPath != null ? folderPath : "";
            //if (Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) != null)
            //{
                //return folderPath;
            //}
            //return "";
        }

        //Получение sid пользователя
        static string SidThisUser(string psGetsidLineArgs)//, string path)
        {
            
            StringBuilder sb = new StringBuilder();
            Process process = new Process
            {
                StartInfo =
                {
                    //WorkingDirectory = @"c:\WINDOWS\System32\",
                    FileName = "PsGetsid.exe",
                    CreateNoWindow = true,
                    Arguments = $"\"{psGetsidLineArgs}\" -nobanner",
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
            string[] sep = { "\n","\r"};
            string[] result = sb.ToString().Split(sep, StringSplitOptions.RemoveEmptyEntries);
            return result[1];
        }        
        //private static string SidThisUser1(string logonUser)
        //{
        //    string value = "";
        //    string sid ="";
        //    string userInSid = "";

        //    int lastIndOf = logonUser.LastIndexOf("\\");
        //    string user = lastIndOf == -1 ? logonUser : logonUser.Substring(lastIndOf + 1);
            
        //    //\string str = user + "\n";
        //    //\string user = logonUser.LastIndexOf("\\") == -1 ? logonUser : logonUser.Substring(lastIndOf + 1);

        //    using (RegistryKey key = Registry.LocalMachine.OpenSubKey(profileList))
        //    {
        //        foreach (string val in key.GetSubKeyNames())
        //        {
        //            using (RegistryKey subKey = Registry.LocalMachine.OpenSubKey(profileList +"\\" + val))
        //            {
        //                value = subKey.GetValue("ProfileImagePath").ToString();
        //            }
        //            userInSid = value.StartsWith(@"C:\Users\") ? value.Substring(@"C:\Users\".Length) : value;
        //            //\str += $"user {userInSid} sid {val}\n";                    
        //            //\str += (value.StartsWith(@"C:\Users\") ? value.Substring(@"C:\Users\".Length) : value) + "\n";
        //            if (user.Equals(userInSid))
        //            {
        //                sid = val;
        //                break;
        //            }
        //        }
        //    };

        //    return sid;
        //}
        
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
            selectQuery.QueryString = @"SELECT * FROM Win32_Printer WHERE Name = '" + printerName.Replace("\\", "\\\\") + "'";
            ManagementObjectSearcher managementObjectSearcher = new ManagementObjectSearcher(managementScope, @selectQuery);
            return managementObjectSearcher.Get();
        }

        //private ManagementObjectCollection GetManagementObject1(string printerName)
        //{
            //ConnectionOptions connectionOptions = new ConnectionOptions();
            //connectionOptions.EnablePrivileges = true;
            //connectionOptions.Impersonation = ImpersonationLevel.Impersonate;
            //managementScope = new ManagementScope(ManagementPath.DefaultPath, connectionOptions);
            //managementScope.Connect();
            //SelectQuery selectQuery = new SelectQuery();
            //selectQuery.QueryString = "SELECT * FROM Win32_Printer \r\n\t            WHERE Name = '" + printerName.Replace("\\", "\\\\") + "'";
            //return new ManagementObjectSearcher(managementScope, selectQuery).Get();
        //}



        //установка принтера по умолчанию через реестр
        //private void SetDefaultPrinterOnReg(string name)
        //{
            ////Дописать
            //string str = "";
            //using (RegistryKey regKey = Registry.Users.OpenSubKey(DFPath()))
            //{
                //str = regKey.GetValue("Device").ToString();
            //}
        //}

        //Получение залогиненного пользователя
        private static string LogonUSER(string logonUI, string regKey)
        {
            string user = "";

            try
            {
                using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(logonUI))
                {
                    user = registryKey.GetValue(regKey).ToString();
                }
                string result = user.StartsWith(".\\") ? user.Substring(2) : user;
                //logger.Info("Logged in " + result);
                return result;
            }
            catch (Exception ex)
            {
                logger.Info(ex, "Could not get user who is logged in pc");
                return "";
            }
        }
        //Создание файла defaut.txt
        private void createDefaultTxt()
        {
            if (IsStartIsLogon())
            {
                string path = LocAppData();
                if (Directory.Exists(path))
                {
                    File.AppendAllText(path + @"\default.txt", defaulPrinter + Environment.NewLine, Encoding.Unicode);
                }
                else
                {
                    File.AppendAllText(def_prn, defaulPrinter, Encoding.Unicode);
                }
            }
            else
            {
                File.AppendAllText(def_prn_sys, defaulPrinter, Encoding.Unicode);
            }
        }
        //Получение принтера по умолчанию

        string GetDefaultPrinter()
        {
            PrinterSettings settings = new PrinterSettings();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                settings.PrinterName = printer;
                if (settings.IsDefaultPrinter)
                    return printer;
            }
            return string.Empty;
        }
        private string GetDefaultPrinter1()
        {
            RegistryKey registryKey = null;
            try
            {
                string defPrn = "";
                logger.Info(" At the beginig GetDefaultPrinter: Application started from " + psWI + ", logged in user " + SidThisUser(psWI) + ", startupPath: " + startupPath + ", str: " + defPrn + ",DFPath(): " + DFPath());
                registryKey = Registry.CurrentUser.OpenSubKey(DFPath());
                if (registryKey != null)
                {
                    defPrn = registryKey.GetValue("Device").ToString();
                }
                registryKey.Close();

                logger.Info(" After GetDefaultPrinter: Application started from " + psWI + ", logged in user " + SidThisUser(psWI) + ", startupPath: " + startupPath + ", str: " + defPrn);
                defPrn = defPrn.Substring(0, defPrn.IndexOf(','));
                int num = defPrn.LastIndexOf("\\");
                return (num == -1) ? defPrn : defPrn.Substring(num + 1);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get default printer");
                registryKey.Close();
                return "";
            }
        }

        //Удаление принтера Старый
        public bool RemovePrinter1(string name)
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
                        logger.Info($"Printer {err} is remove");
                        break;
                    }
                    else
                        continue;
                }
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to remove printer {err}");
                //logger.Error($"Не удалось удалить {err}");
                //File.AppendAllText(err_del_pr, $"Не вышло удалить {err}", Encoding.Unicode);
                return false;
            }

        }

        public bool RemovePrinter(string name)
        {
            string printerName = "";
            ConnectionOptions options = new ConnectionOptions();
            options.EnablePrivileges = true;
            //options.Username = "lastLoggedOnUserSID";
            options.Impersonation = ImpersonationLevel.Impersonate;
            ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath, options);
            scope.Connect();
            SelectQuery query = new SelectQuery("select * from Win32_Printer");
            ManagementObjectSearcher search = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection printers = search.Get();
            try
            {
                
                foreach (ManagementObject printer in printers)
                {
                    printerName = printer["Name"].ToString().ToLower();
                    //logger.Info($"Printers on PC: {printerName}");
                    if (printerName.Equals(name.ToLower()))
                    {
                        
                        printer.Delete();
                        logger.Info($"Printer {printerName} is remove");
                        break;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to remove printer {printerName}");
                //logger.Error($"Не удалось удалить {err}");
                //File.AppendAllText(err_del_pr, $"Не вышло удалить {err}", Encoding.Unicode);
                return false;
            }
        }
        
        public string GetPrinterPort(string printerName)
        {
            //managementObjectCollection = GetManagementObject(printerName);
            string portName = "";
            foreach (ManagementObject managementItem in GetManagementObject(printerName))
            {
                portName = managementItem["PortName"].ToString();
                logger.Info($"Printer: {printerName} port: {portName}");
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

        private string[] ListPrnKey()
        {
            RegistryKey registryKey = null;
            string[] result = null;
            try
            {
                if (!IsStartIsLogon())
                {
                    registryKey = Registry.LocalMachine.OpenSubKey(hklmPrinters);
                    result = registryKey.GetSubKeyNames();
                    logger.Info("Checking printers in Registry.LocalMachine");
                }
                else
                {
                    registryKey = Registry.Users.OpenSubKey(DVPath());
                    result = registryKey.GetValueNames();
                    logger.Info("Checking printers in Registry.Users");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "No read registry");
                registryKey.Close();
            }
            registryKey.Close();
            return result;
        }
        private void readExclude()
        {
            
            if (File.Exists(excludeCSV))
            {
                exclude = File.ReadAllLines(excludeCSV);
                logger.Info("File excludeCSV is read");
            }
            else
            {
                logger.Error("No file exclude.csv");
                Close();
            }

        }

        private void Remove()
        {
            foreach (string nameOfPrinter in ListPrnKey())//listPrn)
            {
                string printerPort = GetPrinterPort(nameOfPrinter);
                byte countPrinters = 0;
                //string[] array3 = exclude;
                foreach (string text2 in exclude)
                {
                    if (printerPort.ToLower().Contains(text2.ToLower()))
                    {
                        countPrinters++;
                        break;
                    }
                }
                if (countPrinters == 0)
                {
                    logger.Info("Removing printer " + nameOfPrinter);
                    RemovePrinter(nameOfPrinter);
                }
            }
        }

        private void SetDef()
        {
            byte countPrinters = 0;
            listPrinters = ListPrnKey();
            //array2 = listPrinters;
            foreach (string setDef in listPrinters)
            {
                if (defaulPrinter.ToLower().Contains(setDef.ToLower()))
                {
                    defaulPrinter = setDef;
                    break;
                }
                countPrinters++;
            }
            if (countPrinters == listPrinters.Length)
            {
                SetDefaultPrinter("SafeQ");
            }
            else if (defaulPrinter.Length <= 1)
            {
                SetDefaultPrinter("SafeQ");
            }
            else
            {
                SetDefaultPrinter(defaulPrinter);
            }
        }
        private void MainTask()
        {
            logger.Info("Application started from " + psWI + ", logged in user " + SidThisUser(psWI) + ", startupPath: " + startupPath);
            defaulPrinter = GetDefaultPrinter();
            readExclude();
            createDefaultTxt();
            //listPrn = ListPrnKey();
            //foreach(string prn in listPrn)
            //{
            //logger.Info($"Printers: {prn}");
            //}


            //string[] array2 = listPrn;
            //foreach (string nameOfPrinter in listPrn)
            //{
            //    string printerPort = GetPrinterPort(nameOfPrinter);
            //    byte countPrinters = 0;
            //    //string[] array3 = exclude;
            //    foreach (string text2 in exclude)
            //    {
            //        if (printerPort.ToLower().Contains(text2.ToLower()))
            //        {
            //            countPrinters++;
            //            break;
            //        }
            //    }
            //    if (countPrinters == 0)
            //    {
            //        logger.Info("Removing printer " + nameOfPrinter);
            //        RemovePrinter(nameOfPrinter);
            //    }
            //}
            //Проверка и удаление принтеров
            Remove();
            SetDef();
            //byte b2 = 0;
            //listPrinters = ListPrnKey();
            ////array2 = listPrinters;
            //foreach (string text3 in listPrinters)
            //{
                //if (defaulPrinter.ToLower().Contains(text3.ToLower()))
                //{
                    //defaulPrinter = text3;
                    //break;
                //}
                //b2 = (byte)(b2 + 1);
            //}
            //if (b2 == listPrinters.Length)
            //{
                //SetDefaultPrinter("SafeQ");
            //}
            //else if (defaulPrinter.Length <= 1)
            //{
                //SetDefaultPrinter("SafeQ");
            //}
            //else
            //{
                //SetDefaultPrinter(defaulPrinter);
            //}
            logger.Info("Exit");
            Close();
        }

        private void ReadFileAndRemove_Click(object sender, EventArgs e)
        {
           MainTask();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label2.Text = LogonUSER(logonUI, "LastLoggedOnSAMUser");
            label3.Text = psWI;
            label4.Text = SidThisUser(psWI);
            if (!isDebug)
            {
                MainTask();
            }

        }
    }
}
