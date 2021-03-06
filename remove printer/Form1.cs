﻿using Microsoft.Win32;
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
        private static string logonUI = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI";
        private string hklmPrinters = @"SYSTEM\CurrentControlSet\Control\Print\Printers";
        private static Logger logger = LogManager.GetCurrentClassLogger();
        //Определение каталога запуска приложения
        private static string startupPath = Application.StartupPath;
        //Путь до файла с исключениями
        private string excludeCSV = $@"{startupPath}\exclude.csv";
        //Пути до файлов с логами
        private string def_prn = $@"{startupPath}\default.txt";
        private string def_prn_sys = $@"{startupPath}\default_sys.txt";
        string[] listPrinters;
        private string defaulPrinter = "";
        private ManagementScope managementScope = null;
        private ManagementObjectCollection managementObjectCollection = null;
        string[] exclude;
        bool isDebug = false; //false;

        Stopwatch stopWatch = new Stopwatch();
        
        //Методы
        public Form1()
        {
            InitializeComponent();
            ReadArgs();
            stopWatch.Start();
            EulaAcceptedPsGetsid();
        }

        string IsUserRunApp()
        {
            string userName = "";
            using (WindowsIdentity wi = WindowsIdentity.GetCurrent())
            {
                userName = wi.Name;
            }
            return userName;
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
            return $@"{SidThisUser(IsUserRunApp())}\Software\Microsoft\Windows NT\CurrentVersion\Devices\";
        }

        bool IsSystemCheck()
        {
            bool owner;
            using (WindowsIdentity wi = WindowsIdentity.GetCurrent())
            {
                owner = wi.IsSystem;
            }

            return owner;
        }

        private void EulaAcceptedPsGetsid()
        {            
            RegistryKey psGetSid = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Sysinternals\PsGetSid");
            psGetSid.SetValue("EulaAccepted", 1);
            psGetSid.Close();
            logger.Info("psGetSid EulaAccepted");
        }

        private string LocAppData()
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return folderPath != null ? folderPath : "";
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
        
        //Получение коллекции объектов
        private ManagementObjectCollection GetManagementObject(string printerName)
        {
            ConnectionOptions options = new ConnectionOptions();
            options.EnablePrivileges = true;
            //options.Username = "lastLoggedOnUserSID";
            options.Impersonation = ImpersonationLevel.Impersonate;
            //options.Impersonation = ImpersonationLevel.Delegate;
            managementScope = new ManagementScope(ManagementPath.DefaultPath, options);
            managementScope.Connect();
            SelectQuery selectQuery = new SelectQuery();
            selectQuery.QueryString = @"SELECT * FROM Win32_Printer WHERE Name = '" + printerName.Replace("\\", "\\\\") + "'";
            ManagementObjectSearcher managementObjectSearcher = new ManagementObjectSearcher(managementScope, @selectQuery);
            return managementObjectSearcher.Get();
        }

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
        private void CreateDefaultTxt()
        {
            if (!IsSystemCheck())
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
        //Удаление принтера Старый
        public bool RemovePrinter1(string name)
        {
            string err = "";
            try
            {
       
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
                if (IsSystemCheck())//!IsStartIsLogon())
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
        private void ReadExclude()
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
            foreach (string nameOfPrinter in ListPrnKey())
            {
                string printerPort = GetPrinterPort(nameOfPrinter);
                byte countPrinters = 0;
                //string[] array3 = exclude;
                foreach (string port in exclude)
                {
                    if (printerPort.ToLower().Contains(port.ToLower()))
                    {
                        countPrinters++;
                        break;
                    }
                }
                if (countPrinters == 0)
                {
                    logger.Info("REMOVING PRINTER: " + nameOfPrinter);
                    RemovePrinter(nameOfPrinter);
                }
            }
        }

        private void SetDef()
        {
            byte countPrinters = 0;
            listPrinters = ListPrnKey();
          
            foreach (string setDef in listPrinters)
            {
                if (setDef.ToLower().Contains(defaulPrinter.ToLower()))
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
            logger.Info("Application started from " + IsUserRunApp() + ", logged in user " + SidThisUser(IsUserRunApp()) + ", startupPath: " + startupPath);
            if (!IsSystemCheck())
            {
                defaulPrinter = GetDefaultPrinter();
            }
            ReadExclude();
            CreateDefaultTxt();           
            //Проверка и удаление принтеров
            Remove();
            logger.Info("The process of checking and removing printers is complete.");
            SetDef();
            logger.Info("Default printer is setup");
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            logger.Info($"Time running: min {ts.Minutes} sec {ts.Seconds}");
            logger.Info("Exit");
            Close();
        }

        private void ReadFileAndRemove_Click(object sender, EventArgs e)
        {
           MainTask();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = "IsSystemCheck = " +IsSystemCheck().ToString();
            label2.Text = "LogonUSER(logonUI, \"LastLoggedOnSAMUser\") = " +LogonUSER(logonUI, "LastLoggedOnSAMUser");
            label3.Text = "IsUserRunApp = " + IsUserRunApp();
            //label4.Text = psWI;//SidThisUser(psWI);
            if (!isDebug)
            {
                MainTask();
            }

        }

        private void User_Click(object sender, EventArgs e)
        {
            ManagementScope theScope = new ManagementScope(ManagementPath.DefaultPath);

            ObjectQuery theQuery = new ObjectQuery("SELECT username FROM Win32_ComputerSystem");

            ManagementObjectSearcher theSearcher = new ManagementObjectSearcher(theScope, theQuery);

            ManagementObjectCollection theCollection = theSearcher.Get();

            foreach (ManagementObject theCurObject in theCollection)

            {

                MessageBox.Show(theCurObject["username"].ToString());

            }
        }
    }
}
