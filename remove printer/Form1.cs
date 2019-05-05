using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management;
using System.Windows.Forms;

namespace remove_printer
{
    public partial class Form1 : Form
    {
        //Расположение дефолтного принтера пользователя в реестре
        readonly string dfpath = @"Software\Microsoft\Windows NT\CurrentVersion\Windows\";
        //Расположение всех принтеров пользователя в реестре
        readonly string dvpath = @"Software\Microsoft\Windows NT\CurrentVersion\Devices\";

        string excludeCSV = @"C:\test\exclude.csv";

        readonly string prn_es = "The default printer is ";
        readonly string prn_ru = "Принтер по умолчанию ";
        string def_prn;
        //string def = "";
        //string new_def = "";
        //Локаль ОС
        readonly static string loc = System.Threading.Thread.CurrentThread.CurrentCulture.ToString();
        readonly string path = $@"C:\Windows\System32\Printing_Admin_Scripts\{loc}\";
        private ManagementScope managementScope = null;
        Process process;
        string[] listPrn;
        List<ListOfPrinter> lst = new List<ListOfPrinter>();
        string namePrinter = "";
        string portName = "";
        string namePrinterUs = "Printer name ";
        string portNameUs = "Port name ";
        string namePrinterRu = "Имя принтера ";
        string portNameRu = "Имя порта ";

        //ExcludePort excludePort = new ExcludePort();

        public Form1()
        {
            InitializeComponent();
            LocationUX();
            label1.Text = loc;
            label2.Text = path;
        }
        //Определение локали и присвоение соответстующих переменных
        private void LocationUX()
        {
            if (loc == "ru-RU")
            {
                def_prn = prn_ru;
                namePrinter = namePrinterRu;
                portName = portNameRu;
            }

            else
            {
                def_prn = prn_es;
                namePrinter = namePrinterUs;
                portName = portNameUs;
            }
        }

        //Принтер по-умолчанию
        //private void DefaultPrinter()
        //{
        //    string str = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(dfpath).GetValue("Device").ToString();

        //    str = str.Substring(0, str.IndexOf(','));
        //    int lastIndOf = str.LastIndexOf("\\");
        //    //Доработать
        //    if (lastIndOf == -1)
        //        label1.Text = str;
        //    else
        //        label2.Text = str.Substring(lastIndOf + 1);
        //}


        private string DefaultPrinter()
        {
            string str = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(dfpath).GetValue("Device").ToString();

            str = str.Substring(0, str.IndexOf(','));
            int lastIndOf = str.LastIndexOf("\\");
            //Доработать
            if (lastIndOf == -1)
                label1.Text = str;
            else
                label1.Text = str.Substring(lastIndOf + 1);
            return str;
        }
        //Список принтеров пользователя
        void ListPrinter()
        {
            listPrn = Registry.CurrentUser.OpenSubKey(dvpath).GetValueNames();
            string t = "";
            foreach (string txt in listPrn)
            {
                t += txt + "\n";
            }
            label1.Text = t;
            label2.Text = listPrn.Length.ToString();
            //return listPrn;
        }
        //Удаление принтера
        public bool RemovePrinter(string name)
        {
            try
            {
                //use the ManagementScope class to connect to the local machine
                ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath);
                scope.Connect();

                //query Win32_Printer
                SelectQuery query = new SelectQuery("select * from Win32_Printer");

                ManagementObjectSearcher search = new ManagementObjectSearcher(scope, query);

                //get all the printers for local machine
                ManagementObjectCollection printers = search.Get();
                foreach (ManagementObject printer in printers)
                {
                    //get the name of the current printer in the list
                    string printerName = printer["Name"].ToString().ToLower();

                    //check for a match, if found then delete it or continue
                    //to the next printer in the list
                    if (printerName.Equals(name.ToLower()))
                    {
                        label2.Text = printerName;
                        //printer.Delete();
                        break;
                    }
                    else
                        continue;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error deleting printer: {0}", ex.Message));
                return false;
            }

        }
        private void Button1_Click(object sender, System.EventArgs e)
        {
            ListPrinter();
        }
        private void Delete_Click(object sender, EventArgs e)
        {
            //RemovePrinter(name);
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            string[] separatingStrings = { "\r", "\n" };
            process = Process.Start(new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = $"/c chcp 1251 & cscript /nologo {path}prnmngr.vbs -l",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            });
            //string list = process.StandardOutput.ReadToEnd();
            //label2.Text = list;
            string[] arrPrinters = process.StandardOutput.ReadToEnd().Split(separatingStrings, StringSplitOptions.RemoveEmptyEntries);
            string text = "";
            string nP = "";
            string pN = "";
            byte count = 0;

            foreach (string txt in arrPrinters)
            {
                if (txt.StartsWith(namePrinter))
                {

                    nP = txt;
                    count++;
                }
                else if (txt.StartsWith(portName))
                {
                    pN = txt;
                    count++;
                }
                if (count == 2)
                {
                    lst.Add(new ListOfPrinter(nP, pN));
                    count = 0;
                }
            }
            foreach (ListOfPrinter t in lst)
            {
                text += t.ToString() + "\n";
            }


            label1.Text = text;
        }

        private void Default_Click(object sender, EventArgs e)
        {
            byte count = 0;
            foreach (string port in ExcludePort.list)
            {
                foreach (ListOfPrinter vs in lst)
                {
                    if (!port.Contains(vs.PortName))
                    {
                        RemovePrinter(vs.NamePrinter);
                        label2.Text += $"Удален принтер {vs.NamePrinter}\n";
                    }
                }
                label2.Text += $"Цикл проверки {count++}\n";
            }
            label1.Text = "Проверка закончена";
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            GetPrinterInfo(DefaultPrinter());
        }

        public void GetPrinterInfo(string printerName)
        {
            managementScope = new ManagementScope(ManagementPath.DefaultPath);
            managementScope.Connect();

            SelectQuery selectQuery = new SelectQuery();
            selectQuery.QueryString = @"SELECT * FROM Win32_Printer 
	            WHERE Name = '" + printerName.Replace("\\", "\\\\") + "'";

            ManagementObjectSearcher managementObjectSearcher =
               new ManagementObjectSearcher(managementScope, @selectQuery);
            ManagementObjectCollection managementObjectCollection = managementObjectSearcher.Get();

            foreach (ManagementObject managementItem in managementObjectCollection)
            {
                //Console.WriteLine("Name : " + oItem["Name"].ToString());
                //Console.WriteLine("PortName : " + oItem["PortName"].ToString());
                //Console.WriteLine("DriverName : " + oItem["DriverName"].ToString());
                //Console.WriteLine("DeviceID : " + oItem["DeviceID"].ToString());
                //Console.WriteLine("Shared : " + oItem["Shared"].ToString());
                //Console.WriteLine("---------------------------------------------------------------");
                label1.Text = managementItem["Name"].ToString();
                label2.Text = managementItem["PortName"].ToString();
            }

        }

        private void ReadFileAndRemove_Click(object sender, EventArgs e)
        {

        }
    }
}
