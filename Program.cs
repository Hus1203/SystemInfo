using Microsoft.Win32;
using System;
using System.IO;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;

namespace ConsoleApp2
{
    class Program
    {
        static string ServerName = "Server1"; // Имя сервера для нахождения расшаренных папок
        static void Main()
        {
            string outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ComputerInfoFiles");
            Directory.CreateDirectory(outputFolder);

            string sessionName = Environment.UserName;
            string datePart = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"ComputerInfo_{datePart}_{sessionName}.txt";
            string outputPath = Path.Combine(outputFolder, fileName);

            // Собираем информацию о железе компьютера
            string hardwareInfo = GetHardwareInformation();

            // Собираем информацию о установленных программах
            string softwareInfo = GetSoftwareInformation();

            // Собираем информацию о сетевых настройках
            string networkInfo = GetNetworkInformation();

            string Shared = GetShared();

            // Запишем информацию в новый файл
            WriteToFile(outputPath, "Информация о железе", hardwareInfo);
            WriteToFile(outputPath, "Установленные программы", softwareInfo);
            WriteToFile(outputPath, "Информация о сетевых настройках", networkInfo);
            WriteToFile(outputPath, "Расшаренные папки", Shared);

        }

        static string GetHardwareInformation()
        {
            string result = "";

            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Processor");  //  Делаем запрос
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query); // Осуществляем поиск по запросу 

            searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard");
            result += "\nМат.плата:\n";
            foreach (ManagementObject baseBoard in searcher.Get())
            {
                foreach (PropertyData prop in baseBoard.Properties)
                {
                    result += $"{prop.Name}: {prop.Value}\n";
                }
            }

            searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");
            result += "\nЖесткий диск:\n";
            foreach (ManagementObject diskDrive in searcher.Get())
            {
                result += $"Серийный номер: {diskDrive["SerialNumber"]}, Сигнатура: {diskDrive["Signature"]}\n";
            }

            searcher = new ManagementObjectSearcher("SELECT * FROM Win32_VideoController");
            result += "\nВидеокарта:\n";
            foreach (ManagementObject videoController in searcher.Get())
            {
                foreach (PropertyData prop in videoController.Properties)
                {
                    result += $"{prop.Name}: {prop.Value}\n";
                }
            }

            searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
            result += "\nИспользуемый процессор:\n";
            foreach (ManagementObject processor in searcher.Get())
            {
                result += $"Производитель: {processor["Manufacturer"]}, Модель: {processor["Name"]}\n";
            }

            searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory");
            result += "\nУстановленная физическая память:\n";
            foreach (ManagementObject memory in searcher.Get())
            {
                result += $"Тип: {memory["MemoryType"]}, Объем: {memory["Capacity"]} байт\n";
            }

            searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0");
            result += "\nПодключенные устройства:\n";
            foreach (ManagementObject device in searcher.Get())
            {
                result += $"Устройство: {device["Caption"]}, Тип: {device["Description"]}\n";
            }

            return result;
        } // Железо

        static string GetSoftwareInformation()
        {
            StringBuilder installedProgramsInfo = new StringBuilder();

            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string subKeyName in key.GetSubKeyNames())
                {
                    using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                    {
                        string displayName = subKey.GetValue("DisplayName") as string;
                        string installDate = subKey.GetValue("InstallDate") as string;
                        string installLocation = subKey.GetValue("InstallLocation") as string;
                        string displayVersion = subKey.GetValue("DisplayVersion") as string;
                        string size = subKey.GetValue("EstimatedSize") as string;
                        if (!string.IsNullOrEmpty(displayName))
                        {
                            installedProgramsInfo.AppendLine("Название: " + displayName);
                            installedProgramsInfo.AppendLine("Версия:" + displayVersion);
                            installedProgramsInfo.AppendLine("Дата установки: " + installDate);
                            installedProgramsInfo.AppendLine("Путь к программе:" + installLocation);
                            installedProgramsInfo.AppendLine("Размер (байт): " + size);
                            installedProgramsInfo.AppendLine();
                        }
                    }
                }
            }

            return installedProgramsInfo.ToString();
        } // Программы

        static string GetNetworkInformation()
        {
            string networkInfo = "\n";

            // Получаем название домена
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");  //  Делаем запрос
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query))
            {
                foreach (ManagementObject mo in searcher.Get())
                {
                    networkInfo += $"Домен: {mo["domain"]}\n\n";
                    networkInfo += $"DNS Имя хоста: {mo["DNSHostName"]}\n\n";
                }

            }

            ObjectQuery query1 = new ObjectQuery("SELECT * FROM Win32_NetworkAdapter");  //  Делаем запрос
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(query1))
            {
                networkInfo += "\nСетевые адаптеры\n";

                foreach (ManagementObject NetAdapter in searcher.Get())
                {
                    foreach (PropertyData prop in NetAdapter.Properties)
                    {
                        networkInfo += $"{prop.Name}: {prop.Value}\n";
                    }
                    networkInfo += $"\n\n";
                }
            }

            // Получаем информацию о сетевых интерфейсах
            NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces();
            networkInfo += "\nСетевые интерфейсы:\n";
            foreach (NetworkInterface inf in interfaces)
            {
                networkInfo += $"Имя: {inf.Name}, Тип: {inf.NetworkInterfaceType}, Скорость: {inf.Speed} bps\n";
            }

            // Получаем информацию о соединениях и IP-адресах
            networkInfo += "\nIP-адреса и соединения:\n";
            foreach (NetworkInterface inf in interfaces)
            {
                IPInterfaceProperties ipProps = inf.GetIPProperties();
                foreach (UnicastIPAddressInformation ip in ipProps.UnicastAddresses)
                {
                    networkInfo += $"Адрес: {ip.Address}, Маска подсети: {ip.IPv4Mask}\n\n";
                }
                foreach (GatewayIPAddressInformation gateway in ipProps.GatewayAddresses)
                {
                    networkInfo += $"Шлюз: {gateway.Address}";
                }
                IPInterfaceStatistics mac = inf.GetIPStatistics(); 
                networkInfo += $"MAC Адрес: {inf.GetPhysicalAddress()}\n"; 
                
            }

            return networkInfo;
        } // Сеть

        static string GetShared()
        {
            string Shared = "\n";

            using (ManagementClass shares = new ManagementClass($@"\\{ServerName}", "Win32_Share", new ObjectGetOptions()))
            {
                foreach (ManagementObject share in shares.GetInstances())
                {
                    Shared += $"Расшаренная папка: {share["Name"]}\n";
                    Shared += $"Путь к папке: {share["Path"]}\n\n";
                }
            }
            return Shared;
        }  // Расшаренные папки

        static void WriteToFile(string filePath, string sectionTitle, string sectionContent)
        {
            using (StreamWriter writer = File.AppendText(filePath))
            {
                writer.WriteLine(sectionTitle + ":");
                writer.WriteLine(sectionContent);
                writer.WriteLine("\n");
            }
        } // Метод записи в файл
    }
}
