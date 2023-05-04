using System;
using System.IO;
using System.Net;
using System.Timers;
using CableRobot.Fins;

namespace Ed_FInsOmron
{
    class Program
    {
        static void Main(string[] args)
        {

            IPAddress ipAddressPlc1 = IPAddress.Parse("192.168.250.50"), ipAdrresPlc2 = IPAddress.Parse("192.168.250.56");
            int portPlc1 = 9600, portPlc2 = 9500;
            IPEndPoint endPointPlc1 = new IPEndPoint(ipAddressPlc1, portPlc1); 
            IPEndPoint endPointPlc2 = new IPEndPoint(ipAdrresPlc2, portPlc2);

            String sendedDatRouteLogAlb1 = "C:/Users/Usuario/Desktop/Code/Ed_FInsOmron/Ed_FInsOmron/SendedDataAlb1.txt";
            String sendedDatRouteLogAlb2 = "C:/Users/Usuario/Desktop/Code/Ed_FInsOmron/Ed_FInsOmron/SendedDataAlb2.txt";
            String receivedDatRouteLog = "C:/Users/Usuario/Desktop/Code/Ed_FInsOmron/Ed_FInsOmron/ReceivedData.txt";
            
            

            FinsClient conPlc1 = new FinsClient(endPointPlc1);
            Console.WriteLine("Plc1 Conectado");

            FinsClient conPlc2 = new FinsClient(endPointPlc2);
            Console.WriteLine("Plc2 Conectado");

            double ScanCycle = 200;

            UInt16[] readData = new UInt16[70];
            UInt16[] dataWorkCh1 = new UInt16[30];
            UInt16[] dataWorkCh2 = new UInt16[30];
            UInt16[] dataWorkCh3 = new UInt16[30];
            UInt16[] dataWorkCh4 = new UInt16[30];
            ushort[] clean = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            String cubeta = "";

            String dataSended = "";
            String dataReceived = "";

            Timer timer = new Timer(ScanCycle);
            timer.Elapsed += (sender, e) =>
            {
                DateTime horaActual = DateTime.Now;
                TimeSpan horaActualDelDia = horaActual.TimeOfDay;

                dataWorkCh1 = conPlc1.ReadWork(100,16);
                dataWorkCh2 = conPlc1.ReadWork(101,16);
                dataWorkCh3 = conPlc2.ReadWork(100,16);
                dataWorkCh4 = conPlc2.ReadWork(101,16);

                if (dataWorkCh1[0] == 1)
                {
                    readData = conPlc1.ReadData(1000,20);
                    conPlc1.WriteWork(100, clean);
                    for(int i = 0; i < 20; i++) {
                        dataSended = dataSended +  Convert.ToString(readData[i],16);
                    }
                    for (int i = 4; i <= 14; i++)
                    {
                        if (i % 2 == 0)
                        {
                            cubeta = cubeta + dataSended[i];
                        }
                    }
                    Console.WriteLine("Data Enviada: " + dataSended);
                    Console.WriteLine("Cubeta Enviada por Sunt: " + cubeta);
                    using (StreamWriter writer = File.AppendText(sendedDatRouteLogAlb1))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual +"] Trama enviada desde el Plc Albaranadora 1: " + dataSended);
                    }
                    dataSended = "";
                    cubeta = "";
                } 
                if (dataWorkCh2[0] == 1) 
                {
                    readData = conPlc1.ReadData(1414, 68);
                    conPlc1.WriteWork(52, clean);
;                   for(int i = 0; i < 68; i++) {
                        dataReceived = dataReceived +  Convert.ToString(readData[i],16);
                    }
                    
                    for(int i = 6; i <= 16; i++)
                    {
                        if(i%2 == 0) { 
                        
                            cubeta = cubeta + dataReceived[i];
                        }
                    }               

                    using (StreamWriter writer = File.AppendText(receivedDatRouteLog))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual +"] Trama recibida al plc: " + dataReceived);
                    }
                    Console.WriteLine("Data Recibida: " + dataReceived);
                    Console.WriteLine("Cubeta Enviada por Cofares: " + cubeta);
                    dataReceived = "";
                    cubeta = "";
                }
                if(dataWorkCh3[0] == 1) 
                {
                    readData = conPlc2.ReadData(1000,20);
                    conPlc2.WriteWork(100, clean);

                    for (int i = 0; i < 20; i++)
                    {
                        dataSended = dataSended + Convert.ToString(readData[i], 16);
                    }
                    Console.WriteLine("Data Enviada: " + dataSended);
                    Console.WriteLine("Cubeta Enviada por Sunt: " + cubeta);

                    using (StreamWriter writer = File.AppendText(sendedDatRouteLogAlb2))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual + "] Trama enviada desde el Plc Albaranadora 2: " + dataSended);
                    }
                    dataSended = "";
                    cubeta = "";
                }
                if (dataWorkCh4[0] == 1)
                {
                    readData = conPlc2.ReadData(1050, 68);
                    conPlc2.WriteWork(101, clean);
                    ; for (int i = 0; i < 68; i++)
                    {
                        dataReceived = dataReceived + Convert.ToString(readData[i], 16);
                    }

                    for (int i = 6; i <= 16; i++)
                    {
                        if (i % 2 == 0)
                        {

                            cubeta = cubeta + dataReceived[i];
                        }
                    }

                    using (StreamWriter writer = File.AppendText(receivedDatRouteLog))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual + "] Trama recibida al plc: " + dataReceived);
                    }
                    Console.WriteLine("Data Recibida: " + dataReceived);
                    Console.WriteLine("Cubeta Enviada por Cofares: " + cubeta);
                    dataReceived = "";
                    cubeta = "";
                }
            };

            timer.Enabled = true;
        }
    }
}
