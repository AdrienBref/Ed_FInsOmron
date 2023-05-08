using System;
using System.IO;
using System.Net;
using System.Timers;
using CableRobot.Fins;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace Ed_FInsOmron
{
    class Program
    {
        static void Main(string[] args)
        {

            IPAddress ipAddressPlc1 = IPAddress.Parse("10.50.10.106"), ipAdrresPlc2 = IPAddress.Parse("10.50.10.140");
            int portPlc1 = 9600, portPlc2 = 9600;
            IPEndPoint endPointPlc1 = new IPEndPoint(ipAddressPlc1, portPlc1); 
            IPEndPoint endPointPlc2 = new IPEndPoint(ipAdrresPlc2, portPlc2);

            String sendedDatRouteLogAlb1 = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/SendedDataAlb1.txt";
            String sendedDatRouteLogAlb2 = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/SendedDataAlb2.txt";
            String receivedDatRouteLog = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/ReceivedData.txt";

            String loggingNormon = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/LoggingNormon.xlsx";
            
            FinsClient conPlc1 = new FinsClient(endPointPlc1);
            Console.WriteLine("Plc1 Conectado");

            FinsClient conPlc2 = new FinsClient(endPointPlc2);
            Console.WriteLine("Plc2 Conectado");

            XSSFWorkbook LogNormon;

            using (FileStream file = new FileStream(loggingNormon, FileMode.Open, FileAccess.Read))
            {
                LogNormon = new XSSFWorkbook(file);
            }

            ISheet sheet = LogNormon.GetSheetAt(0);

            double ScanCycle = 200;

            UInt16[] readData = new UInt16[70];
            UInt16[] dataWorkCh1 = new UInt16[30];
            UInt16[] dataWorkCh2 = new UInt16[30];
            UInt16[] dataWorkCh3 = new UInt16[30];
            UInt16[] dataWorkCh4 = new UInt16[30];
            ushort[] clean = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            String completeStream = "";
            String bulto = "";
            String hojasALeer = "";
            String hojasTotales = "";
            String numeroMaquina = "";
            String resultadoLectura = "";

            String dataSended = "";
            String dataReceived = "";

            int filaTabla1 = 3, filaTabla2 = 3, filaTabla3 = 3;

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
                            completeStream = completeStream + dataSended[i];
                        }
                    }
                    Console.WriteLine("Data Enviada: " + dataSended);
                    Console.WriteLine("Cubeta Enviada por Sunt: " + completeStream);
                    using (StreamWriter writer = File.AppendText(sendedDatRouteLogAlb1))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual +"] Trama enviada desde el Plc Albaranadora 1: " + dataSended);
                    }
                    dataSended = "";
                    completeStream = "";
                } 
                if (dataWorkCh2[0] == 1) 
                {
                    readData = conPlc1.ReadData(1050, 68);
                    conPlc1.WriteWork(101, clean);
;                   for(int i = 0; i < 68; i++) {
                        dataReceived = dataReceived +  Convert.ToString(readData[i],16);
                    }
                    
                    for(int i = 5; i <= 55; i++)
                    {
                        if(i%2 != 0) { 
                        
                            completeStream = completeStream + dataReceived[i];
                        }
                    }
                    
                    for(int i = 0; i < completeStream.Length; i++)
                    {
                        if(i < 18)
                        {
                            bulto = bulto + completeStream[i];
                        } else if (i == 19 && i < 20)
                        {
                            hojasALeer = hojasALeer + completeStream[i];
                       
                        } else if (i == 21 && i < 22)
                        {
                            hojasTotales = hojasTotales + completeStream[i];
                        }
                        else if (i == 23 && i < 24)
                        {
                            numeroMaquina = numeroMaquina + completeStream[i];

                        } 
                    }

                    Console.WriteLine("Bulto: " + bulto);
                    Console.WriteLine("Hojas a leer: " + hojasALeer);
                    Console.WriteLine("Hojas totales: " + hojasTotales);
                    Console.WriteLine("Numero máquina: " + numeroMaquina);

                    
                    IRow row = sheet.GetRow(filaTabla1) ?? sheet.CreateRow(filaTabla1);
                    ICell codigoCajaCell = row.GetCell(1) ?? row.CreateCell(1);
                    ICell hojasTotalesCell= row.GetCell(2) ?? row.CreateCell(2);
                    ICell hojasALeerCell = row.GetCell(3) ?? row.CreateCell(3);
                    ICell numMaqCell= row.GetCell(4) ?? row.CreateCell(4);
                    ICell tramaOriginalC= row.GetCell(5) ?? row.CreateCell(5);

                    codigoCajaCell.SetCellValue(bulto);
                    hojasTotalesCell.SetCellValue(hojasTotales);
                    hojasALeerCell.SetCellValue(hojasALeer);
                    numMaqCell.SetCellValue(numeroMaquina);
                    tramaOriginalC.SetCellValue(completeStream);

                    using (FileStream file = new FileStream(loggingNormon, FileMode.Create, FileAccess.Write))
                    {
                        LogNormon.Write(file);
                        filaTabla1++;
                    }

                    using (StreamWriter writer = File.AppendText(receivedDatRouteLog))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual +"] Trama recibida al plc: " + dataReceived);
                    }
                    Console.WriteLine("Data Recibida: " + dataReceived);
                    Console.WriteLine("Cubeta Enviada por Cofares: " + completeStream);
                    dataReceived = "";
                    completeStream = "";
                    bulto = "";
                    hojasALeer = "";
                    hojasTotales = "";
                    numeroMaquina = "";
                }
                if(dataWorkCh3[0] == 1) 
                {
                    readData = conPlc2.ReadData(1000,20);
                    conPlc2.WriteWork(100, clean);

                    for (int i = 0; i < 20; i++)
                    {
                        dataSended = dataSended + Convert.ToString(readData[i], 16);
                    }

                    for(int i = 1; i <= 40; i++)
                    {
                        if(i%2 == 0) { 
                        
                            completeStream = completeStream + dataSended[i];
                        }
                    }

                    Console.WriteLine("Data Enviada: " + dataSended);
                    Console.WriteLine("Cubeta Enviada por Sunt: " + completeStream);

                    for(int i = 0; i < completeStream.Length; i++)
                    {
                        if(i < 17)
                        {
                            bulto = bulto + completeStream[i];
                        } else if (i == 18 && i < 20)
                        {
                            resultadoLectura = resultadoLectura + completeStream[i];
                        }
                    }

                    Console.WriteLine("Bulto: " + bulto);
                    Console.WriteLine("Resultado Lectura: " + resultadoLectura);


                    IRow row = sheet.GetRow(filaTabla2) ?? sheet.CreateRow(filaTabla2);
                    ICell codigoCajaCell = row.GetCell(8) ?? row.CreateCell(8);
                    ICell resultadoLecturaCell= row.GetCell(9) ?? row.CreateCell(9);
                    ICell tramaOriginalC= row.GetCell(10) ?? row.CreateCell(10);

                    codigoCajaCell.SetCellValue(bulto);
                    resultadoLecturaCell.SetCellValue(resultadoLectura);
                    tramaOriginalC.SetCellValue(dataSended);

                    using (FileStream file = new FileStream(loggingNormon, FileMode.Create, FileAccess.Write))
                    {
                        LogNormon.Write(file);
                        filaTabla2++;
                    }

                    using (StreamWriter writer = File.AppendText(sendedDatRouteLogAlb2))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual + "] Trama enviada desde el Plc Albaranadora 2: " + dataSended);
                    }
                    dataSended = "";
                    completeStream = "";
                }
                if (dataWorkCh4[0] == 1)
                {
                    //readData = conPlc2.ReadData(1050, 68);
                    //conPlc2.WriteWork(101, clean);
                    ; for (int i = 0; i < 68; i++)
                    {
                        dataReceived = dataReceived + Convert.ToString(readData[i], 16);
                    }

                    for (int i = 6; i <= 16; i++)
                    {
                        if (i % 2 == 0)
                        {

                            completeStream = completeStream + dataReceived[i];
                        }
                    }

                    /*using (StreamWriter writer = File.AppendText(receivedDatRouteLog))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual + "] Trama recibida al plc: " + dataReceived);
                    }*/
                    Console.WriteLine("Data Recibida: " + dataReceived);
                    Console.WriteLine("Cubeta Enviada por Cofares: " + completeStream);
                    dataReceived = "";
                    completeStream = "";
                }
            };

            timer.Enabled = true;
        }
    }
}
