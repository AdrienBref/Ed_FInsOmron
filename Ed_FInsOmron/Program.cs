using System;
using System.IO;
using System.Net;
using System.Timers;
using CableRobot.Fins;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using Ed_FInsOmron.Excel;
using System.Collections;

namespace Ed_FInsOmron
{
    class Program
    {
        static void Main(string[] args)
        {

            Plc conPlc1 = new Plc("10.50.10.106", 9600);
            Plc conPlc2 = new Plc("10.50.10.140", 9600);

            conPlc1.connect();
            conPlc2.connect();

            String sendedDatRouteLogAlb1 = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/SendedDataAlb1.txt";
            String sendedDatRouteLogAlb2 = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/SendedDataAlb2.txt";
            String receivedDatRouteLog = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/ReceivedData.txt";

            String loggingNormon = "C:/code/github.com/AdrienBref/Ed_FinsOmron/Ed_FInsOmron/resources/LoggingNormon.xlsx";
            IExcelManager excelManager = new ExcelManager(loggingNormon);

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

            ArrayList streamPackage = new ArrayList();

            String dataSended = "";
            String dataReceived = "";

            int filaTabla1 = 3, filaTabla2 = 3, filaTabla3 = 3;

            Timer timer = new Timer(ScanCycle);
            timer.Elapsed += (sender, e) =>
            {
                DateTime horaActual = DateTime.Now;
                TimeSpan horaActualDelDia = horaActual.TimeOfDay;

                dataWorkCh1 = conPlc1.read("w",100,16);
                dataWorkCh2 = conPlc1.read("w",101,16);
                dataWorkCh3 = conPlc2.read("w",100,16);
                dataWorkCh4 = conPlc2.read("w",101,16);

                if (dataWorkCh1[0] == 1)
                {
                    readData = conPlc1.read("d",1000,20);
                    conPlc1.write("w",100, clean);
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

                    for (int i = 0; i < completeStream.Length; i++)
                    {
                        if (i < 17)
                        {
                            bulto = bulto + completeStream[i];
                            streamPackage.Add(bulto);
                        }
                        else if (i == 18 && i < 20)
                        {
                            resultadoLectura = resultadoLectura + completeStream[i];
                            streamPackage.Add(resultadoLectura);
                        }
                    }

                    Console.WriteLine("Data Enviada: " + dataSended);
                    Console.WriteLine("Cubeta Enviada por Sunt: " + completeStream);

                    using (StreamWriter writer = File.AppendText(sendedDatRouteLogAlb1))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual +"] Trama enviada desde el Plc Albaranadora 1: " + dataSended);
                    }

                    int j = 6;
                    foreach (string stream in streamPackage)
                    {
                        excelManager.WriteData(filaTabla1, j + 1, stream);
                    }

                    filaTabla2++;


                    dataSended = "";
                    completeStream = "";
                } 
                if (dataWorkCh2[0] == 1) 
                {
                    readData = conPlc1.read("d", 1050, 68);
                    conPlc1.write("w", 101, clean);
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
                            streamPackage.Add(bulto);
                        } else if (i == 19 && i < 20)
                        {
                            hojasALeer = hojasALeer + completeStream[i];
                            streamPackage.Add(hojasALeer);

                        } else if (i == 21 && i < 22)
                        {
                            hojasTotales = hojasTotales + completeStream[i];
                            streamPackage.Add(hojasTotales);
                        }
                        else if (i == 23 && i < 24)
                        {
                            numeroMaquina = numeroMaquina + completeStream[i];
                            streamPackage.Add(numeroMaquina);

                        } 
                    }

                    streamPackage.Add(completeStream);

                    Console.WriteLine("Bulto: " + bulto);
                    Console.WriteLine("Hojas a leer: " + hojasALeer);
                    Console.WriteLine("Hojas totales: " + hojasTotales);
                    Console.WriteLine("Numero máquina: " + numeroMaquina);

                    int j = 0;
                    foreach (string stream in streamPackage) {
                        excelManager.WriteData(filaTabla1, j + 1, stream);
                    }

                    filaTabla1++;

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
                    readData = conPlc2.read("d",1000,20);
                    conPlc2.write("w",100, clean);

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
                            streamPackage.Add(bulto);
                        } else if (i == 18 && i < 20)
                        {
                            resultadoLectura = resultadoLectura + completeStream[i];
                            streamPackage.Add(resultadoLectura);
                        }
                    }

                    streamPackage.Add(completeStream);

                    Console.WriteLine("Bulto: " + bulto);
                    Console.WriteLine("Resultado Lectura: " + resultadoLectura);


                    int j = 12;
                    foreach (string stream in streamPackage)
                    {
                        excelManager.WriteData(filaTabla1, j + 1, stream);
                    }

                    filaTabla3++;

                    using (StreamWriter writer = File.AppendText(sendedDatRouteLogAlb2))
                    {
                        writer.WriteLine("[LOG]>> [" + horaActual + "] Trama enviada desde el Plc Albaranadora 2: " + dataSended);
                    }
                    dataSended = "";
                    completeStream = "";
                }
                if (dataWorkCh4[0] == 1)
                {
                    readData = conPlc2.read("d",1050, 68);
                    conPlc2.write("w",101, clean);
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
                //streamPackage.Clear();
            };

            timer.Enabled = true;
        }
    }
}
