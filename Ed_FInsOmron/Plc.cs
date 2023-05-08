using System;
using System.Net;
using System.IO;
using CableRobot.Fins;


    class Plc
    {
        String Ip = "";
        int Port = 0;
        IPAddress ipAddressPlc;
        IPEndPoint endPointPlc;
        FinsClient conPlc;

        public Plc(String Ip, int Port)
        {
            this.Ip = Ip;
            this.Port = Port;
        }

        public void connect()
        {
            this.ipAddressPlc = IPAddress.Parse(Ip);
            this.endPointPlc = new IPEndPoint(ipAddressPlc, this.Port);

            conPlc = new FinsClient(endPointPlc);

            Console.WriteLine("Plc con Ip " + this.Ip + " se ha conectado.");
        }

        /// <summary>
        /// Read the specified Area starting at specified Memory of the plc.
        /// </summary>
        /// <param name="area">"w" to read w Area. "d" to read DM area.</param>
        /// <param name="startMemory">The method star to read in this number of area selected</param>
        /// <param name="bitsToRead">Quantity of bits want to read</param>
        /// <returns>An Aray Uint16[] filled of data from the Plc.</returns>
        public UInt16[] read(String area, ushort startMemory, ushort bitsToRead)
        {
            UInt16[] hostArray = new UInt16[bitsToRead];

            if(area.Equals("w"))
            {
                hostArray = conPlc.ReadWork(startMemory, bitsToRead);
            } else if (area.Equals("d"))
            {
                hostArray = conPlc.ReadData(startMemory, bitsToRead);
            }

            return hostArray;

        }

        /// <summary>
        /// Read the specified Area starting at specified Memory of the plc.
        /// </summary>
        /// <param name="area">"w" to read w Area. "d" to read DM area.</param>
        /// <param name="startMemory">The method star to read in this number of area selected</param>
        /// <param name="dataToWrite">Data that will send to Plc</param>
        public void  write(String area, ushort startMemory, ushort[] dataToWrite)
        {
            if(area.Equals("w")) {

                conPlc.WriteWork(startMemory, dataToWrite);
            } else if (area.Equals("d"))
            {
                conPlc.WriteData(startMemory, dataToWrite);
            }
        }


    }
