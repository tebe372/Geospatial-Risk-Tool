using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GeospatialRiskAnalysisTool
{
    public class MyStruct
    {
        /// <summary>
        /// Time
        /// </summary>
        public double TimeStampSF { get; set; }
        /// <summary>
        /// Pressure
        /// </summary>
        public double Pressure { get; set; }
        /// <summary>
        /// AccX
        /// </summary>
        public double AccelerationX { get; set; }
        /// <summary>
        /// AccY
        /// </summary>
        public double AccelerationY { get; set; }
        /// <summary>
        /// AccZ
        /// </summary>
        public double AccelerationZ { get; set; }
        /// <summary>
        /// magn(acc)
        /// </summary>
        public double Acceleration { get; set; }
        /// <summary>
        /// RotX
        /// </summary>
        public double RotationX { get; set; }
        /// <summary>
        /// RotY
        /// </summary>
        public double RotationY { get; set; }
        /// <summary>
        /// RotZ
        /// </summary>
        public double RotationZ { get; set; }
        /// <summary>
        /// deg/s
        /// </summary>
        public double Rotation { get; set; }
        public double Temperature { get; set; }
        public double VBattery { get; set; }
        public double Voltage33 { get; set; }
        public double Voltage55 { get; set; }
        public double Status { get; set; }

        public double Mag_X { get; set; }
        public double Mag_Y { get; set; }
        public double Mag_Z { get; set; }
        public double Mag { get; set; }

        public MyStruct(double time, double pressure, double accelaration, double rotation)
        {
            TimeStampSF = time;
            Pressure = pressure;
            Acceleration = accelaration;
            Rotation = rotation;
        }

        public MyStruct(double time, double pressure, double accX, double accY, double accZ, double accelaration, double rotX, double rotY, double rotZ, double rotation,
            double temperature, double vBattery, double voltage33, double voltage55, double status, double magX, double magY, double magZ, double mag)
        {
            TimeStampSF = time;
            Pressure = pressure;
            AccelerationX = accX;
            AccelerationY = accY;
            AccelerationZ = accZ;
            Acceleration = accelaration;
            RotationX = rotX;
            RotationY = rotY;
            RotationZ = rotZ;
            Rotation = rotation;
            Temperature = temperature;
            VBattery = vBattery;
            Voltage33 = voltage33;
            Voltage55 = voltage55;
            Status = status;
            Mag_X = magX;
            Mag_Y = magY;
            Mag_Z = magZ;
            Mag = mag;
        }

        public MyStruct(double time, double pressure, double accX, double accY, double accZ, double accelaration, double rotX, double rotY, double rotZ, double rotation,
            double status)
        {
            TimeStampSF = time;
            Pressure = pressure;
            AccelerationX = accX;
            AccelerationY = accY;
            AccelerationZ = accZ;
            Acceleration = accelaration;
            RotationX = rotX;
            RotationY = rotY;
            RotationZ = rotZ;
            Rotation = rotation;
            Status = status;
        }

        public MyStruct(double time, double pressure, double accX, double accY, double accZ, double accelaration, double rotX, double rotY, double rotZ, double rotation)
        {
            TimeStampSF = time;
            Pressure = pressure;
            AccelerationX = accX;
            AccelerationY = accY;
            AccelerationZ = accZ;
            Acceleration = accelaration;
            RotationX = rotX;
            RotationY = rotY;
            RotationZ = rotZ;
            Rotation = rotation;
        }
    }
}
