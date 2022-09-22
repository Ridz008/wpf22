using System.IO.Ports;

//Console.WriteLine(SerialPort.GetPortNames().ToArray());
// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");
SerialPort port = new SerialPort("COM1", 9000, Parity.None, 8, StopBits.One);

port.Open();
port.WriteLine("Hello");
port.Close(); 
//new SerialPortDemo();
Console.ReadLine();
port.WriteLine("Hello");
public class SerialPortDemo
{
    //SerialPort port = new SerialPort("COM6", 9000, Parity.None, 8, StopBits.One);
    public SerialPortDemo()
    {
        //Console.WriteLine("Incoming data");
        //port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
        //port.Open();
        //port.WriteLine("sending from COM6");
        //port.Write("");
    }
    public void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        //Console.WriteLine(port.ReadExisting());
    }
}