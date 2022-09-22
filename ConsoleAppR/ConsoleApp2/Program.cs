using System.IO.Ports;


//Console.WriteLine(SerialPort.GetPortNames().ToArray());
// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");
new SerialPortDemo();
Console.ReadLine();
public class SerialPortDemo
{
    
     SerialPort port = new SerialPort("COM4", 9000, Parity.None, 8, StopBits.One);
    public SerialPortDemo()
    {
        Console.WriteLine("Incoming data");
        
        try
        {
            
            port.Open();
            //port.Write("Y"); 
            //port.WriteLine("ss");
            port.Close();   
            
            //Console.WriteLine(port.ReadLine());
        }
        catch (Exception ex)
        {
            Console.WriteLine("exception " + ex.ToString());
        }
        //port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
        
    } 
    public void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        Console.WriteLine(port.ReadExisting());
        //Console.Write(port.ReadLine());
    }
}