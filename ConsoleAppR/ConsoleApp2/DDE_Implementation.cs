using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using DDEML;

//namespace ConsoleApp2
//{

//    internal class DDE_Implementation
//    {


//    }
//}

// Matlab Interface Library
// by Emanuele Ruffaldi 2002
// http://www.sssup.it/~pit/
// mailto:pit@sssup.it
//
// Description: DDE Client Library (string and XLTable transfers)

/*
 * This is a DDE Library for C# that uses DDEML to implement a DDE Client
* I use this library to access MATLAB DDE
*
* Classes:
*  DDE	(P/Invoke wrapper and base service class for DDEClient and future DDEServer)
*  DDEClient  (client)
*  DDEChannel (open channel with an application)
*  DDEItem    (a text for efficient multicall)
* 
* Implemented: poke request execute
* Not implemented because not interesting: adv and unadv
* 
* TODO: BUSY support
* TODO: METAFILE support
*/
using System;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;

namespace DDELibrary
{
    /// <summary>
    /// The type of transaction
    /// </summary>
    public enum TransactionType
    {
        /// <summary>
        /// register
        /// </summary>
        XTYP_REGISTER = 0x00A0 | 0x8000 | 0x0002,
        /// <summary>
        /// unregister
        /// </summary>
        XTYP_UNREGISTER = 0x00D0 | 0x8000 | 0x0002,
        /// <summary>
        /// advise data
        /// </summary>
        XTYP_ADVDATA = 0x0010 | 0x4000,
        /// <summary>
        /// xact complete
        /// </summary>
        XTYP_XACT_COMPLETE = 0x0080 | 0x8000,
        /// <summary>
        /// disconnection
        /// </summary>
        XTYP_DISCONNECT = 0x00C0 | 0x8000 | 0x0002,
        /// <summary>
        /// execute
        /// </summary>
        XTYP_EXECUTE = 0x0050 | 0x4000,
        /// <summary>
        /// request
        /// </summary>
        XTYP_REQUEST = 0x00B0 | 0x2000,
        /// <summary>
        /// poke
        /// </summary>
        XTYP_POKE = 0x0090 | 0x4000
    }

    /// <summary>
    /// Result status from calls
    /// </summary>
    internal enum DDEReturnStatus
    {
        // DDE_
        FBUSY = 0x4000,
        FACK = 0x8000,
        FNOTPROCESSED = 0x0000
    }

    /// <summary>
    /// Errors
    /// </summary>
    public enum DDEErrors
    {
        // DMLERR_
        NOTPROCESSED = 0x4009,
        NOERROR = 0,
        BUSY = 0x4001,
        EXECACKTIMEOUT = 0x4005,
        POKEACKTIMEOUT = 0x400b,
        DATAACKTIMEOUT = 0x4002
    }

    /// <summary>
    /// Some usefule Clipboard Formats
    /// </summary>
    internal enum ClipboardFormats
    {
        CF_NONE = 0,
        CF_TEXT = 1,
        CF_BITMAP = 2,
        CF_METAFILEPICT = 3,
        CF_UNICODETEXT = 13

    }

    /// <summary>
    /// A class to wrap DDE services. It uses DDEML (a standard Windows component)
    /// to implement the DDE services. It's the base class for the DDEClient
    /// and for a (possible) DDEServer
    /// </summary>
    public class DDE : IDisposable
    {

        #region DDE P/Invoke
        #region Main
        //---------- DDE LIBRARY
        // initialize the system using unicode
        [DllImport("user32.dll")]
        protected static extern int DdeInitializeW(ref int id, DDECallback cb, int afcmd, int ulres);

        // uninitialize
        [DllImport("user32.dll")]
        static extern int DdeUninitialize(int id);
        #endregion
        #region Channel

        [DllImport("user32.dll")]
        internal static extern DDEErrors DdeGetLastError(int idInst);

        //---------- CHANNELS

        // create a channel connection to a service:topic
        [DllImport("user32.dll")]
        static extern IntPtr DdeConnect(
              int idInst,             // instance identifier
            IntPtr hszService,      // handle to service name string
            IntPtr hszTopic,        // handle to topic name string
            IntPtr pCC              // context data
            );

        // disconnects the channel
        [DllImport("user32.dll")]
        internal static extern int DdeDisconnect(
              IntPtr hc);

        [DllImport("user32.dll")]
        internal static extern IntPtr DdeClientTransaction(
            IntPtr pData,       // pointer to data to pass to server
          int cbData,       // length of data
          IntPtr hConv,        // handle to conversation
          IntPtr hszItem,        // handle to item name string
          ClipboardFormats wFmt,          // clipboard data format
          TransactionType wType,         // transaction type
          int dwTimeout,    // time-out duration
          int pdwResult   // pointer to transaction result
          );

        // start a transaction using a string as parameter. Should exchange the charset to CF_TEXT (ANSI)
        [DllImport("user32.dll", EntryPoint = "DdeClientTransaction", CharSet = CharSet.Auto)]
        internal static extern IntPtr DdeClientTransactionString(
              string pData,       // pointer to data to pass to server
            int cbData,       // length of data
            IntPtr hConv,        // handle to conversation
            IntPtr hszItem,        // handle to item name string
            ClipboardFormats wFmt,          // clipboard data format
            TransactionType wType,         // transaction type
            int dwTimeout,    // time-out duration
            int pdwResult   // pointer to transaction result
            );

        // start a transaction using a string as parameter. Should exchange the charset to CF_TEXT (ANSI)
        [DllImport("user32.dll", EntryPoint = "DdeClientTransaction", CharSet = CharSet.Ansi)]
        internal static extern IntPtr DdeClientTransactionStringA(
              string pData,       // pointer to data to pass to server
            int cbData,       // length of data
            IntPtr hConv,        // handle to conversation
            IntPtr hszItem,        // handle to item name string
            ClipboardFormats wFmt,          // clipboard data format
            TransactionType wType,         // transaction type
            int dwTimeout,    // time-out duration
            int pdwResult   // pointer to transaction result
            );

        // start a transaction using a byte [] as parameter. Should exchange the charset to CF_TEXT (ANSI)
        [DllImport("user32.dll", EntryPoint = "DdeClientTransaction", CharSet = CharSet.Ansi)]
        internal static extern IntPtr DdeClientTransactionXL(
            byte[] data,                // pointer to data to pass to server
            int cbData,                 // length of data
            IntPtr hConv,               // handle to conversation
            IntPtr hszItem,             // handle to item name string
            int wFmt,                   // clipboard data format
            TransactionType wType,         // transaction type
            int dwTimeout,          // time-out duration
            int pdwResult           // pointer to transaction result
            );

        #endregion
        #region Data Handle
        //---------- DATAHANDLE

        // returns the data pointer associated with the handle and its size
        [DllImport("user32.dll")]
        internal static extern IntPtr DdeAccessData(IntPtr p, out int datasize);

        // same as above but use P/Invoke to get the data as an ANSI string (because of CF_TEXT)
        [DllImport("user32.dll", EntryPoint = "DdeAccessData", CharSet = CharSet.Ansi)]
        internal static extern string DdeAccessDataString(IntPtr p, out int datasize);

        // free a data handle
        [DllImport("user32.dll")]
        internal static extern void DdeFreeDataHandle(IntPtr data);

        // create a data handle from string
        [DllImport("user32.dll", CharSet = CharSet.Ansi, EntryPoint = "DdeCreateDataHandle")]
        internal static extern IntPtr DdeCreateDataHandleString(int id, string data, int len, int off, IntPtr hszitem,
              ClipboardFormats wFmt, int flags);

        [DllImport("user32.dll")]
        internal static extern IntPtr DdeCreateDataHandle(int id, IntPtr data, int len, int off, IntPtr hszitem,
            ClipboardFormats wFmt, int flags);

        #endregion
        #region Strings
        //---------- STRINGS

        // creates a string handle from a string
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr DdeCreateStringHandleW(
              int idInst,  // instance identifier
            string psz,    // pointer to null-terminated string
            int iCodePage  // code page identifier 1200
            );

        // destroy a string handle
        [DllImport("user32.dll")]
        static extern int DdeFreeStringHandle(
              int idInst,
              IntPtr hsz
          );

        // fill in a string from the string
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern int DdeQueryString(int idInst,
              IntPtr hsz,
              StringBuilder text,
              int length,
              int cp);  // 1200 because of unicode		
        #endregion

        protected internal delegate IntPtr DDECallback(
              TransactionType uType,
              int uFmt,
              IntPtr hConv,
              IntPtr hsz1,
              IntPtr hsz2,
              IntPtr hdata,
              IntPtr data1,
              IntPtr data2
              );

        #endregion

        [DllImport("user32.dll")]
        internal static extern int RegisterClipboardFormat(string lpszFormat);

        /// <summary>
        /// Constructor of the DDE Instance. Usually only one of this class per application 		
        /// </summary>
        public DDE()
        {
            instanceID = 0;
        }

        /// <summary>
        /// Closes the Service
        /// </summary>
        public void Close()
        {
            Dispose();
        }

        /// <summary>
        /// IDisposable implementation
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Tells if the connection is active
        /// </summary>
        public bool Active
        {
            get { return instanceID != 0; }
        }

        /// <summary>
        /// Opens a Channel of communication with the specified service and the specific topic
        /// </summary>
        /// <param name="service">the name of the service</param>
        /// <param name="topic">the name of topic</param>
        /// <returns></returns>
        public DDEChannel OpenChannel(string service, string topic)
        {
            if (!Active) return null;

            DDEItem diService = new DDEItem(CreateString(service));
            DDEItem diTopic = new DDEItem(CreateString(topic));
            DDEChannel c = OpenChannel(diService, diTopic);
            DestroyString(diService.Item);
            DestroyString(diTopic.Item);
            return c;
        }

        /// <summary>
        /// Opens a Communication Channel with the Service and Topic
        /// Use DDEItem objects to reduce memory overhead
        /// </summary>
        /// <param name="hService"></param>
        /// <param name="hTopic"></param>
        /// <returns></returns>
        public DDEChannel OpenChannel(DDEItem hService, DDEItem hTopic)
        {
            if (!Active) return null;
            IntPtr channel = DdeConnect(instanceID, hService.Item, hTopic.Item, new IntPtr(0));
            if (channel == IntPtr.Zero)
                return null;
            else
                return new DDEChannel(this, channel);
        }

        static internal string QueryStringFromDataHandle(IntPtr h)
        {
            int datasize;
            return DdeAccessDataString(h, out datasize);
        }

        internal IntPtr CreateDataHandleFromString(string s)
        {
            return DdeCreateDataHandleString(instanceID, s, s.Length, 0, IntPtr.Zero, ClipboardFormats.CF_TEXT, 0);
        }

        internal IntPtr CreateDataHandleFromString(string s, string item)
        {
            IntPtr hitem = CreateString(item);
            return DdeCreateDataHandleString(instanceID, s, s.Length, 0, hitem, ClipboardFormats.CF_TEXT, 0);
        }

        // String Access
        protected internal DDEItem CreateStringItem(string n)
        {
            return new DDEItem(CreateString(n));
        }

        protected internal void DestroyStringItem(DDEItem di)
        {
            DestroyString(di.Item);
        }

        internal IntPtr CreateString(string n)
        {
            return instanceID == 0 ? IntPtr.Zero : DdeCreateStringHandleW(instanceID, n, 1200);
        }

        internal void DestroyString(IntPtr hs)
        {
            if (instanceID != 0)
                DdeFreeStringHandle(instanceID, hs);
        }

        internal string QueryString(IntPtr hs)
        {
            int n = DdeQueryString(instanceID, hs, null, 0, 1200);
            StringBuilder q = new StringBuilder(n + 1);
            DdeQueryString(instanceID, hs, q, q.Capacity, 1200);
            return q.ToString();
        }

        /// <summary>
        /// Returns the last error
        /// </summary>
        /// <returns></returns>
        public DDEErrors GetLastError()
        {
            return DdeGetLastError(instanceID);
        }

        /// <summary>
        /// A Testing function
        /// </summary>
        public void Test()
        {
            // STRING TEST
            IntPtr pp = CreateString("Hello World");
            if (pp != IntPtr.Zero)
            {
                Console.WriteLine("(HSTRING) Hello World = " + QueryString(pp));
                DestroyString(pp);
            }

            // DATA HANDLE TEST AS STRINGS
            IntPtr ppd = CreateDataHandleFromString("Hello World");
            if (ppd != IntPtr.Zero)
            {
                string s = QueryStringFromDataHandle(ppd);
                Console.WriteLine("(DATAHANDLE) Hello World = " + s);
                DdeFreeDataHandle(ppd);
            }
        }

        /// <summary>
        /// IDisposable implementation
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (instanceID != 0)
            {
                DdeUninitialize(instanceID);
                instanceID = 0;
            }
        }

        ~DDE()
        {
            Dispose(false);
        }

        /// <summary>
        /// The instance variable associated with the DDE session
        /// </summary>
        protected int instanceID;
    }

    /// <summary>
    /// A DDE Client class
    /// </summary>
    public class DDEClient : DDE
    {
        /// <summary>
        /// Initialize the system
        /// </summary>
        public DDEClient()
        {
            // 0x00000010L means Client Only
            DdeInitializeW(ref instanceID, new DDECallback(this.DDECallbackX), 0x00008000 | 0x003c0000 | 0x00000010, 0);
        }

        /// <summary>
        /// DDECallback for a Client ... do Nothing
        /// </summary>
        /// <param name="uType"></param>
        /// <param name="uFmt"></param>
        /// <param name="hConv"></param>
        /// <param name="hsz1"></param>
        /// <param name="hsz2"></param>
        /// <param name="hdata"></param>
        /// <param name="data1"></param>
        /// <param name="data2"></param>
        /// <returns></returns>
        internal IntPtr DDECallbackX(
              TransactionType uType,
              int uFmt,
              IntPtr hConv,
              IntPtr hsz1,
              IntPtr hsz2,
              IntPtr hdata,
              IntPtr data1,
              IntPtr data2
              )
        {
            return IntPtr.Zero;
        }
    }

    /// <summary>
    /// efficient DDE item storage
    /// BE CAREFULE: it's not Dispose protected!!!
    /// </summary>
    public struct DDEItem
    {
        internal DDEItem(IntPtr p)
        {
            hs = p;
        }

        internal IntPtr Item
        {
            get { return hs; }
        }

        IntPtr hs;
    }

    /// <summary>
    /// A DDE Channel handles the connection between a Client and a DDE Server
    /// </summary>
    public class DDEChannel : IDisposable
    {
        /// <summary>
        /// Created by a DDE object
        /// </summary>
        /// <param name="d">the parent dde object</param>
        /// <param name="c">the channel handler</param>
        internal DDEChannel(DDE d, IntPtr c)
        {
            parent = d;
            channel = c;

            if (xlFormat == 0)
            {
                xlFormat = DDE.RegisterClipboardFormat("XlTable");
            }
        }

        ~DDEChannel()
        {
            Dispose(false);
        }

        /// <summary>
        /// IDisposable requirement
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disconnect the Channel
        /// </summary>
        public void Disconnect()
        {
            Dispose();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (channel != IntPtr.Zero)
            {
                DDE.DdeDisconnect(channel);
                channel = IntPtr.Zero;
            }
        }

        /// <summary>
        /// Tells if Active
        /// </summary>
        public bool Active
        {
            get { return channel.ToInt32() != 0; }
        }

        /// <summary>
        /// Executes a DDE Command on the specified item using the
        /// default timeout
        /// </summary>
        /// <param name="command"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Execute(string command, string item)
        {
            return Execute(command, item, defTimeout);
        }

        /// <summary>
        /// Executes a DDE Command on the specified item using
        /// a timeout
        /// </summary>
        /// <param name="command"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool Execute(string command, int timeout)
        {
            return Execute(command, new DDEItem(IntPtr.Zero), timeout);
        }

        /// <summary>
        /// Executes a DDE Command without an item and with the default
        /// timeout
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        public bool Execute(string command)
        {
            return Execute(command, defTimeout);
        }

        /// <summary>
        /// Executes a DDE command on item specifying the timeout
        /// </summary>
        /// <param name="command"></param>
        /// <param name="item"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool Execute(string command, string item, int timeout)
        {
            if (!Active)
                return false;
            DDEItem di = new DDEItem(parent.CreateString(item));
            bool b = Execute(command, di, timeout);
            parent.DestroyString(di.Item);
            return b;
        }

        /// <summary>
        /// Same as the other overloaded classes but the item is specified
        /// by a DDEItem to make it faster
        /// </summary>
        /// <param name="command"></param>
        /// <param name="item"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool Execute(string command, DDEItem item, int timeout)
        {
            if (!Active) return false;
            lastTimeouted = false;
            do
            {
                IntPtr da = DDE.DdeClientTransactionString(command, command.Length * 2 + 2, channel, item.Item, 0, TransactionType.XTYP_EXECUTE, timeout, 0);
                if (da == IntPtr.Zero)
                {
                    DDEErrors e = parent.GetLastError();
                    if (e == DDEErrors.BUSY)
                    {
                        Thread.Sleep(100);
                        continue;
                    }
                    else if (e == DDEErrors.EXECACKTIMEOUT)
                    {
                        lastTimeouted = true;
                        return false;
                    }
                    else
                        return false;
                }
                else
                {
                    DDE.DdeFreeDataHandle(da);
                    return true;
                }
            } while (true);
        }

        /// <summary>
        /// Request a specific item from the DDE Server
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public bool Request(string item, out string result)
        {
            return Request(item, out result, defTimeout);
        }

        /// <summary>
        /// Request from the DDE Server using a timeout
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool Request(string item, out string result, int timeout)
        {
            result = null;
            if (!Active)
                return false;
            DDEItem di = new DDEItem(parent.CreateString(item));
            bool b = Request(di, out result, timeout);
            parent.DestroyString(di.Item);
            return b;
        }

        /// <summary>
        /// Request of an item from the DDE Server in XLTable format
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool RequestXL(DDEItem item, out byte[] result, int timeout)
        {
            result = null;
            lastTimeouted = false;
            if (!Active)
                return false;
            do
            {
                IntPtr da = DDE.DdeClientTransactionXL(null, 0, channel, item.Item, xlFormat, TransactionType.XTYP_REQUEST, timeout, 0);
                if (da == IntPtr.Zero)
                {
                    DDEErrors e = parent.GetLastError();
                    if (e == DDEErrors.BUSY)
                    {
                        Thread.Sleep(100);
                        continue;
                    }
                    else if (e == DDEErrors.DATAACKTIMEOUT)
                    {
                        lastTimeouted = true;
                        return false;
                    }
                    else
                        return false;
                }
                else
                {
                    int datasize;
                    IntPtr p = DDE.DdeAccessData(da, out datasize);
                    if (p != IntPtr.Zero)
                    {
                        result = new byte[datasize];
                        Marshal.Copy(p, result, 0, datasize);
                    }
                    DDE.DdeFreeDataHandle(da);
                    return true;
                }
            } while (true);
        }

        /// <summary>
        /// Faster request of an item from the server using a DDEItem
        /// </summary>
        /// <param name="item"></param>
        /// <param name="result"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool Request(DDEItem item, out string result, int timeout)
        {
            result = null;
            lastTimeouted = false;
            if (!Active)
                return false;
            do
            {
                IntPtr da = DDE.DdeClientTransaction(IntPtr.Zero, 0, channel, item.Item, ClipboardFormats.CF_TEXT, TransactionType.XTYP_REQUEST, timeout, 0);
                if (da == IntPtr.Zero)
                {
                    DDEErrors e = parent.GetLastError();
                    if (e == DDEErrors.BUSY)
                    {
                        Thread.Sleep(100);
                        continue;
                    }
                    else if (e == DDEErrors.DATAACKTIMEOUT)
                    {
                        lastTimeouted = true;
                        return false;
                    }
                    else
                        return false;
                }
                else
                {
                    result = DDE.QueryStringFromDataHandle(da);
                    DDE.DdeFreeDataHandle(da);
                    return true;
                }
            } while (true);
        }

        /// <summary>
        /// Writes (poke) a value into the specific item
        /// </summary>
        /// <param name="item"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool Poke(string item, string data)
        {
            return Poke(item, data, defTimeout);
        }

        /// <summary>
        /// Writes (poke) a value into the specific item using a timeout
        /// </summary>
        /// <param name="item"></param>
        /// <param name="data"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool Poke(string item, string data, int timeout)
        {
            if (!Active) return false;
            DDEItem di = new DDEItem(parent.CreateString(item));
            bool b = Poke(di, data, timeout);
            parent.DestroyString(di.Item);
            return b;
        }

        /// <summary>
        /// Writes (poke) a value into an item using XLTable format
        /// </summary>
        /// <param name="item"></param>
        /// <param name="data"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool PokeXL(DDEItem item, byte[] data, int timeout)
        {
            lastTimeouted = false;
            if (!Active)
                return false;
            do
            {
                IntPtr da = DDE.DdeClientTransactionXL(data, data.Length, channel, item.Item, xlFormat, TransactionType.XTYP_POKE, timeout, 0);
                if (da == IntPtr.Zero)
                {
                    DDEErrors e = parent.GetLastError();
                    if (e == DDEErrors.BUSY)
                    {
                        Thread.Sleep(100);
                        continue;
                    }
                    else if (e == DDEErrors.POKEACKTIMEOUT)
                    {
                        lastTimeouted = true;
                        return false;
                    }
                    else
                        return false;
                }
                else
                {
                    DDE.DdeFreeDataHandle(da);
                    return true;
                }
            } while (true);
        }

        /// <summary>
        /// Faster Write (poke) using a DDEItem
        /// </summary>
        /// <param name="item"></param>
        /// <param name="data"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public bool Poke(DDEItem item, string data, int timeout)
        {
            lastTimeouted = false;
            if (!Active)
                return false;
            do
            {
                IntPtr da = DDE.DdeClientTransactionString(data, (data.Length + 1) * 2, channel, item.Item, ClipboardFormats.CF_TEXT, TransactionType.XTYP_POKE, timeout, 0);
                if (da == IntPtr.Zero)
                {
                    DDEErrors e = parent.GetLastError();
                    if (e == DDEErrors.BUSY)
                    {
                        Thread.Sleep(100);
                        continue;
                    }
                    else if (e == DDEErrors.POKEACKTIMEOUT)
                    {
                        lastTimeouted = true;
                        return false;
                    }
                    else
                        return false;
                }
                else
                {
                    DDE.DdeFreeDataHandle(da);
                    return true;
                }
            } while (true);
        }

        /// <summary>
        /// The default timeout for Execute, Request and Poke calls
        /// </summary>
        public int Timeout
        {
            get { return defTimeout; }
            set { defTimeout = value; }
        }

        /// <summary>
        /// Tells if the last operation timeouted
        /// </summary>
        public bool LastTimeouted
        {
            get { return lastTimeouted; }
        }

        /// <summary>
        /// The parent DDE session
        /// </summary>
        DDE parent;
        /// <summary>
        /// The channel with the DDE Server
        /// </summary>
        IntPtr channel;

        int defTimeout = 3000;
        bool lastTimeouted;

        static int xlFormat = 0;
    }

}