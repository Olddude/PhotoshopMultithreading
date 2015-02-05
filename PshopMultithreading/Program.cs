using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Photoshop;

namespace PshopMultithreading
{
    class Program
    {
        public static Application App = new Application();
        public static Thread PhotoshopThread = null;

        static void Main(string[] args)
        {
            Document document = App.Documents.Add(500, 500, 96, "MyDocument") as Document;
            document.Info.Keywords = new object[] { "editing" };

            Start(ref App);

            Console.WriteLine("You have finished everything! You are a great man!!!");
            Console.ReadKey();
        }

        static void Start(ref Application appHandler)
        {
            var appHandler2 = appHandler;
            PhotoshopThread = new Thread(delegate() { PhotoshopJob(ref appHandler2); });
            PhotoshopThread.Start();
            PhotoshopThread.Join();   // Uncomment this line to ensure HandleClientComm has finished running.
            appHandler = appHandler2; // handler2 may or may not have changed by now.
        }

        static void PhotoshopJob(ref Application photoshop)
        {
            try
            {
                string status = "";
                while(status != "done")
                {
                        Thread.Sleep(1000);
                        status = photoshop.ActiveDocument.Info.Keywords[0].ToString();
                        Console.WriteLine(status);
                }

                Console.WriteLine(photoshop.Documents.Count.ToString() + " " + status);

                JPEGSaveOptions saveOptions = new JPEGSaveOptions();
                saveOptions.EmbedColorProfile = true;
                saveOptions.FormatOptions = PsFormatOptionsType.psStandardBaseline;
                saveOptions.Matte = PsMatteType.psNoMatte;
                saveOptions.Quality = 12;

                photoshop.ActiveDocument.SaveAs(@"C:\Users\Dude\Documents\00.jpg", saveOptions, true, PsExtensionType.psLowercase);

                Thread.CurrentThread.Abort();
            }
            catch
            {
                if (PhotoshopThread.ThreadState != ThreadState.AbortRequested)
                {
                    Console.WriteLine(PhotoshopThread.ThreadState.ToString());
                    Start(ref photoshop);
                }
            }
        }
    }
}
