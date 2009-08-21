using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
//using System.IO.Compression;
//using System.Net;
//using System.Net.Mail;
//using System.Net.Sockets;
//using System.Runtime.InteropServices;


namespace MarkLogic_PowerPointLoader
{
    class Program
    {

        static void Main(string[] args)
        {
            //parameters reqd
            //path of dir with ppt(x)s  convert, load to ML
            //path of output dir for image dirs ? (right now just create dir of images sibling to .pptx)
            //url to ML for saving the ppt and images 
            //user
            //pwd
            //could add- folder to save in ML, right now defaults to '/'

            //this assumes pipeline already installed 

            if (args.Length != 1)
            {
                Console.WriteLine(args.Length);
                Console.WriteLine(args);
                foreach (string s in args)
                {
                    Console.WriteLine("arg: " + s);
                }
                Console.WriteLine("Expected Params: (tbd) ....\nPress any key too continue...");
                Console.ReadLine();
                return;
            }

            string sourceDirectory = args[0];


            object missing = System.Type.Missing;
            DirectoryInfo root = new DirectoryInfo(sourceDirectory);
            DirectoryInfo[] dirs = root.GetDirectories("*", SearchOption.AllDirectories);

            foreach (DirectoryInfo d in dirs)
            {
                //if (d.Name.Equals("ug2009"))//only for testing purposes
                //{
                    Console.WriteLine(d.FullName);
                    FileInfo[] files = d.GetFiles();
                    Console.WriteLine("files:");


                    Application ppt = new Application();
                    foreach (FileInfo file in files)
                    {
                        Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                        string imgdirwithpath = "";
                        string saveasdir = "";
                        if (file.Name.EndsWith(".pptx"))
                        {
                            imgdirwithpath = file.FullName.Replace(".pptx", "_PNG");
                            saveasdir = "/" + file.Name.Replace(".ppt", "_PNG");
                            //}
                            //else
                            //{
                            //    imgdirwithpath = file.FullName.Replace(".pptx", "_PNG");
                            //    saveasdir = "/"+file.Name.Replace(".pptx", "_PNG");
                            //}
                            Console.WriteLine("BEFORE:   " + file.FullName + "|" + file.Directory + "|" + file.Name);
                            pres.SaveAs(imgdirwithpath, PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoFalse);
                            pres.Close();
                        }

                        Console.WriteLine("AFTER :" + file.FullName + "|" + file.Directory + "|" + file.Name);

                    }
                    ppt.Quit();


                //}

            }

            Console.WriteLine("Press any key too continue...");
            Console.ReadLine();
        }
    }
}
