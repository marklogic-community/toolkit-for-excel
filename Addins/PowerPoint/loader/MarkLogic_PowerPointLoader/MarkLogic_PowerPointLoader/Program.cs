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
        private static void uploadData(string url, byte[] content, string user, string pwd)
        {
            string message = "";

            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Headers.Add("enctype", "multipart/form-data");
                Client.Headers.Add("Content-Type", "application/octet-stream");
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);

                Client.UploadData(url, "POST", content);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }


            // return message;

        }

        public static string savePPTXToML(string filefullname,string url, string user, string pwd)
        {
            
            string message = "";
            try
            {
                string fname = filefullname.Split(new Char[] { '\\' }).Last();
                string fileuri =  "/" + fname;
                string uri = url + fileuri;

                FileStream fs = new FileStream(filefullname, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                int length = (int)fs.Length;
                byte[] content = new byte[length];
                fs.Read(content, 0, length);

                try
                {
                    uploadData(uri, content, user, pwd);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    // MessageBox.Show("message1 :" + message);
                }

                fs.Dispose();
                fs.Close();
                Console.WriteLine("-----------------------message" + message);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
            }

            return message;
        }

 
        public static string saveImagesToML(string outputdir, string saveasdir, string url, string user, string pwd)
        {
            string message = "";
            string[] imgfiles = Directory.GetFiles(outputdir);

           // string user = "oslo";
           // string pwd = "oslo";

            foreach (string i in imgfiles)
            {
               // MessageBox.Show("filename: " + i);
                string fname = i.Split(new Char[] { '\\' }).Last();
                string fileuri = saveasdir + "/" + fname; // "/"+outputdir + "/" + fname;
                //convert this uri to .pptx slide.xml
                //MessageBox.Show("i"+i);
                //MessageBox.Show("FileUri"+fileuri);
                //als get index from here
                // add as parameters for upload.xqy doc properties

                string uri = url + fileuri;
                //MessageBox.Show("url"+url);
                try
                {
                   
                    FileStream fs = new FileStream(i, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);

                    try
                    {
                        uploadData(uri, content,user,pwd);
                    }
                    catch (Exception e)
                    {
                        string errorMsg = e.Message;
                        message = "error: " + errorMsg;
                        
                    }
                    
                    fs.Dispose();
                    fs.Close();
                   
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                   
                }
            }
            return message;
        }
     
    
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

            if (args.Length != 4)
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
            string url = "@"+args[1];
            string user = args[2];
            string pwd = args[3];


            object missing = System.Type.Missing;
           // string user ="oslo";
           // string pwd ="oslo";
            //string url = "http://localhost:8023/ppt/api/upload.xqy?uid="; 
            DirectoryInfo root = new DirectoryInfo(sourceDirectory);
            //DirectoryInfo root = new DirectoryInfo(@"C:\Documents and Settings\paven\Desktop\presentationsToConvert\");
            DirectoryInfo[] dirs = root.GetDirectories("*", SearchOption.AllDirectories);

            foreach(DirectoryInfo d in dirs)
            {
                if (d.Name.Equals("ug2009"))//only for testing purposes
                {
                    Console.WriteLine(d.FullName);
                    FileInfo[] files = d.GetFiles();
                    Console.WriteLine("files:");


                    Application ppt = new Application();
                    foreach (FileInfo file in files)
                    {  
                        Presentation pres = ppt.Presentations.Open(file.FullName, MsoTriState.msoFalse, MsoTriState.msoTrue,MsoTriState.msoFalse);
                        string imgdirwithpath = "";
                        string saveasdir = "";
                        if (file.Name.EndsWith(".ppt"))
                        {
                            imgdirwithpath = file.FullName.Replace(".ppt", "_PNG");
                            saveasdir = "/"+file.Name.Replace(".ppt", "_PNG");
                        }
                        else
                        {
                            imgdirwithpath = file.FullName.Replace(".pptx", "_PNG");
                            saveasdir = "/"+file.Name.Replace(".pptx", "_PNG");
                        }

                        pres.SaveAs(imgdirwithpath, PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoFalse);
                        pres.Close();

                        saveImagesToML(imgdirwithpath, saveasdir, url, user,pwd);
                        savePPTXToML(file.FullName,url,user,pwd);
                     
                        Console.WriteLine(file.FullName + "|" + file.Directory + "|" + file.Name);
                        
                    }
                    ppt.Quit();
                   
                    
                }
                
            }
         
            Console.WriteLine("Press any key too continue...");
            Console.ReadLine();
        }
    }
}
