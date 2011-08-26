using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;

using Marklogic.Xcc;

namespace DroidPublish
{
    class GalleryFactory
    {
        public void publishGallery(String galleryURI)
        {
            Uri serverUri = new Uri("xcc://oslo:oslo@localhost:8032");
            //MessageBox.Show("URI: " + galleryURI);
            String galleryFetcherUri = "publish-fetcher.xqy"; //args[1];

            string[] tempSourcePres = galleryURI.Split('/');

            //update this to get the name of the file from the gallery.xml
            string fileName = tempSourcePres[tempSourcePres.Length-1].Replace(".xml",".pptx");


           // MessageBox.Show("Filename: " + fileName);
            ModuleRunner fetchRunner = new ModuleRunner(serverUri);
            //String result = runner.InvokeToSingleString(moduleUri, "\n");
            fetchRunner.Request.SetNewStringVariable("doc", galleryURI);
            String[] result = fetchRunner.InvokeToStringArray(galleryFetcherUri);

            int length = result.Length;
            ContentFetcher fetcher = new ContentFetcher(serverUri);
            Stream tmpDestDocStream = null;
            Stream tmpSrcDocStream = null;
            MemoryStream destDocStream = new MemoryStream();
            MemoryStream final = new MemoryStream();
          
            
                for (int i = 0; i < length; i++)
                {
                    MemoryStream srcDocStream = new MemoryStream();
                   // MessageBox.Show("FILE NAME: " + result[i]);
                    String docUri = result[i];
                    ResultItem binDoc = fetcher.Fetch(result[i]);


                    if (i == 0)
                    {
                        tmpDestDocStream = binDoc.AsInputStream();
                        destDocStream = MergeDecks.CopyToMemory(tmpDestDocStream);
                        final = destDocStream;
                    }
                    else
                    {
                        tmpSrcDocStream = binDoc.AsInputStream();
                        srcDocStream = MergeDecks.CopyToMemory(tmpSrcDocStream);
                        final = MergeDecks.Assemble(srcDocStream, final);
                        
                    }

                    srcDocStream.Close();

                } //end of for

            //TODO:have to insure these directories exist
            //create if not present under TMP

                string tmpPath = System.IO.Path.GetTempPath() + @"publish\";
                string deletePath = System.IO.Path.GetTempPath() + @"delete\";

                string tmpSourcePres = tmpPath + fileName;
                string deleteSourcePres = deletePath + fileName;

                //Stream outStream1 = new FileStream(@"C:\Users\paven\Desktop\test\output.pptx", FileMode.Create);

                 Stream outStream1 = new FileStream(tmpSourcePres, FileMode.Create);
                 final.WriteTo(outStream1);
                 //final.Seek(0, SeekOrigin.Begin); 


                 FileInfo[] sourceFile = new FileInfo[1];
                 sourceFile[0] = new FileInfo(tmpSourcePres);

                 string[] sourceURI = { "/out/"+fileName };

                ContentLoader loader = new ContentLoader(serverUri);
                loader.Load(sourceURI, sourceFile);

                long start = DateTime.Now.Ticks;
          
                destDocStream.Close();
                tmpDestDocStream.Close();
                tmpSrcDocStream.Close();
                final.Close();
                outStream1.Close();

                if (File.Exists(deleteSourcePres))
                     File.Delete(deleteSourcePres);

                File.Move(tmpSourcePres, deleteSourcePres);

         
        }

        public void deleteGallery(string galleryURI)
        {
            try
            {
                Uri serverUri = new Uri("xcc://oslo:oslo@localhost:8032");
                String galleryFetcherUri = "publish-delete.xqy"; //args[1];
                ModuleRunner fetchRunner = new ModuleRunner(serverUri);
                fetchRunner.Request.SetNewStringVariable("doc", galleryURI);
                String[] result = fetchRunner.InvokeToStringArray(galleryFetcherUri);
            }
            catch (Exception e)
            {
                MessageBox.Show("ERROR" + e.Message);
            }
        }

    }
}
