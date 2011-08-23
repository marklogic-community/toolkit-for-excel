using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Xml;
using System.Windows.Forms;
using Marklogic.Xcc;


namespace DroidPublish
{
    class MergeDecks
    {
        static uint uniqueId;
        public static MemoryStream Assemble(MemoryStream sourceDoc, MemoryStream destDoc)
        {
            int id = 1;
            //Open up the destination deck

            using (PresentationDocument myDestDeck = PresentationDocument.Open(destDoc, true))
            {
                PresentationPart destPresPart = myDestDeck.PresentationPart;

                //Open up the source deck
                using (PresentationDocument mySourceDeck = PresentationDocument.Open(sourceDoc, true))
                {
                    PresentationPart sourcePresPart = mySourceDeck.PresentationPart;

                    //Need to get a unique ids for slide master and slide lists (will use this later)
                    uniqueId = GetMaxIdFromChild(destPresPart.Presentation.SlideMasterIdList);
                    uint maxSlideId = GetMaxIdFromChild(destPresPart.Presentation.SlideIdList);

                    //Copy each slide in my source deck in order to my destination deck
                    foreach (SlideId slideId in sourcePresPart.Presentation.SlideIdList)
                    {
                        SlidePart sp;
                        SlidePart destSp = null;
                        SlideMasterPart destMasterPart;
                        string relId = "";
                        SlideMasterId newSlideMasterId;
                        SlideId newSlideId;

                        //come up with a unique relationship id
                        id++;
                        try
                        {
                            sp = (SlidePart)sourcePresPart.GetPartById(slideId.RelationshipId);
                            Random random = new Random();
                            int randomNumber = random.Next(0, 10000);

                            string tmp = "Foo" + randomNumber;//id;//sourceDoc.Remove(sourceDeck.IndexOf('.'));
                            string[] path = tmp.Split('\\');
                            //relId =path[path.Length-1] + id;
                            relId = tmp;
                            //MessageBox.Show(relId);
                            destSp = destPresPart.AddPart<SlidePart>(sp, relId);

                        }
                        catch (XmlException e)
                        {
                            MessageBox.Show(e.Message);
                        }
                        //Master part was added, but now we need to make sure the relationship is in place
                        destMasterPart = destSp.SlideLayoutPart.SlideMasterPart;
                        destPresPart.AddPart(destMasterPart);

                        //Add slide master to slide master list
                        uniqueId++;
                        newSlideMasterId = new SlideMasterId();
                        newSlideMasterId.RelationshipId = destPresPart.GetIdOfPart(destMasterPart);
                        newSlideMasterId.Id = uniqueId;

                        //Add slide to slide list
                        maxSlideId++;
                        newSlideId = new SlideId();
                        newSlideId.RelationshipId = relId;
                        newSlideId.Id = maxSlideId;

                        destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);
                        destPresPart.Presentation.SlideIdList.Append(newSlideId);
                    }
                    //Make sure all slide ids are unique
                    FixSlideLayoutIds(destPresPart);

                }
                destPresPart.Presentation.Save();

                return destDoc;




                //using (Stream file = File.OpenWrite(destDeck))
                //{
                //  CopyStream(destDoc, file);
                //}


            }


        }

        public static MemoryStream CopyToMemory(Stream input)
        {
            // It won't matter if we throw an exception during this method;
            // we don't *really* need to dispose of the MemoryStream, and the
            // caller should dispose of the input stream
            MemoryStream ret = new MemoryStream();

            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                ret.Write(buffer, 0, bytesRead);
            }
            // Rewind ready for reading (typical scenario)
            ret.Position = 0;
            return ret;
        }


        public static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }


        static void FixSlideLayoutIds(PresentationPart presPart)
        {
            //Need to make sure all slide layouts have unique ids
            foreach (SlideMasterPart slideMasterPart in presPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    uniqueId++;
                    slideLayoutId.Id = (uint)uniqueId;
                }
                slideMasterPart.SlideMaster.Save();
            }
        }

        static uint GetMaxIdFromChild(OpenXmlElement el)
        {
            uint max = 1;
            //Get max id value from set of children
            foreach (OpenXmlElement child in el.ChildElements)
            {
                OpenXmlAttribute attribute = child.GetAttribute("id", "");

                uint id = uint.Parse(attribute.Value);

                if (id > max)
                    max = id;
            }
            return max;
        }
    }
    
}
