using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Marklogic.Xcc;

namespace DroidPublish
{
    class ContentFetcher
    {
        private readonly Session session;
        private RequestOptions options = null;

        /// <summary> <p>
        /// Set (or clear) the RequestOptions instance to associate
        /// with submitted queries.
        /// </p>
        /// <p>
        /// Note: It's a good idea to set CachedResult=false.  Since the
        /// data is being written straight out to a Stream
        /// there is no need to buffer the document first.  Streaming will
        /// also accommodate arbitrarily large documents without running out
        /// of memory.  Setting an options value of null will use defaults,
        /// which includes cached result mode.
        /// </p>
        /// </summary>
        /// <param name="options">An instance of RequestOptions or null.
        /// </param>
        public virtual RequestOptions RequestOptions
        {
            set { this.options = value; }

        }

        /// <summary> Construct an instance that may be used to fetch documents.</summary>
        /// <param name="serverUri">A URI identifying a ContentSource,
        /// in the format expected by
        /// ContentSourceFactory.newContentSource(URI).
        /// </param>
        /// <throws>  XccConfigException Thrown if a Session cannot
        /// be created.  This usually indicates that the host/port or
        /// user credentials are incorrect.
        /// </throws>
        public ContentFetcher(Uri serverUri)
        {
            ContentSource cs = ContentSourceFactory.NewContentSource(serverUri);

            session = cs.NewSession();

            options = new RequestOptions();

            options.CacheResult = false;		// stream by default
        }

        public ResultItem Fetch(String docUri)
        {
            Request request = session.NewAdhocQuery("doc (\"" + docUri + "\")", options);
            ResultSequence rs = session.SubmitRequest(request);
            ResultItem item = rs.Next();

            if (item == null)
            {
                throw new ArgumentException("No document found with URI '" + docUri + "'");
            }

            return item;
        }

        /// <summary> Command-line main() method to fetch a document.</summary>
        /// <param name="args">Arg 1: A server URL as per
        /// ContentSourceFactory.newContentSource(URI).
        /// Arg 2: A document URI.  Optional Args 3 and 4: "-o outputfilename"
        /// </param>
        /*  [STAThread]
          public static void Main(String[] args)
          {
              if (args.Length < 2)
              {
                  usage();
                  return;
              }

              Uri serverUri = new Uri(args[0]);
              String docUri = args[1];

              ContentFetcher fetcher = new ContentFetcher(serverUri);
              ResultItem item = fetcher.Fetch(docUri);

              if (args.Length == 4)
              {
                  if (args[2].Equals("-o"))
                  {
                      Stream outStream = new FileStream(args[3], FileMode.Create);
                      item.WriteTo(outStream);
                      outStream.Close();
                  }
                  else
                  {
                      usage();
                      return;
                  }
              }
              else
              {
                  item.WriteTo(Console.Out);
              }


              //			Console.WriteLine ("Fetched " + docUri + " in " + formatTime (DateTime.Now.Millisecond - start));
          }

          private static void usage()
          {
              Console.WriteLine("usage: serveruri docuri [-o outfilename]");
          }
         * */
    }
}
