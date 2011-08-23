using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Marklogic.Xcc;

namespace Poller
{
    class Program
    {
        public class ModuleRunner
        {
            private readonly Session session;
            private readonly ModuleInvoke request;
            private RequestOptions options;

            /// <summary> The Request object used internally to submit
            /// requests.  This object can be used to set external variables
            /// that will be bound to the query when submitted.  You should
            /// not set your own RequestOptions object, use
            /// the RequestOptions property instead.
            /// </summary>
            /// <returns> An instance of Request.
            /// </returns>
            public virtual Request Request
            {
                get { return request; }
            }

            /// <summary> Set (or clear) the RequestOptions instance to associate
            /// with submitted queries.
            /// </summary>
            /// <param name="options">An instance of RequestOptions or null.
            /// </param>
            public virtual RequestOptions RequestOptions
            {
                set { this.options = value; }
            }

            /// <summary> Construct an instance that will invoke modules on the
            /// server represented by the given URI.  Note that the URI will
            /// not be validated at this time.
            /// </summary>
            /// <param name="serverUri">A URI that specifies a server as per
            /// ContentSourceFactory.NewContentSource(URI).
            /// </param>
            /// <throws>  XccConfigException If the URI is not a valid XCC server URL. </throws>
            public ModuleRunner(Uri serverUri)
            {
                ContentSource cs = ContentSourceFactory.NewContentSource(serverUri);
                session = cs.NewSession();
                request = session.NewModuleInvoke(null);
            }

            /// <summary> Invoke the module with the given URI and return the resulting ResultSequence.
            /// </summary>
            /// <param name="moduleUri">A URI that specifies a module to invoke on the server.
            /// </param>
            /// <returns> An instance ResultSequence, possibly with size zero.
            /// </returns>
            /// <throws>  RequestException If an unrecoverable error occurs when submitting
            /// or evaluating the request.
            /// </throws>
            public virtual ResultSequence Invoke(String moduleUri)
            {
                request.ModuleUri = moduleUri;
                request.Options = options;

                return session.SubmitRequest(request);
            }

            /// <summary> Invoke the module with the given URI and return the resulting ResultSequence
            /// as an array of Strings.
            /// </summary>
            /// <param name="moduleUri">A URI that specifies a module to invoke on the server.
            /// </param>
            /// <returns> An array of Strings, possibly with size zero.
            /// </returns>
            /// <throws> RequestException If an unrecoverable error occurs when submitting
            /// or evaluating the request.
            /// </throws>
            public virtual String[] InvokeToStringArray(String moduleUri)
            {
                ResultSequence rs = Invoke(moduleUri);

                return rs.AsStrings();
            }

            /// <summary> Invoke the module with the given URI and return the resulting ResultSequence
            /// as a single String concatenation of all the values.
            /// </summary>
            /// <param name="moduleUri">A URI that specifies a module to invoke on the server.
            /// </param>
            /// <param name="separator">A String value to insert between the String
            /// representation of each item in the result sequence.
            /// </param>
            /// <returns> A String, possibly with size zero.
            /// </returns>
            /// <throws> RequestException If an unrecoverable error occurs when submitting
            /// or evaluating the request.
            /// </throws>
            public virtual String InvokeToSingleString(String moduleUri, String separator)
            {
                ResultSequence rs = Invoke(moduleUri);
                
                String str = rs.AsString(separator);

                rs.Close();

                return str;
            }

            /// <summary> Command-line main() method to invoke a module.</summary>
            /// <param name="args">Arg 1: A server URL as per
            /// ContentSourceFactory.newContentSource(URI).
            /// Arg 2: A module URI
            /// </param>
            [STAThread]
            public static void Main(String[] args)
            {
                //bool debug = true;
               /* if (args.Length < 2)
                {
                    usage();
                    return;
                }
                * */
                //xcc://username:password@localhost:8021 
                //hello.xqy
                //Uri serverUri = new Uri(args[0]);
                Uri serverUri = new Uri("xcc://oslo:oslol0g1c@localhost:8031");
                String moduleUri = "staging-count.xqy"; //args[1];

                ModuleRunner runner = new ModuleRunner(serverUri);
                //String result = runner.InvokeToSingleString(moduleUri, "\n");
                String[] result = runner.InvokeToStringArray(moduleUri);

           
                int length = result.Length;
                
                //if ( Int32.Parse(result) > 0)
                if(length > 0)
                {
                    Console.WriteLine("FILES FOUND: " + length);
                    DroidIngest di = new DroidIngest();
                    for (int i = 0; i < length; i++)
                    {
                        Console.WriteLine("FILE NAME: " + result[i]);
                        di.ingestDocs(result[i]);
                    }

                }
                else
                {
                    Console.WriteLine("NOTHING HERE: " + length);
                    
                }

                
                    Console.WriteLine("Press any key too continue...");
                    Console.ReadLine();
                
            }

            private static void usage()
            {
                Console.WriteLine("usage: serveruri docuri [-o outfilename]");
            }
        }
    }
}
