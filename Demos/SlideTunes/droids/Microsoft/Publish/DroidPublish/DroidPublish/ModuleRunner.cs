using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Marklogic.Xcc;

namespace DroidPublish
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

    }
}
