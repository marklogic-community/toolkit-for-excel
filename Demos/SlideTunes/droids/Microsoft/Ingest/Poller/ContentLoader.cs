using System;
using MarkLogic.Xcc;
namespace Marklogic.Xcc.Examples
{
public class ContentLoader
{
	public ContentLoader()
	{
		private readonly Session session;
		private ContentCreateOptions options = null;

		/// <summary> Set (or clear) an instance of ContentCreateOptions
		/// which defines creation options to apply to each document loaded.
		/// This is null (defaults) unless explictly set.
		/// </summary>
		public virtual ContentCreateOptions Options
		{
			set { this.options = value; }
		}

		/// <summary> Construct an instance that may be used to insert
		/// content.
		/// </summary>
		/// <param name="serverUri">A URI identifying a ContentSource,
		/// in the format expected by
		/// ContentSourceFactory.newContentSource(URI).
		/// </param>
		/// <throws>  XccConfigException Thrown if a Session cannot
		/// be created.  This usually indicates that the host/port or
		/// user credentials are incorrect.
		/// </throws>
		public ContentLoader (Uri serverUri)
		{
			ContentSource cs = ContentSourceFactory.NewContentSource (serverUri);
			session = cs.NewSession();
		}

		/// <summary> Load the provided files (represented by FileInfo objects),
		/// using the provided URIs, into the content server.
		/// </summary>
		/// <param name="uris">An array of URIs (identifiers) that correspond to the
		/// FileInfo instances given in the "files" parameter.
		/// </param>
		/// <param name="files">An array of FileInfo objects representing disk
		/// files to be loaded.  The ContentCreateOptions object
		/// set with the Options property,
		/// if any, will be applied to all documents when they are loaded.
		/// </param>
		/// <throws> RequestException If there is an unrecoverable problem
		/// with sending the data to the server.  If this exception is
		/// thrown, none of the documents will have been committed to the
		/// contentbase.
		/// </throws>
		public virtual void Load (String[] uris, FileInfo[] files)
		{
			    Content [] contents = new Content [files.Length];

			    for (int i = 0; i < files.Length; i++) {
				    contents [i] = ContentFactory.NewContent (uris [i], files [i], options);
			    }

			    session.InsertContent (contents);
		}

		/// <summary> Load the provided files into the contentbase,
		/// using the absolute pathname of each FileInfo as
		/// the document URI.
		/// </summary>
		/// <param name="files">An array of FileInfo objects representing disk
		/// files to be loaded.  The ContentCreateOptions object
		/// set with the Options property,
		/// if any, will be applied to all documents when they are loaded.
		/// </param>
		/// <throws>  RequestException If there is an unrecoverable problem
		/// with sending the data to the server.  If this exception is
		/// thrown, none of the documents will have been committed to the
		/// contentbase.
		/// </throws>
		public virtual void Load (FileInfo[] files)
		{
			String [] uris = new String [files.Length];

			for (int i = 0; i < files.Length; i++) {
				uris [i] = files [i].FullName.Replace ("\\", "/");
			}

			Load (uris, files);
		
	     }
    }
}
}


