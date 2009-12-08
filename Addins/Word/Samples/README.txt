Copyright 2008-2009 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

README.txt - instructions for enabling Samples with MarkLogic Server and Microsoft Word


SETUP INSTRUCTIONS:

To use the Samples with the Addin framework:

1) Create a new HTTP app server named samples in MarkLogic Server.
2) Set the root to the location of your this Samples directory.
    Example :
           C:\tmp\Samples
3) Set the port of the app server to a port not currently being used by MarkLogic or any other applications
    Example: 
           9000
4) Set the database to the database on the system where you are saving your Word documents.  
    Example:
           Documents

4) Set the URL for the Addin in the msi to the url of this new app server.  To see how to update the .msi configuration with the URL for the Addin, please refer to the FrameworkForWordGuide.pdf provided with this distribution.
    Example: http://localhost:9000

Now, assuming you've installed the Addin and it's url is properly set, you will see a default screen in the pane within Word when you open the Word application.  There will be 2 links: 1 for Search and 1 for Metadata.  Click the links to advance to the examples.


Search: 
---------------------------------

Assuming you've previously saved a .docx to MarkLogic.  Enter a term or phrase to search.  

Any paragraph (<w:p>) containing that text will be returned in the results below the Search box.  Double-Click the text and it will automatically be inserted into the active document.

Notice that though the text looks plain in the results; however, the style the content has in the original docx may be retained when it's inserted here.*

   *Style will be retained for default Word styles that are present for all .docx packages. You can however dynamically add custom and other styles to the the active document, but you'll need to write some more code.  If this interests you, please see word-processing-ml-support.xqy and the javascript api documentation. There are tool available to help make it happen.


Metadata:
---------------------------------

Enter details about the active document in the fields provided in the pane.  These fields will be saved as dublin core metadata in a custom XML part in the .docx package.  

Here we only save 1 custom XML part, but using the javascript api, you can add multiple custom parts of any well-formed XML to the .docx package.




