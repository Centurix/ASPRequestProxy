## ASP Request Proxy ##

What is the ASP Request Proxy?

There are several ways of getting information from the web client browser to a web server. For example sending text to the server from an input form or posting a URL with parameters. A problem arises when you need to send an actual file to the web server for processing. Normally there are no built in methods to manage the information from the browser, and the information has to be gathered by other means. This script allows you to send files from the browser to the server by actually deciphering the binary information sent from a browser using the multipart/form-data encoding type. There are other components out there which do the same thing. A lot of them are COM components, which require installation onto a web server by an administrator, which you may not have the rights to do. This component is written in entirely script, which means that you can upload it and get the same results. There are also some other scripts out there which do the same thing as well, so why is this one any different? Easy, it's fast.

How does it work?
There is a standard outlined by a body called the W3C. This standard specifies 
how binary information should be sent from a browser to a web server. This is achieved by posting a file through a form on a web page. Same as any other form, with the exception that the form tag contains a specific encoding type. This is shown in later examples. This is used in conjunction with a special attribute on the input tag. This gives a special button on a form which allows the user to select a file from their hard drive to upload. When the form is submitted, the file is transferred with the form data. This is fairly straight forward until the server retrieves the file and tries to do something with it. This is where the ASPProxyRequest script comes into play. The first thing it does is scan the incoming request information for form data and then processes it, looking for binary information and normal field information. Once this is all found, it marks it's spots so that when the data is referenced, it is redily available.

Requirements
The basic requirements for ASPProxyRequest include a web server running IIS, VBScript version 5 and higher, and a version of ADO greater then 2.6.

Installation Instructions
Copy class_request.asp to a directory on your web server which has script access.

Usage
To utilise this script, you create an instance of the class. This will 
immediately search the Request object for valid data. From that point onwards, the data can be acquired as if the object was the Request object itself. Example (save this text as a file called file2.asp):

```
<--#include file="class_request.asp"-->
Dim objRequest, sText

Set objRequest = New ProxyRequest
sText = objRequest("field1")
```

This example first includes the ASPProxyRequest file and creates an instance of the class. It then grabs the contents of the "field1" element from the previous page's form and places the contents into the variable sText. The previous page would simply look like the following:

```
<form action="file2.asp" enctype="multipart/form-data" method="post">
<input type="file" name="field1">
<input type="submit" value="Upload...">
</form>
```

Note the inclusion of the enctype attribute above. This is important in order for this to work correctly. If you have entered this 
correctly into your page, you should see a button like the one below:

```
<input type="file">
```

Clicking the browse button will bring up the file selection dialog. Examples have also been&nbsp;supplied with this script which show how to upload a file. You will also note that this script is able to upload multiple files at any one time. This is done simply by adding more file elements to the form.

Methods

BinaryRead() - Retrieves the binary information from the request object.
SaveItem(vIndex,sFilename) - Saves an element to file, vIndex is the name of the element to save. sFilename is the file and path to save the data to.
ExportRequest - Creates a serialized string of the current request object, so it can be re-loaded into the request object at a later time.
ImportRequest(sText) - Imports a previously serialized version of the request object created by ExportRequest()

Properties
Cookies - Cookies collection, see the Request object
ClientCertificate - ClientCertificate collection, see the Request object
ServerVariables - ServerVariables collection, see the Request object
QueryString - QueryString collection, see the Request object
TotalBytes - Total number of bytes in the request object.
Form - Form element collection, see the Request object
Item - General item collection from both forms and query strings. See the Request object. Default property.

Extra file information
Along with all of the fields uploaded via a form, the file fields are accompanied with extra information about the file type and original file name. These can be accessed like they were form information. These extra variables add an extension to the form element names. If the file form element was called "file1" then the extra file information is stored in:

file1_mimetype - The approximate file content type
file1_filename - The original file name from the browser
file1_sourcepath - The original path of the uploaded file

Why is this quicker than other request script only request parsers?
This script is different in that it circumnavigates the VBScript string concatenation issues which generally hurt performance in regular applications. It does this by using ADO Streams to generate data. Very little is done in terms of transferring blocks of memory back and forth. Most other script uploaders will process the binary data with VB strings, adding a huge amount of overhead to an application.
