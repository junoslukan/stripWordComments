# stripWordComments
Change the author name and/or remove timestamps from comments in Microsoft Word documents

This is a simple script that allows hiding the time and date from comments in Word or changing the name of the comments' author. 
This is done by extracting the .docx file as a .zip archive and changing the metadata in .xml files.

Use this script by specifying the filename -f, the new author name, -a and whether the date should be hidden, -d:
```console
foo@bar:~/pathToFile/$ stripWordComments -f "filename.docx" -a "Anonymous author" -d
```

A file with "_new" appended is produced.

Note: The timestamps are preserved in the .xml files (they are simply ignored) and the author initials are left unchanged.
There is also other identifiable information in .docx files and this script is not meant to make the author anonymous.
