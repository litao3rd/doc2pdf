# doc2pdf
---
This is a very dirty code about convert Windows doc files to pdf files. It's my first time to write python code. Fortunately, it works fine. 

It's implemented by Windows Office COM interfaces, and call COM interfaces by **Python
for Windows Extensions**. You can get it from 
    [**HERE**](http://sourceforge.net/projects/pywin32/files%2Fpywin32).

This program accepts a directory path, and convert all the .doc files in that path
to .pdf file format. It works recursively. So it will covert all doc files in
its sub-directorys.
``` bash
$ doc2pdf.py [source-directory-path]

$ doc2pdf.py [source-doc-file-path]

$ doc2pdf.py [source-doc-file-path] [target-pdf-file-path]
```

I complete these code by learning some guys' works. The program contains some codes from
[**HERE**](http://blog.csdn.net/rumswell/article/details/7434302). Thanks for
your blog article and source code.
