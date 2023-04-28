A Python script that automatically converts PPT/PPTX  under working dir to PNG files and combines them to PDF after applying a watermark (optional, if you want).
requirement:

- math
- win32com
- shutil
- Pillow
- Reportlab

this repository is forked from [PPT2PDF of ernestyao's](https://github.com/ernestyao/PPT2PDF), I make serval changes in it to make it more useful.

- remove unused packages.
- lint the code to make it more readable.
- add the option of whether to add a watermark in your PDF.
- you can choose a '.txt' file as you watermark text.

BTW, you can add this script to your system path, then you can call it in your CMD/Powershell.

