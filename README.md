# omelette
VB6 Database Application Framework

Origin of omelette:
I had a delicious omelette dish for dinner before this project started.

Objective of this project is to provide a framework to develop database driven applications in vb6 for windows. I found source code in vb6 can be open or read using notepad. My idea is to store some frequently used templates and generate the source code files and use VB6 to compile them as an EXE or DLL.

VB6 project source code can be generated on the fly. With the addition of a "fake" command line form, developer can generate project files and ms access database using a few commands.

Example: 

vb create Customer

vb open Customer 

vb generate model Customer


The above commands create a vb6 project inside Projects folder name Customer with an empty form and a folder name Data with a MS Access database name Data.mdb with a Customer table.
