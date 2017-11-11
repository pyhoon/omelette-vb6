# Omelette
## *VB6 Database Application Framework*

**Omelette** is an open source project intended to be a RAD framework to build database driven applications in Visual Basic 6.0 for 32/64bit Windows XP/ME/7/8.X/10 and also Windows server 2000/2003. The omelette name was coming from a delicious dinner before this project started.

I found source code in VB6 can be opened or read using notepad. My idea is to store some frequently used templates and generate the source code files and use VB6 to compile them as an EXE or DLL.

VB6 project source code can be generated on the fly. With the addition of a "fake" command line form, developer can generate project files and MS Access database using a few commands.

### Example: 

`vb create Customer`

`vb open Customer`

`vb generate model Customer`


The above commands create a vb6 project inside _Projects_ folder name _Customer_ with an empty form and a folder name _Data_ with a MS Access database name _Data.mdb_ with a _Customer_ table.
