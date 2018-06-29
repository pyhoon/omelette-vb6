# Omelette
## *VB6 Database Application Framework*

**Omelette** is an open source project intended to be a RAD framework to build database driven applications in Visual Basic 6.0 for 32 bit WinXP, WinME but also supported in 64 bit Windows such as Win7, Win8.X, Win10, Win2K and 2003. The project name Omelette was coming from a delicious dinner before this project started.

I found source code in VB6 can be opened or read using notepad. My idea is to store some frequently used templates and generate the source code files and use VB6 to compile them as an EXE or DLL.

VB6 project source code can be generated on the fly. With the addition of a "fake" command line form, developer can generate project files and MS Access database using a few commands.

### Example: 

`vb create Customer`

`vb open Customer`

`vb generate model Customer`


The above commands create a vb6 project inside _Projects_ folder name _Customer_ with an empty form and a folder name _Data_ with a MS Access database name _Data.mdb_ with a _Customer_ table.

**Preview:**
![Compile and Run](https://github.com/pyhoon/omelette/blob/master/Source/Preview/Omelette_Compile_Run.png)

For more screenshots, check Omelette [Wiki](https://github.com/pyhoon/omelette/blob/master/Wiki) page.
