
 This is a DEMO Application on Using my FileMapping VB Class Module
   to Simulate the use of File Mapping APIs

       Used in Visual C++ and Win32 Assembly, Why in Visual Basic, Why Not?


   FileMapping Class Module
   ========================

   Tired of using the traditional Visual Basic File Handling Routines, or
   even FileSystemObject failed to read to a file byte-by-byte, or you want to
   Access a File Using Pointers for Better Scalability Performance, while
   Some Anti-Virus Blocks the use of very useful FileSystemObject, 

   File Mapping Functions comes into a handy set of routines just waiting for
   you to use them, and take note, we people as Visual C++ and Win32 Assembly 
   uses this functions, because we dont have easy macros.

   Why use easy routines, and limiting your performance, use File Mapping, thats
   the only way!!!

   yup, Visual Basic doesn't support Pointer Variables, but there is a CopyMemory
   "RtlMoveMemory" to do the same effects (a little slower than the real pointers
   but that will do just fine)

   Included in this ZIP file is my Class Module to perform FileMapping and a demo
   Project on how to use the File Mapping by Traversing to an Executable format
   and checks whether it is a valid Win32 Portable Executable Format and finally
   checks what type of PE it is, PE32 or PE64



   ========================================================================
   This code checks for a valid PE Executable Format

   To understand this Application, you need to consult your nearest
   PE Documentation.

   Win32 Assembly Codes are included in Comments are 100% working on
   TASM32 Compiler



   Created by: Chris Vega [gwapo@models.com]
               http://trider.8m.com
