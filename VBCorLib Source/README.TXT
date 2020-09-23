** VBCorLib v1.5 ** 

	** BROKE ALL BINARY COMPATIBILITY ** (sorry)

The majority of the time was used to refactor the current code base.
Parameter checking has now become more centralized and easier to maintain.

Version 1.5
	- Added Example Projects
	- Refactored All Around Code
	- Renamed cObject To IObject
	- Renamed MappedFile To MemoryMappedFile
	- Added Ticker Class
	- Added StopWatch Class
	- Added INIFile Class
	- Added INISectionWriter Class
	- Added ExceptionBase Class
	- Added ResourcWriter Class
	- Added Several Resource Decoder Classes
	- Added Several Resource Encoders Classes
	- Upgraded WinResourceReader Class
	- Upgraded MemoryMappedFile Class
	- Upgraded Console Class
	- Fixed WeakReference Class
	- Fixed NumberFormatInfo Class
	- Fixed DateTimeFormatInfo Class
	- Fixed StringBuilder Class
	- Fixed DirectoryInfo Class


Version 1.2.2
	Extended the Console to .NET 2.0
	Extended Environment to .NET 2.0
	Extended File to .NET 2.0 (except security)
	Extended FileInfo to .NET 2.0 (except security)
	Extended cArray to .NET 2.0
	Added DriveInfo from .NET 2.0
	Added ConsoleKeyInfo from .NET 2.0
	Extended StringBuilder to .NET 2.0
	Extended cDateTime to .NET 2.0
	Extended RegistryKey to .NET 2.0
	Extended Registry to .NET 2.0
	Extended Constructors to .NET 2.0
	Extended Convert to .NET 2.0
	Fixed some bugs


Version 1.0.9
	Fixed Stream.ReadToEnd - Copied too many characters from buffer
	Updated StringBuilder.AppendFormat - Recognizes no formatting performed by a custom formatter
	Updated NumberFormatInfo.Format - Formats Integers upto 4 characters in hexadecimal
	Updated StringBuilder.AppendFormat - Allows NumberFormatInfo to format Integers now

** INTRODUCTION **
VBCorLib is a Visual Basic 6 reimplementation of the mscorlib.dll found in Dot NET.
This is the base library from which all of Dot Net builds its functionality.
VBcorLib brings much of that functionality to Visual Basic 6 through dozens of
classes and interfaces.

** SUPPORT WEBSITE **
A dedicated website for VBCorLib is located at www.kellyethridge.com/vbcorlib.
This site has documentation, tutorial, news and a forum for discussion. 

** BLOG **
I have added a blog site: http://vbcorlib.blogspot.com/

I seem to be better at giving updates on the blog than the support site.

** REQUIRED FILES **
The VBCorLib.dll file requires a second file 'vbcultures.nlp'. This file can
be in one of two places. Either in the same folder as the VBCorLib.dll file or
in the Windows directory. Not the Windows\System(32) folder.

** UNIT TESTS **
The source code comes with a set of unit tests that can be used to maintain the
integrity of VBCorLib. The unit tests require the SimplyVBUnit testing harness
with can be found at www.pscode.com or www.sourceforge.net/projects/simplyvbunit.
The unit tests do not need to be loaded in order to compile VBCorLib. Simply load
the VBCorLib project located in the VBCorLibGroup\VBCorLib folder.

The unit tests will also show how to use the classes in VBCorLib. Some of the classes
can be instantiated using New, some cannot. Viewing the unit tests will help see
how each class is used.

** TYPE LIBRARIES **
In order to load and compile VBCorLib, the VBCorLibTypes.tlb and VBVM6Lib.tlb need
to be registered. NOTE that this version of VBVM6Lib.tlb is not the same version that
is available from www.xbeat.net/vbspeed. This version has additional methods added.
These files are found in the 'VBCorLibGroup\VBCorLib\Type Libraries' folder.

** BINARY COMPATIBILITY **
If binary compatibility with my release of the compiled version is wanted, then it
can be acquired from www.kellyethridge.com/vbcorlib to download the binary version
of the file.

** USING BINARY INSTALL **
The binary install version of VBCorLib can be downloaded at www.kellyethridge.com/vbcorlib.
The installer allows the installation of VBCorLib and template classes, if you are
developing. The VBCorLib.dll and VBCultures.nlp files are placed into the Windows
folder. The VBCorLib.dll is then registered.

The template classes are placed in the VB98\Template\Classes folder. These classes
are base classes used to easily create VBCorLib supported classes.

** TOOLS **
There are tools used to create some of the required information used by VBCorLib. Currently
there is a VBCultures.nlp file builder and an EncodingData.bin file. More information
is available in those folders.

The tools are located in the 'VBCorLib Source v0.7\Tools' folder.

** FEEDBACK **
Comments or suggestions can be emailed to me at vbcorlib@kellyethridge.com
