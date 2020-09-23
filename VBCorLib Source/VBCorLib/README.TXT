** VBCORLIB **
This folder contains all of the files necessary to build the VBCorLib library project. 

The VBCultures.nlp must be in the Windows directory or in the directory 
the VBCorLib.dll file is located. 

Then EncodingData.bin file is loaded as a resource within then IDE. The
name of the resource must be "ENCODINGDATA" with a resource ID of 101.

The type libraries located in the TypeLibraries folder need to be registered.

The template folder contains classes that should be copied to the VB98/Template/Classes
folder to allow for easy creation of VBCorLib supported classes.
