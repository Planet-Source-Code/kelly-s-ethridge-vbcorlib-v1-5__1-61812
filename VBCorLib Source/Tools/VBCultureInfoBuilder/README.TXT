VBCultureInfoBuilder

This application is used to build the VBCultures.nlp file. The VBCultures.nlp file is used
by the CultureTable class in VBCorLib. It contains the culture specific information used
in such areas as number and date formatting.

The file is built from data stored in XML files. One XML per culture. A set of 203 XML files
comes with the release of VBCorLib. VBCultures.nlp is already built and included with the
release. It is located in the VBCorLib folder.

To build the VBCultures.nlp from the XML files, run the VBCultureInfoBuilder application and
browse to the accompanied CultureInfoData folder found in the VBCultureInfoBuilder folder.

Once the file has been built, VBCorLib will attempt to access it in the App.Path folder, so 
where the executing application is located is where VBCultures.nlp should be placed. However,
if VBCorLib fails to find the file in the App.Path folder it will check in the Windows folder.

The purpose of allowing VBCorLib to use a file in the App.Path folder is to allow a custom
VBCultures.nlp file to be created for a specific application. The file found in the Windows
folder is shared amongst all application that use VBCorLib.

If the file is not in either location, then the only culture supported is Invariant.

