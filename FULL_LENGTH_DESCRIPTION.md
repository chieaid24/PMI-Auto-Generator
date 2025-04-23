<h1 align="center">PMI Auto Generator</h1>

> The first of its kind, this script uses 2D annotations from a technical drawing to attach the relevant PMI to a blank 3D model.



![Untitled design](https://github.com/user-attachments/assets/a33199ed-4859-445d-99cd-a515ee304dd3)

All it requires is an Excel spreadsheet containing the annotations and a 3D model for the part you wish to annotate. It then cross-references the nominal values with the geometry information to **attach all valid diameter annotations with PMI** (Product Manufacturing Information).

Check out the demo video [here!]

<h2>⚡️Two ways to use it</h2>
<h3>With MBDVidia (preferred)</h3>

Built with *AutoHotKey*, the program will guide you with step-by-step instructions through converting your bubble drawings and .stp files to the correct formats. </br>

**Setup steps**:
 - Set your display resolution to `1920 x 1080p` - _allows for image recognition with AutoHotKey_
 - Within MBDVidia, bind `Export Report Set...` to `(Ctrl + W)`
____

<h3>Without MBDVidia</h3>

Requires more setup and a software that can convert 3D geometry files into QIF format (_ex. Inventor, FreeCAD, SOLIDWORKS_). </br>

**Setup steps**:
 - Use the Excel template located in `resources/QIFParser Excel Template.xlsx` and fill out the blanks in conformance to the example provided
   - Save the `.xlsx` file into `INPUT FILES`
 - Convert your blank 3D model into a .QIF file with your software of choice
   - Save the `.qif` file into `INPUT FILES`
 - **Note:** These files will be removed after the program is completed
___
<h2>⚙️ Additional Capabilities</h2>

 - Recognize and assign `asymmetric tolerances`
 - Recognize `duplicate features` and define annotations for each
 - Prompt user for `default tolerances` and apply when necessary
___
<h2>🔭 Further Steps</h2>

 - Check out my [MBD Macro](https://github.com/chieaid24/MBD-Macro) for an even further expedited PMI process!
