# ChemNumbering

**ChemNumbering** is a Microsoft Word-template that allows user to automotize the molecule numbering within the main body of the document.   
Microsoft Word equivalent of the ChemNum package in LaTex (minus ChemDraw modification).


# How to use ChemNumbering
Create a new document using ChemNumbering.dotm template.    
To add a new reference type _\cmpd{ref}_ or use the insert reference button in the **ChemNumbering** ribbon.
Then use the button in the ribbon to either get the number or the reference.      

More details about installation and use in the WIKI.


# Features
 - Allow multi numbering. ex: _\cpmd{ref1,ref2,ref3,ref4}_ will give _1-4_ or _\cpmd{ref1,ref3,ref4,ref5}_ will give _1,3-5_ given that "ref1" and "ref2" were already cited before.
 - Create a CSV file in the same directory where your document is saved.
 - Tested with Office365, Office 2016, Office 2013, Office 2011.
 
 
# Limitations
- This macro will replace only in the main body (table included). Textboxes and headers are excluded.
- This will not change the reference in the ChemDraw structure. Possible if scheme saved as CDXML file and added as linked object (more details to come). 
- Only available for Windows users


# Example
testNumber.docm is part of the introduction of my thesis. Please refrain from commenting the content.

## Enjoy


