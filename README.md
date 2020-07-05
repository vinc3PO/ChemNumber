
# CHEMNUMBERING


**ChemNumbering** is a Microsoft Word-template that allows user to automate the molecule numbering within the main body of the document.   
Microsoft Word equivalent of the ChemNum package in LaTex.  
This version also allows the modification of the ChemDraw scheme (Follow tutorial for instructions).


# How to use ChemNumbering
Create a new document using ChemNumbering.dotm template.    
To add a new reference in the word document type _\cmpd{ref}_ or use the insert reference button in the **ChemNumbering** ribbon.
Set-up the \{ref} markers in the ChemDraw scheme.  
Then use the buttons in the ribbon to either get the number or the reference.      

More details about installation and use in the tutorial.


# Features
 - Allow multi-numbering.  
 ex: _\cpmd{ref1,ref2,ref3,ref4}_ will give _1-4_ or _\cpmd{ref1,ref3,ref4,ref5}_ will give _1,3-5_ given that "ref1" and "ref2" were already cited before.
 - Create a CSV file in the same directory where your document is saved.
 - Tested with Office365, Office 2016, Office 2013 and Office 2011. In Office 2007 the macros work but there is no ribbon.
 
 
# Limitations
- This macro will replace only in the main body (table included). Textboxes and headers are excluded.
- If large document, it can take few minutes to complete the changes. 
- ChemDraw option is still very experimental and raise some errors.
- Only available for Windows users (I do not own any Apple devices, donation accepted ;)).

# Disclaimer
I am just a chemist who does not like to do boring and repetitive tasks when it is possible to automated them. I am not a developer, so the code might not be too power efficient, but a least it does the job.  

## Enjoy & Feel free to join and contribute


