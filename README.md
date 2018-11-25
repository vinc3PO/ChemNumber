# ChemNumbering

ChemNumbering is a template document that allows user to automotize the numbering of the molecule within the main body of the document. 
Similar to the ChemNum package in LaTex.


# How to use ChemNumbering
Create a new document using ChemNumbering.dotm template. 
To add a new reference type \cmpd{ref} or use the insert reference button in the ChemNumbering ribbon.
Then use the button in the ribbon to either get the number or the reference. 


# Features
 - Allow multi numbering. ex: \cpmd{ref1,ref2,ref3,ref4} will give 1-4 or \cpmd{ref1,ref3,ref4,ref5} will give 1,3-5 given that ref1 and ref2 were already cited before.
 - Create a CSV file in the same directory where your document is saved.
 
 
# Limitation
- This macro will include reference in the main body. Textboxes and headers are excluded.
- This will not change the reference in the ChemDraw structure. 


