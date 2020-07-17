
# Chemical Numbering For Word


**ChemNumber** is a Microsoft Word template which allows the user to automate the numbering of the molecules within the main body of a document.   
Microsoft Word equivalent of the ChemNum package in LaTex.  
This version also allows the modification of the ChemDraw scheme (See tutorial).

## Principle

This is a simple reference/number system. You only need to define each molecule with a unique identifier.  
Each identifier gets a number and the original identifier is stored as a hidden bookmark.  
Make it possible to revert the process.  


## How to use ChemNumbering (basics)
* Create a new document based on the ChemNumber.dotm template.    
* To add a new identifer\reference: 
  * type `\cmpd{ref}` 
  * Click the reference button in the ribbon.
* Click on the ribbon buttons to toggle between ref and numbers.      

See tutorial/instruction for ChemDraw editing possibilities.


## Features
 - Allow multi-numbering.  
    - `\cpmd{ref1,ref2,ref3,ref4}` will give `1-4`.
    - `\cmpd{ref1, ref2} \cpmd{ref1,ref3,ref4,ref5}` will give `1,2 13-5`.
 - CSV file containing id/number pairs.
 - Tested on Office365, Office 2016, 2013, 2011. In Office 2007, the macros work but  no ribbon.
 
 
## Limitations
- This macro will replace only in the main body (table included) of the document. Textboxes, headers, etc will be excluded.
- More content, more time. It's not that fast
- ChemDraw update option is still very experimental and might be giving some errors.
- Only available for Windows users.

## Disclaimer
This work, I have done during my PhD thesis to save me from the most boring task to number molecule in my thesis.  
This was my first contact with scripting and I have tried to optimized it as my knowledge in VBA grew.  
This code might not be super efficient, but a least it does the job it was meant to do.  

## Enjoy & Feel free to join and contribute


