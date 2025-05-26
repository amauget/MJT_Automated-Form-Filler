# MJT_Automated-Form-Filler
This is a very simple automation solution that I've provided for my current employer. 
The sample register has 37 line items. It takes approximately 3 minutes to manually create a cover sheet. 
_**That means this solution saves 1 hour and 51 minutes for this job alone!**_ 
Now imagine the savings considering that our company generates hundreds of coversheets per year.

## Note: Read Below Before Trying for Yourself!
Wherever this sheet is located **must** have an associated directory named "AF3000_Wkr_Fldr". I tried to add error handling that would create the directory if it didn't exist, but VBA doesn't support the creation of directories on Sharefile (our company's file service).

## Background:
The excel file in this repository is called a submittal register. In oversimplified terms, this document tracks every item that we are required to send to our reviewing agency. Every line item is a legal package that requires a filled out cover sheet describing the item. 

Originally, the cover sheet was a word doc that was manually filled in, saved to PDF, and attached to the associated package. 

## What automated and how it works:
I integrated the coversheet (see page two, or the "AF3000" sheet) into the excel document and added a hyperlink to each row of the submittal register.
When the hyplink ("PDF") is selected, it captures the value of the submittal number, contract reference number, and item description, and pastes them into the AF3000 cover sheet on page 2.

After then relevant values are added to the coversheet, another macro is called, saving the coversheet to PDF, and naming it accordingly.

