# MicrosoftAccessImportCSV
An intelligent CSV Import


'I assume this file is just a plain text file
'MicrosoftAccessImportCSV
'Features:
'takes a CSV file and tries to figure out which columns should go to what columns in your Microsoft Access Database automatically
'It will also regonize which columns must be imported at the same time (see example 1). It will also regonize if you have repeated columns
'and import them as a second entry (example 2).

'Example 1: Consider a database of two tables, one of a person and the other of a phone number. the person has a first and last name field
'and the phone number just has a phone number field. When you import, you most likely want the first and last name 
'on the same row in the database, the program should regonize this during import

'Example 2: consider an import with three phone numbers and phone number type, the program should import each CSV row as three rows in your
'databse


'HOW TO USE:
'Import all VBA files into Microsoft Access
'Set the path variable to the directory pointing to your CSV file
'run and follow the prompts as provided
'DISCLAMER: I wouldn't be surprised if this is still pretty buggy. 
