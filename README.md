# time_sheet_generator_script
A scripts that I use for Time Sheet generation for european and other projects

The scripts are used for loading data from prepared XLSX style sheets.
It contains a local meomory for each time sheet for each person and it combines information from the work presence also included in the shteets.
It has hardcoded dates of public holidays and number of hourse each month have. the information are spread throu all files, mostli in Loaing script, that is used in both other scripts. 

It requires an Input folder which should contain XLSX sheet with all the information that are loaded in Loding script. Also a exporting template should be presented. It is referenced in Generate script. 
Person is a memory structure for the information, that is filled in the Loading script. 

Emailsend is a script that uses a SMTP server for sending the email - all files belonigng to single person (by email) at once. It uses a public SMTP withou any passowrd or authentication. 

The input folder should also contain an HTML encoded body of the email and TXT version as well. 