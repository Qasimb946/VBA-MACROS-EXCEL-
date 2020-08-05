# VBA-MACROS-EXCEL-

The project was based on MACROS/VBA for merging two files and Creating a random username and password for unlimited data of customers.  The project was done using VBA excel and some other tools in excel. 

The final output was just 3 buttons that need to press to do the following tasks:

1- TO copy data from User file to Merged File
2- To copy data from Organizational data to the merged file
3- A button to create a Random Username and Random password for millions of members.


Some of the major Formulas are:

=CONCATENATE(LEFT(C2,20),RANDBETWEEN(1,10),RANDBETWEEN(1,10),CHAR(RANDBETWEEN(97,122)),CHAR(RANDBETWEEN(97,122)))

=CHAR(RANDBETWEEN(65,90))&CHAR(RANDBETWEEN(65,90))&CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(35,38))&RANDBETWEEN(1111,9999)

=CONCATENATE(LEFT(C2,20),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"))
