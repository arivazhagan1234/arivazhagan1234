"""

. Contact Book

Concepts: dictionary, file handling

Features:

Add contact

Search contact

Delete contact

Save contacts in a JSON file

Example structure:

contacts = {
    "Ram": "9876543210",
    "John": "9998887777"9807
}
bn 
1.Requirment
* Save user phone number.
* Duplicate phone number
* Validation message handle
* Ensure valid  phone number
* Search phone number valid case

* Search invalid phone number
* Validation message for search
* Delete contact 
* Delete contact validation 
* search cantact after delete 
* Save contact on Json File

2.Pseudocode
START 
!
INPUT PHONE NUMBER
!
IF VALIDATE NUMBER
    DISPLAY VALID NUMBER
ELSE 
    DISPLAY ERROR MESSAGE
SAVE PHONE NUMBER IN FILE 

IF SEARCH PHONE NUMBER
    DISPLAY NUMBER
ELSE 
    DISPLAY NUMBER NOT AVAILALE
IF DELETE PHONE NUMBER
    SUCCESS MESSAGE
ELSE 
    NUMBER NOT AVAIABLE

3.Files
*Main files
*json
"""

contact_input=input("Enter a contact Name and Phone number: ").split(',')
print(contact_input)