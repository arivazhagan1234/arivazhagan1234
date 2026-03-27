import json
import os 

#file name
contact_file="contact.json"

def load_contact():
    if os.path.exists(contact_file) and os.path.getsize(contact_file)>0:
            with open(contact_file, 'r') as f:
                return json.load(f)
    else:
        return {} 
  
def Add_contact(num):
    
    if len(phone)==10 and phone.isdigit():
        data=load_contact()

        if name in data:
            return "Contact already exists"
        
        data[name]=phone
        sort_data=dict(sorted(data.items()))
        Save_contact(sort_data)
        return "Contact successfully saved!!!"      
    else:
        return "Enter valid length phone of number"

def Save_contact(data):
    with open(contact_file, 'w') as f:
        json.dump(data, f)
       
def Search_contact(num):
    data=load_contact()  
    name=name
    print(data)
    return data.get(name, 'Contact is not present')

def Update_contact(num):
    if len(phone)==10 and phone.isdigit():
        data=load_contact()

        if name not in data:
            return "Contact not exists"
        
        data[name]=phone
        sort_data=dict(sorted(data.items()))
        Save_contact(sort_data)
        return "Contact successfully updated!!!"      
    else:
        return "Enter valid length phone of number"

def Delete_contact(num):
    name=name
    data=load_contact()    

    if name in data:
        print(f'Before delete contact {name} and {data[name]}')
        del data[name]
        Save_contact(data)
        return "Deleted successfully"
    else:
        return f'Contact {name} not present'

def View_contacts():                                                                                                                 
    data=load_contact()
    for name,number in data.items():
        print(name,':', number, end=' , ') 


try:
    while True:
        contact_option=input("Select any one of contact options in 'Add', 'Search','Update', 'View', 'Delete', 'Exit': ")

        if contact_option=='View':
            View_contacts()
        elif contact_option in ['Exit']:
            exit()
        else:
            contact_input=[x.strip() for x in input("Enter a contact Name and Phone number: ").split( )]
        
        print()

        if len(contact_input) !=2 and contact_option=='Add':
            print("Invalid format enter Name, Phone number")
            continue


        if contact_option =='Add':
            print(Add_contact(contact_input))
        elif contact_option =='Search':
            print(f'The searched contact of {contact_input}: ', Search_contact(contact_input))
        elif contact_option=='Delete':
            print(Delete_contact(contact_input))
        elif contact_option=='Update':
            print(Update_contact(contact_input))
        else:
            print("Please Select any one of contact options in 'Add', 'Search', 'View', 'Delete', 'Exit: ")

except Exception as e:
    print("Error Error", e)