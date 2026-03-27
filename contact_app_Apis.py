import json
from flask  import Flask, jsonify, request
import os 


app=Flask(__name__)

#file name
contact_file="contact.json"

def load_contact():
    if os.path.exists(contact_file) and os.path.getsize(contact_file)>0:
            try:
                with open(contact_file, 'r') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                return {}
    return {} 

@app.route('/add', methods=['POST'])
def Add_contact():
    req=request.json
    name=req.get("name")
    phone=req.get("phone")

    if len(phone)==10 and phone.isdigit():
        data=load_contact()
        if name in data:
            return jsonify({"message":"Contact already exists"})
        
        data[name]=phone
        sort_data=dict(sorted(data.items()))
        Save_contact(sort_data)
        return jsonify({"message":"Contact successfully saved!!!"})      
    else:
        return jsonify({"message":"Enter valid length phone of number"})

def Save_contact(data):
    with open(contact_file, 'w') as f:
        json.dump(data, f, indent=4)

@app.route('/search/<name>', methods=['GET'])      
def Search_contact(name):

    data=load_contact()  
    print(data)
    return jsonify({name:data.get(name, "Contact doesn't exist")})

@app.route('/update', methods=['PUT'])
def Update_contact():
    req=request.json
    name=req.get("name")
    phone=req.get("phone")

    if len(phone)==10 and phone.isdigit():
        data=load_contact()

        if name not in data:
            return jsonify({"message":"Contact not exists"})
        
        data[name]=phone
        sort_data=dict(sorted(data.items()))
        Save_contact(sort_data)
        return jsonify({"messafe":"Contact successfully updated!!!"})      
    else:
        return jsonify({"message":"Enter valid length of phone number"})

@app.route('/delete/<name>', methods=['DELETE'])
def Delete_contact(name):
    
    data=load_contact()    

    if name in data:
        print(f'Before delete contact {name} and {data[name]}')
        del data[name]
        Save_contact(data)
        return jsonify({"message":"Deleted successfully"})
    else:
        return jsonify({"message":f"Contact {name} doesn't exist"})

@app.route('/view', methods=['GET'])
def View_contacts():                                                                                                                 
    data=load_contact()
    return jsonify(data) 


if __name__=='__main__':
    app.run(debug=True)
