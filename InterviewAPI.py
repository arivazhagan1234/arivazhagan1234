from flask import Flask, request, jsonify

app = Flask(__name__)

contact = {}

@app.route("/addition", methods = ["POST"])
def addition():
    data = request.get_json()
    val1 = data.get("val1")
    val2 = data.get("val2")
    val3 = data.get("val3")

    if val2 == 0:
        raise ZeroDivisionError("Cannot divide by zero")
    if val1 and val2:
        return jsonify( {"result": val1 + val2 + val3})

@app.route("/multiplication", methods = ["POST"])
def multiplication():
    data = request.get_json()
    val1 = data.get("val1")
    val2 = data.get("val2")
    val3 = data.get("val3")
    return jsonify({"result": val1 * val2 * val3})


@app.route("/addcontact", methods = ["POST"])
def addcontact():
    data = request.get_json()
    name = data.get("val1").strip()
    number = data.get("val2")
    email = data.get("val3").strip()

    print(f"name : {type(name)} | number : {type(number)} | email : {type(email)}")

    if not isinstance( name, str) or name in (" ", None) or not isinstance(number, int) or number in ("", None, str) or not isinstance(email, str) or email in ("",None):
        return jsonify( {"message" : "The credencial should be valid"})

    if name not in contact :
        contact[name] = {"number" : number, "email" : email}
        return jsonify({"message": "contact saved",
                        "contact": contact}), 200
    else:
        return jsonify({"message" : "contact already exists"}), 200
 

if __name__ == "__main__":
    app.run(debug = True)