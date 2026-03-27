"Calculator function addition, subtraction. multipilication"


def Add(*user_inputs):
    if len(user_inputs) < 2:
       return "Please enter atleast two numbers"
    
    add_val=user_inputs[0]
    
    for num in user_inputs[1:]:
        add_val +=num
    return add_val

def Sub(*user_inputs):
    if len(user_inputs) < 2:
        return "Please enter atleast two numbers"

    sub_val=user_inputs[0]

    for num in user_inputs[1:]:
        sub_val -=num
    return sub_val

def Mul(*user_inputs):
    if len(user_inputs) < 2:
        return "Please enter atleast two numbers"
    
    mul_val=user_inputs[0]

    for num in user_inputs[1:]:
        mul_val *=num
    return mul_val

def Div(*user_inputs):
    if len(user_inputs) < 2:
        return "Please enter atleast two numbers"

    div_val=user_inputs[0]
    
    for num in user_inputs[1:]:
        if num==0:
            return "Zero division error"
        div_val /=num
    return div_val


try: 
    user_choice=input("Enter the any options in ADD, SUB, MUL, DIV : ")
    option=user_choice.upper()
    user_inputs= list(map(int, input("Enter the numbers: ").split( )))

    if option in ['ADD', '+']:
       print(f'Addition of {user_inputs} is :', Add(*user_inputs)) 
    elif option in ['SUB', '-']:
        print(f'Subrtaction of {user_inputs} is :', Sub(*user_inputs))
    elif option in ['MUL', '*']:
        print(f'Multiplication of {user_inputs} is :', Mul(*user_inputs))
    elif option in ['DIV', '/']:
        print(f'Divison of {user_inputs} is :', Div(*user_inputs))
    else:
        print("Please select any of the options in ('ADD', 'SUB', 'MUL', '+', '-', '*')")
except Exception as e:
    print(e)
finally:
    print("Execution is completed!!!")