import random

userchoice=[]
systemchoice=[]

for n in  range(10):
    number=random.randint(1, 5)
    #print('number)
    usernumb=int(input("Enter the user choice number: "))
                
    if usernumb==number:
        print("User chages number:", usernumb)
        userchoice.append(usernumb)
    else:
        print("System number", number)
        systemchoice.append(number)
   
print("User choice number is ", userchoice)
print("System choice numberL",systemchoice)