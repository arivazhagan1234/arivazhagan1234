import random   

for numb_otp in range(5):
    otp=random.randint(min(400534, 20016), max(400556,200561))
    print("The OTP is: ",otp)
