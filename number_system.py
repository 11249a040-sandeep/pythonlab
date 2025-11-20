def binary_to_decimal(binary_str):
    try:
        decimal=int(binary_str,2)
        return decimal
    except valueError:
        return"invalid binary number!"
def octal_to_hexadecimal(octal_str):
    try:
        decimal=int(octal_str,8)
        hexa=hex(decimal).upper(),replace("OX","")
        return hexa
    except valueError:
        return"invalid octal number!"
print("Number system convertion")
print("1.binary to decimal")
print("2.octal to hexadecimal")
choice=input("enter your choice(1 or 2):")
if choice=="1":
    binary_str=input("enter the binary number :")
    print("decimal value :",binary_to_decimal(binary_str))
elif choice=="2":
    octal_str=input("enter the octal number :")
    print("hexadecimal value :",octal_to_hexadecimal(octal_str))
else:
    print("invalid choice!please enter 1 or 2")
