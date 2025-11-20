number=input("enter to count the apperence :")
digit=input("enter which value you want to count :")
count=0
for element in number:
    if(element==digit):
        count=1
print("apperence of digit:\n",count)
