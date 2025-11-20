a=0
b=1
times=int(input("enter numberr for fibonacci series :"))
for seq in range(times):
    print(a,end=" ")
    a,b=b,a+b
print("\n")
