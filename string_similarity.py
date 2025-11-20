string1 = input("Enter the first string: ")
string2 = input("Enter the second string: ")

similar_count = 0
length = min(len(string1), len(string2))

for i in range(length):
    if string1[i] == string2[i]:
        similar_count += 1

print("Number of Matching Characteristics:", similar_count)
