filename = input("Enter the file name: ")
n = int(input("Enter how many lines to display: "))
word = input("Enter the word to find frequency: ")

with open(filename, "r") as file:
    lines = file.readlines()

print("\n--- First", n, "lines of the file ---")
for i in range(min(n, len(lines))):
    print(lines[i], end="")

count = 0
for line in lines:
    words = line.split()
    for w in words:
        if w == word:
            count += 1

print("\n\nFrequency of the word '", word, "' is:", count)
