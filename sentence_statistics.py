Sentence = input("Enter a Sentence: ")
words = Sentence.split()
word_count = len(words)
upper_count = 0
lower_count = 0
digit_count = 0

for ch in Sentence:
    if ch.isupper():
        upper_count += 1
    elif ch.islower():
        lower_count += 1
    elif ch.isdigit():
        digit_count += 1

print("Num of words:", word_count)
print("No of upper case:", upper_count)
print("Lower case letters:", lower_count)
print("Digits:", digit_count)
