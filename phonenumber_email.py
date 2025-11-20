import re  

text = input("Enter a text: ")
phones = []
emails = []

phone_pattern = re.compile(r'\b\d{10}\b')
email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')

phones = phone_pattern.findall(text)
emails = email_pattern.findall(text)

print("Phone Numbers found:", phones)
print("Email IDs found:", emails)
