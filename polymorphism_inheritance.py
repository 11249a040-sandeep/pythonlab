class PalindromeChecker:
    def is_palindrome(self, value):
        pass


class StringPalindrome(PalindromeChecker):
    def is_palindrome(self, value):
        value = str(value)
        return value == value[::-1]


class IntegerPalindrome(PalindromeChecker):
    def is_palindrome(self, value):
        value = str(value)
        return value == value[::-1]


user_input = input("Enter a string or integer: ")

if user_input.isdigit():
    checker = IntegerPalindrome()
    result = checker.is_palindrome(int(user_input))
else:
    checker = StringPalindrome()
    result = checker.is_palindrome(user_input)

print("Palindrome:", result)
