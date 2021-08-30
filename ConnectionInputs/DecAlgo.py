def decrypt(text):
    result = ""
    s = 9  #string length of config.txt file
    for i in range(len(text)):
        char = text[i]
        if (char.isupper()):  #returns True if all characters in the string are uppercase,
            # otherwise, returns “False”
            result += chr((ord(char) + (26-s) - 65) % 26 + 65)  # encrypt the plain text
        elif char.isdigit():  #isdigit() method returns “True” if all characters in the string are digits,
            # Otherwise, It returns “False”
            c_new = (int(char) - s) % 10
            result += str(c_new)
        elif (char.islower()):  ##returns True if all characters in the string are lowercase,
            # otherwise, returns “False”
            result += chr((ord(char) + (26-s) - 97) % 26 + 97)  # encrypt the plain text
        else:
            result += char
    return result

def encrypt(text, s):
    result = ""
    for i in range(len(text)):
        char = text[i]
        if (char.isupper()):
            result += chr((ord(char) + s - 65) % 26 + 65)
        elif char.isdigit():
            c_new = (int(char) + s) % 10
            result += str(c_new)
        elif (char.islower()):
            result += chr((ord(char) + s - 97) % 26 + 97)
        else:
            result += char
    return result
