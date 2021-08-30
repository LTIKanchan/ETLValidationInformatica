from ConnectionInputs.DecAlgo import *  #Here we have taken all the data from DecAlgo file from ConnectionInput folder
from pathlib import Path

def readtext(wordfile,servernameintextfile):
    cred = []  #empty array
    flag=0  # Condition false
    try:
        file = open(wordfile, "r")  # it will open credfile
        linesintext = []  # Empty array, read data line by line
        for line in file:
            linesintext.append(line.strip())  # strip function will delete empty spaces
        file.close()  # to close the file

        for i in linesintext:
            servername = servernameintextfile
            if i.__contains__(servername) :  # will check for i contains servername or not
                credentials = i.split("-")  #split() method splits the string with the - separator
                # and returns a list object with string elements.
                #print("New User Name is " + decrypt(credentials[1]))
                #print("New Password is " + decrypt(credentials[2]))
                newusename = decrypt(credentials[1])
                newpass = decrypt(credentials[2])
                cred = [newusename,newpass]  #array of username& password
                flag=1  #condition true
                break  #terminate the loop
        if(flag==0):
            raise Exception("Incorrect server name or Credentials not Found....")

    except Exception as ex:
        raise ex
    return cred

def main():
    wordfile = str(Path(__file__).parent.parent) + "\\Test Data\\config.txt"
    readtext(wordfile,"USTRDD49.GENRE.COM")

if __name__ == "__main__":
    main()