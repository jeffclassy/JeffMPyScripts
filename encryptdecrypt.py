def split(word): #encrypts each letter of the url its ascii code + 1
    return [format(ord(char)+88,'03d') for char in word]
def chop(double): #decrypts every 3char from a string of encrypted url
	return [chr(int(double[i:i+3])-88) for i in range(0, len(double), 3)]
     

word = IN[0]
intword = IN[1]

encryptword = ''.join(split(word))
decryptword = ''.join(chop(intword))

#rint(decryptword,encryptword)
OUT=decryptword,encryptword

