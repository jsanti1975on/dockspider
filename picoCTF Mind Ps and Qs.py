# picoGym (picoCTF) Mind your Ps and Qs
# cat values // command to use to read the file named values
# Decrypt my super sick RSA:
'''
c: 421345306292040663864066688931456845278496274597031632020995583473619804626233684
n: 631371953793368771804570727896887140714495090919073481680274581226742748040342637
e: 65537   
'''
# Lets get two prime factors when multiplied they get the n: value 
# Below the factors were retrieved from factordb
# p = 1461849912200000206276283741896701133693
# q = 431899300006243611356963607089521499045809
# Now we have a value for p and a value for q, Lets set up all the values we have now 
c = 421345306292040663864066688931456845278496274597031632020995583473619804626233684
n = 631371953793368771804570727896887140714495090919073481680274581226742748040342637
e = 65537
p = 1461849912200000206276283741896701133693
q = 431899300006243611356963607089521499045809 
t = (p - 1) * (q - 1)
# Calculate totient <- Notes 
#T = (p-1) * (q-1) 
# Get d value 
# e * d = 1 mod t => d = e^-1 mod t
d = pow(e, -1, t)
p = pow(c, d, n)
print(bytearray.fromhex(hex(p)[2:]).decode('ascii'))
                                           
                                           



# use p for plain text, p = pow(c, d, n)
# print(p)  This will be a decimal value and needs to be hex decimal therefore convert to hex which is used to convert to ASCI
# print(bytearray.fromhex(hex(p)[2:]).decode(‘ascii’))



