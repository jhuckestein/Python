import sys

#### Now I need to take in the input from the command line where sys.argv[0] is the name of this
#### file, and sys.argv[1] is the "password" that was supplied.

passString = sys.argv[1]
length = len(passString)
print(passString)
passArray = list(passString)  ####Convert the input into an array of characters
print("passArray = ", passArray)
violation = False
notAllowed = ">!<"   #### This definition can be updated if new unacceptable characters are found
for counter in range(0, length): #### Take the passArray and look at each character
	test = notAllowed.find(passArray[counter])
	#print("Test = ", test)
	if (test != -1):
	    violation = True
	#print("V= ", violation)
if (violation == True):
	#### Now stop execution and send the user back to the login etc.
	print("Please choose a valid password '0-9', 'a-z', 'A-Z', ')(*&^%$#@}{?'")
else:
	#### Now go do whatever was requested etc.
	print("Password accepted as valid")
