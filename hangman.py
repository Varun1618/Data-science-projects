
import random
name=input("Enter the name: ")
print("Good Luck! ",name)
wordslist=['rainbow','computer','science','programming','python','mathematics','player','condition','reverse','water','board']
word=random.choice(wordslist)
print("Start the Game!")
guesses=''
turns=12
while turns>0:
    fail=0
    for char in word:
        if char in guesses:
            print(char)
        else:
            print("_")
            fail+=1
    if fail==0:
        print("You won the game!")
        print("The word is: ",word)
        break
    guess=input("\nGuess a character: ")
    guesses+=guess
    if guess not in word:
        turns-=1
        print("Wrong!")
        print("You have ",turns,"more guesses")
    if turns==0:
        print("You lost the game! Try again...")
        print("The word is: ",word)
