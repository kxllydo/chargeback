
def validateOption (prompt, options):
    valid = False

    while not valid:
        try:
            option = int(input(prompt))
            if option in options:
                valid = True
            else:
                print("That was an invalid response. Try again.\n")
        except ValueError:
            print ("Input a number. Try again.\n")
    return option
