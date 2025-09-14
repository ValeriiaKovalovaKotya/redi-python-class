while True:
    user_input = input("Enter a string: ")
    if user_input == "":
        print("Thanks for playing!")
        break
    print("First letter capitalized:\t".expandtabs(1) + user_input.capitalize())
    print("All lowercase:\t".expandtabs(13) + user_input.lower())
    print("Title case:\t".expandtabs(26) + user_input.title())
    print("All uppercase:\t".expandtabs(13) + user_input.upper())