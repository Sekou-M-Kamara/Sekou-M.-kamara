import random
import math
point_play = 0


def number_operation():

    def multiplication():
        end_num = input("\nlevel?⬇️\n")
        a = random.randint(0, int(end_num))
        b = random.randint(0, int(end_num))
        play_now = input(f"\n{a} x {b}⬇️\n")
        if a*b == int(play_now):
            global point_play
            point_play += 10
            print(f"\nCorrect😀\nAccumulated Point: {point_play}")
            continue_mul = input(
                "\nDo you want to  continue?\ny for yes\nn for no⬇️\n")
            if continue_mul == "y":
                multiplication()
            else:
                print(f"Thanks for  playing🙏\nAccumulated Point: {point_play}")
        else:
            print("\nYou are wrong😔\n")
            help_mul = input(
                "\nDo you need a help?\ny for yes\nn for no⬇️\n")
            if help_mul == "y":
                print("\nMultiplication is one of the four basic arithmetic operations, alongside addition, subtraction, and division. In math, multiply means the repeated addition of groups of equal sizes.\nTo understand better, let us take a multiplication example of the ice creams.\nEach group has ice creams, and there are two such groups. Total ice creams are 3 + 3 = 6.However, you have added two groups of 3 ice creams. Therefore, you have multiplied three ice creams by two. You may also write it as 2 x 3 = 6 ")
                continue_mul = input(
                    "\nDo you want to continue?\ny for yes\nn for no⬇️\n")
                if continue_mul == "y":
                    multiplication()
                else:
                    print(
                        f"Thanks for playing🙏\nAccumulated Point: {point_play}")

    def division():
        end_num = input("\nlevel?⬇️\n")
        try:
            a = random.randint(0, int(end_num))
            b = random.randint(1, int(end_num))
        except ValueError:
            print(f"{ValueError}: Empty  range")
        else:
            play_now = input(f"\n{a} / {b}⬇️\n")
            if math.floor(a/b) == math.floor(int(play_now)):
                global point_play
                point_play += 10
                print(f"\nCorrect😀\nAccumulated Point: {point_play}")
                continue_div = input(
                    "\nDo you want to continue?\ny for yes\nn for no⬇️\n")
                if continue_div == "y":
                    division()
                else:
                    print(
                        f"Thanks for playing🙏\nAccumulated Point: {point_play}")
            else:
                print("\nYou are wrong😔\n")
                help_mul = input(
                    "\nDo you need a help?\ny for yes\nn for no⬇️\n")
                if help_mul == "y":
                    print("\nThe division is the process of repetitive subtraction. It is the inverse of the multiplication operation. It is defined as the act of forming equal groups. While dividing numbers, we break down a larger number into smaller numbers such that the multiplication of those smaller numbers will be equal to the larger number taken. For example, 4 ÷ 2 = 2. This can be written as a multiplication fact as 2 x 2 = 4.")
                    continue_div = input(
                        "\nDo you want to continue?\ny for  yes\nn for no⬇️\n")
                    if continue_div == "y":
                        division()
                    else:
                        print(
                            f"Thanks for playing🙏\nAccumulated Point: {point_play}")

    def mixed_mul_div():
        end_num = input("\nlevel?⬇️\n")
        try:
            a = random.randint(0, int(end_num))
            b = random.randint(1, int(end_num))
        except ValueError:
            print(f"{ValueError}: Empty range")
        else:
            play_now = input(f"\n{a} / {b} x {a}⬇️\n")
            if math.floor(a/b*a) == math.floor(int((play_now))):
                global point_play
                point_play += 10
                print(f"\nCorrect😀\nAccumulated point: {point_play}")
                continue_mix = input(
                    "\nDo you want to continue?\ny for yes\nn for no⬇️\n")
                if continue_mix == "y":
                    mixed_mul_div()
                else:
                    print(
                        f"Thanks for playing🙏\nAccumulated Point: {point_play}")
            else:
                print("\nYou are wrong😔\n")
                help_mul = input(
                    "\nDo you need a help?\ny for yes\nn for no⬇️\n")
                if help_mul == "y":
                    print("Multiplication is one of the four basic arithmetic operations, alongside addition, subtraction, and division. In math, multiply means the repeated addition of groups of equal sizes.\nTo understand better, let us take a multiplication example of the ice creams.\nEach group has ice creams, and there are two such groups. Total ice creams are 3 + 3 = 6.However, you have added two groups of 3 ice creams. Therefore, you have multiplied three ice creams by two. You may also write it as 2 x 3 = 6. \n\nThe division is the process of repetitive subtraction. It is the inverse of the multiplication operation. It is defined as the act of forming equal groups. While dividing numbers, we break down a larger number into smaller numbers such that the multiplication of those smaller numbers will be equal to the larger number taken. For example, 4 ÷ 2 = 2. This can be written as a multiplication fact as 2 x 2 = 4.")
                    continue_mix = input(
                        "\nDo you want to  continue?\ny for yes\n for no⬇️\n")
                    if continue_mix == "y":
                        mixed_mul_div()
                    else:
                        print(
                            f"Thanks for playing🙏/nAccumulated  Point: {point_play}")

    def choice_game():
        print("\n**********************************************\n")
        print("WELCOME TO THE GAME OF MIND, NUMERICAL OPERATION\n")
        print("\n**********************************************\n")
        choose_game = input(
            "\nplease choose the game of your choice\nMultiplication(-m)\nDivision(-v)\nMised Multiplication Division(-mv)⬇️\n")
        if choose_game == "-m":
            multiplication()
        elif choose_game == "-v":
            division()
        elif choose_game == "-mv":
            mixed_mul_div()
        else:
            print("\nYou have to choose one  of the above⬆️")
            choice_game()
    choice_game()


number_operation()
