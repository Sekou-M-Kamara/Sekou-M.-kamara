import random

print("\nPlease make sure you have an identification number before creating an account.This helps us reaching you faster using numerical algorithm\n")


def identity_generator():
    full_name = input("\nPlease enter a prospective user name⬇️.\n")
    Idn_num = " "
    rand_code = random.sample(range(0, 9, 1), 3)
    for i in rand_code:
        Idn_num += str(i)
    Idn_num = int(Idn_num)

    print(f"\n{full_name}, your identification number has successfully being generated👍.\n\nIndentifiction_Number = {Idn_num}.")

    def inform_user():
        print(f"\n{full_name}, please use  both your identification number and prospective user name to  create a valid account.")
        print("\n\nFormat: user_name =  Projrct_A.account_creation('user_name', idn_num)\nIdn stands for  identificcation")
    inform_user()
    request_prom = input(
        "\nTo  understand our format of request, please input -h for help; else, any letter.\n")
    if request_prom == "-h":
        print("\nThese helps are useful after you have created an account\n\nuser_name.add_stu: To add a student to the list.\nuser_name.:  Gives you several methods you have at your diposal.\n")


class account_creation:
    def __init__(self, user_name, Idn_num):
        self.User_Name = user_name
        self.Idn_num = Idn_num
        self.student_num = random.sample(range(0, 9, 2), 4)
        self.stu_num = ""
        print(f"\n>>>>> {self.User_Name}\n")

    def add_stu(self):
        for i in self.student_num:
            self.stu_num += str(i)
        student_name = input("\nPlease enter a student  name⬇️.\n")
        with open("stu_list", "a") as f:
            f.write(
                f"name: {student_name} ---- admin#: {int(self.stu_num)}\n")
        with open("stu_list", "r") as f:
            read_list = f.read()
        print(f"\n{student_name} is  added to  the list.\n{read_list}")
        prom_again = input("\nAgain?\nY or N\n")
        if prom_again == "Y":
            self.add_stu()
        else:
            print("You exit!")

    class Grades_Depositor:
        def __init__(self, stu_num, grade, subject):
            self.stu_num = stu_num
            self.grade = grade
            self.subject = subject

            with open("grades_list", "a") as f:
                f.write(
                    f"stu_num: {self.stu_num} -- grade: {self.grade} -- sub: {self.subject}\n")
            with open("grades_list", "r") as f:
                read_grade = f.read()
            print(
                f"\nGrade successfully  added\nSee info below\n\n{read_grade}")
