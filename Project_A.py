# import random


# def identity_generator():
#     list_check = []
#     full_name = input("\nPlease enter a prospective user name.\n")
#     if full_name in list_check:
#         return
#     else:
#         rand_code = random.sample(range(10), 3).
#         list_check.append(full_name)
#         print(f"\n{full_name}, your identification number has successfully being generated.\nIndentifiction_Number = {rand_code}.")

#     def inform_user():
#         print(f"{full_name}, please use  both your identification number and prospective user name to  create a valid account")
#         print("\n\nFormat: user_name =  Grade_Repository('user_name', idn_num)\nidn is  identificcation")
#     inform_user()


# identity_generator()

# def inform_user():
#     print("")

list_stu = []


class teacher_account:
    def __init__(self, user_name, idn_num):
        # print("Successfully created👍")
        self.User_Name = user_name
        self.Idn_num = idn_num

    def add_stu(self):
        global list_stu
        student_num = 0
        student_name = input("\nPlease enter a student  name.\n")
        list_stu.append(student_name)
        student_num += 1
        print(
            f"\n{student_name} is  added to  the list.\n{list_stu}\nStudent_Number = {student_num}")
        prom_again = input("\nAgain?\nY or N\n")
        if prom_again == "Y":
            self.add_stu()
        else:
            return

    def student_list(self):
        print(list_stu)


Sekou = Grade_Repository("sekou", 145)
print(Sekou.add_stu())
