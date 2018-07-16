import random

class NumberGenerator:
    def __init__(self):
        self.numbers = []
        self.first = []
        self.second = []
        self.third = []
        self.fourth = []
        self.fifth = []
        self.sixth = []

    def generate_numbers(self):
        self.numbers = random.sample(range(1, 91), 78)

    def fill_lists(self):
        for i in range(0, len(self.numbers)):
            if int(i / 13) == 0:
                self.first.append(self.numbers[i])
            elif int(i / 13) == 1:
                self.second.append(self.numbers[i])
            elif int(i / 13) == 2:
                self.third.append(self.numbers[i])
            elif int(i / 13) == 3:
                self.fourth.append(self.numbers[i])
            elif int(i / 13) == 4:
                self.fifth.append(self.numbers[i])
            elif int(i / 13) == 5:
                self.sixth.append(self.numbers[i])

    def view_numbers(self, num_list):
        print(num_list)