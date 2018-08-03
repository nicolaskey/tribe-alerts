class Person(object):

    def __init__(self, name: str, ph_num: str, dues: str):
        self.name = name
        self.ph_num = ph_num.replace('$', '')
        self.dues = dues

        try:
            self.ph_num = int(self.ph_num)
        except ValueError:
            pass


    def __str__(self) -> str:
        return '{} owes {}. Their phone number is {}.'.format(self.name, self.dues, self.ph_num)

    __repr__ = __str__