import logging

# DEBUG(10), INFO(20), WARNING(default level)(30), ERROR(40), CRITICAL(50)

logging.basicConfig(filename='test.log', level=logging.INFO,
                    format='%(asctime)s :%(levelname)s: :%(message)s')

def add(x, y):
    return x + y

def subtract(x, y):
    return x - y

def multiply(x, y):
    return x * y

def divide(x, y):
    return x / y

num_1 = 10
num_2 = 5

add_result = add(num_1, num_2)
logging.debug('Add: {} + {} = {}'.format(num_1, num_2, add_result))

sub_result = subtract(num_1, num_2)
logging.debug('Subtract: {} - {} = {}'.format(num_1, num_2, sub_result))

mul_result = multiply(num_1, num_2)
logging.debug('Multiply: {} * {} = {}'.format(num_1, num_2, mul_result))

div_result = divide(num_1, num_2)
logging.debug('Divide: {} / {} = {}'.format(num_1, num_2, div_result))

class Employee:
    """A sample Employee class"""

    def __init__(self, first, last):
        self.first = first
        self.last = last

        logging.debug('Created Employee: {} - {}'.format(self.fullname, self.email))

    @property
    def email(self):
        return '{}.{}'.format(self.first, self.last)

    @property
    def fullname(self):
        return '{} {}'.format(self.first, self.last)

empl_1 = Employee('John', 'Smith')
empl_2 = Employee('Corey', 'Schafer')
empl_3 = Employee('Jane', 'Doe')

bankAccountNumber = '0000000016'

print(bankAccountNumber[-4:])