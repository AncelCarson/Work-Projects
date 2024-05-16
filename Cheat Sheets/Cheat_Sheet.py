# Basics: Printing and Input
print("Hello, World!")

name = input("Enter your name: ")
print("Hello, {}".format(name))  # String formatting with .format()

# Variables and Data Types
age = 25
height = 5.9  # Floating-point number
is_student = True  # Boolean

# Basic Arithmetic
x = 10
y = 5
sum_result = x + y
difference = x - y
product = x * y
division = x / y
remainder = x % y

# Lists and Indexing
fruits = ["apple", "banana", "cherry"]
first_fruit = fruits[0]
fruits.append("orange")
num_fruits = len(fruits)

# Control Structures: If-Else
if age >= 18:
    print("You are an adult.")
else:
    print("You are a minor.")

# Loops: For and While
for i in range(5):
    print("Iteration {}".format(i))  # String formatting with .format()

counter = 0
while counter < 5:
    print("While loop iteration {}".format(counter))  # String formatting with .format()
    counter += 1

# Functions
def greet(name):
    return "Hello, {}".format(name)  # String formatting with .format()

message = greet("Alice")

# Classes and Objects
class Dog:
    def __init__(self, name, breed):
        self.name = name
        self.breed = breed

    def bark(self):
        return "{} barks!".format(self.name)  # String formatting with .format()

my_dog = Dog("Buddy", "Golden Retriever")
dog_sound = my_dog.bark()

# Built-in Functions
num_list = [3, 1, 4, 1, 5, 9, 2, 6, 5, 3, 5]
min_value = min(num_list)
max_value = max(num_list)
sum_values = sum(num_list)

# File Handling
with open("sample.txt", "w") as file:
    file.write("This is a sample file.")

with open("sample.txt", "r") as file:
    file_contents = file.read()

# Exception Handling
try:
    result = 10 / 0
except ZeroDivisionError:
    result = "Error: Division by zero"

# Modules and Libraries
import math
sqrt_2 = math.sqrt(2)

# Lambda Functions in Python
add = lambda x, y: x + y
sum_result = add(3, 4)  # sum_result is 7

# Printing Results
print("Variables and Data Types:")
print(age, height, is_student)
print("\nBasic Arithmetic:")
print(sum_result, difference, product, division, remainder)
print("\nLists and Indexing:")
print(fruits, first_fruit, num_fruits)
print("\nControl Structures:")
print("You are an adult." if age >= 18 else "You are a minor.")
print("\nLoops:")
print("For Loop:")
for i in range(5):
    print("Iteration {}".format(i))  # String formatting with .format()
print("While Loop:")
counter = 0
while counter < 5:
    print("While loop iteration {}".format(counter))  # String formatting with .format()
    counter += 1
print("\nFunctions:")
print(message)
print("\nClasses and Objects:")
print(dog_sound)
print("\nBuilt-in Functions:")
print(min_value, max_value, sum_values)
print("\nFile Handling:")
print(file_contents)
print("\nException Handling:")
print(result)
print("\nModules and Libraries:")
print(sqrt_2)
print("\nLambda Functions in Python:")
print("Sum:", sum_result)
