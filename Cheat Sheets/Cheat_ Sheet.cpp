#include <iostream>
#include <string>
#include <vector>

using namespace std;

int main() {
    // Basics: Printing and Input
    cout << "Hello, World!" << endl;

    string name;
    cout << "Enter your name: ";
    getline(cin, name);
    cout << "Hello, " << name << endl;

    // Variables and Data Types
    int age = 25;
    double height = 5.9;
    bool isStudent = true;

    // Basic Arithmetic
    int x = 10;
    int y = 5;
    int sumResult = x + y;
    int difference = x - y;
    int product = x * y;
    double division = static_cast<double>(x) / y;
    int remainder = x % y;

    // Vectors and Indexing
    vector<string> fruits = {"apple", "banana", "cherry"};
    string firstFruit = fruits[0];
    fruits.push_back("orange");
    int numFruits = fruits.size();

    // Control Structures: If-Else
    if (age >= 18) {
        cout << "You are an adult." << endl;
    } else {
        cout << "You are a minor." << endl;
    }

    // Loops: For and While
    for (int i = 0; i < 5; ++i) {
        cout << "Iteration " << i << endl;
    }

    int counter = 0;
    while (counter < 5) {
        cout << "While loop iteration " << counter << endl;
        ++counter;
    }

    // Functions
    auto greet = [](const string& name) -> string {
        return "Hello, " + name;
    };
    string message = greet("Alice");

    // Classes and Objects
    class Dog {
    public:
        string name;
        string breed;

        string bark() {
            return name + " barks!";
        }
    };

    Dog myDog;
    myDog.name = "Buddy";
    myDog.breed = "Golden Retriever";
    string dogSound = myDog.bark();

    // Lambda Functions in C++
    auto add = [](int x, int y) { return x + y; };
    int sum = add(3, 4);  // sum is 7

    // Printing Results
    cout << "Variables and Data Types:" << endl;
    cout << age << ", " << height << ", " << isStudent << endl;
    cout << "\nBasic Arithmetic:" << endl;
    cout << sumResult << ", " << difference << ", " << product << ", " << division << ", " << remainder << endl;
    cout << "\nVectors and Indexing:" << endl;
    for (const auto& fruit : fruits) {
        cout << fruit << ", ";
    }
    cout << "\nControl Structures:" << endl;
    cout << "You are an adult." << endl;
    cout << "\nLoops:" << endl;
    cout << "For Loop:" << endl;
    for (int i = 0; i < 5; ++i) {
        cout << "Iteration " << i << endl;
    }
    cout << "While Loop:" << endl;
    counter = 0;
    while (counter < 5) {
        cout << "While loop iteration " << counter << endl;
        ++counter;
    }
    cout << "\nFunctions:" << endl;
    cout << message << endl;
    cout << "\nClasses and Objects:" << endl;
    cout << dogSound << endl;
    cout << "\nLambda Functions in C++:" << endl;
    cout << "Sum: " << sum << endl;

    return 0;
}
