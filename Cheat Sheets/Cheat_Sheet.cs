using System;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Basics: Printing and Input
        Console.WriteLine("Hello, World!");

        Console.Write("Enter your name: ");
        string name = Console.ReadLine();
        Console.WriteLine("Hello, " + name);

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
        double division = (double)x / y;
        int remainder = x % y;

        // Lists and Indexing
        List<string> fruits = new List<string> { "apple", "banana", "cherry" };
        string firstFruit = fruits[0];
        fruits.Add("orange");
        int numFruits = fruits.Count;

        // Control Structures: If-Else
        if (age >= 18)
        {
            Console.WriteLine("You are an adult.");
        }
        else
        {
            Console.WriteLine("You are a minor.");
        }

        // Loops: For and While
        for (int i = 0; i < 5; i++)
        {
            Console.WriteLine("Iteration " + i);
        }

        int counter = 0;
        while (counter < 5)
        {
            Console.WriteLine("While loop iteration " + counter);
            counter++;
        }

        // Functions
        string Greet(string n) => "Hello, " + n;
        string message = Greet("Alice");

        // Classes and Objects
        class Dog
        {
            public string Name { get; set; }
            public string Breed { get; set; }

            public string Bark()
            {
                return Name + " barks!";
            }
        }

        Dog myDog = new Dog
        {
            Name = "Buddy",
            Breed = "Golden Retriever"
        };
        string dogSound = myDog.Bark();

        // Lambda Functions in C#
        Func<int, int, int> add = (x, y) => x + y;
        int sum = add(3, 4);  // sum is 7

        // Printing Results
        Console.WriteLine("Variables and Data Types:");
        Console.WriteLine(age + ", " + height + ", " + isStudent);
        Console.WriteLine("\nBasic Arithmetic:");
        Console.WriteLine(sumResult + ", " + difference + ", " + product + ", " + division + ", " + remainder);
        Console.WriteLine("\nLists and Indexing:");
        Console.WriteLine(string.Join(", ", fruits));
        Console.WriteLine("\nControl Structures:");
        Console.WriteLine("You are an adult.");
        Console.WriteLine("\nLoops:");
        Console.WriteLine("For Loop:");
        for (int i = 0; i < 5; i++)
        {
            Console.WriteLine("Iteration " + i);
        }
        Console.WriteLine("While Loop:");
        counter = 0;
        while (counter < 5)
        {
            Console.WriteLine("While loop iteration " + counter);
            counter++;
        }
        Console.WriteLine("\nFunctions:");
        Console.WriteLine(message);
        Console.WriteLine("\nClasses and Objects:");
        Console.WriteLine(dogSound);
        Console.WriteLine("\nLambda Functions in C#:");
        Console.WriteLine("Sum: " + sum);
    }
}
