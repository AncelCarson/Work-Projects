import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        // Basics: Printing and Input
        System.out.println("Hello, World!");

        System.out.print("Enter your name: ");
        java.util.Scanner scanner = new java.util.Scanner(System.in);
        String name = scanner.nextLine();
        System.out.println("Hello, " + name);

        // Variables and Data Types
        int age = 25;
        double height = 5.9;
        boolean isStudent = true;

        // Basic Arithmetic
        int x = 10;
        int y = 5;
        int sumResult = x + y;
        int difference = x - y;
        int product = x * y;
        double division = (double) x / y;
        int remainder = x % y;

        // Lists and Indexing
        List<String> fruits = new ArrayList<>();
        fruits.add("apple");
        fruits.add("banana");
        fruits.add("cherry");
        String firstFruit = fruits.get(0);
        fruits.add("orange");
        int numFruits = fruits.size();

        // Control Structures: If-Else
        if (age >= 18) {
            System.out.println("You are an adult.");
        } else {
            System.out.println("You are a minor.");
        }

        // Loops: For and While
        for (int i = 0; i < 5; i++) {
            System.out.println("Iteration " + i);
        }

        int counter = 0;
        while (counter < 5) {
            System.out.println("While loop iteration " + counter);
            counter++;
        }

        // Functions
        String message = greet("Alice");

        // Classes and Objects
        class Dog {
            String name;
            String breed;

            String bark() {
                return name + " barks!";
            }
        }

        Dog myDog = new Dog();
        myDog.name = "Buddy";
        myDog.breed = "Golden Retriever";
        String dogSound = myDog.bark();

        // Lambda Functions in Java
        BinaryOperator<Integer> add = (x, y) -> x + y;
        int sum = add.apply(3, 4);  // sum is 7

        // Printing Results
        System.out.println("Variables and Data Types:");
        System.out.println(age + ", " + height + ", " + isStudent);
        System.out.println("\nBasic Arithmetic:");
        System.out.println(sumResult + ", " + difference + ", " + product + ", " + division + ", " + remainder);
        System.out.println("\nLists and Indexing:");
        for (String fruit : fruits) {
            System.out.print(fruit + ", ");
        }
        System.out.println("\nControl Structures:");
        System.out.println("You are an adult.");
        System.out.println("\nLoops:");
        System.out.println("For Loop:");
        for (int i = 0; i < 5; i++) {
            System.out.println("Iteration " + i);
        }
        System.out.println("While Loop:");
        counter = 0;
        while (counter < 5) {
            System.out.println("While loop iteration " + counter);
            counter++;
        }
        System.out.println("\nFunctions:");
        System.out.println(message);
        System.out.println("\nClasses and Objects:");
        System.out.println(dogSound);
        System.out.println("\nLambda Functions in Java:");
        System.out.println("Sum: " + sum);
    }

    // Functions
    static String greet(String name) {
        return "Hello, " + name;
    }
}
