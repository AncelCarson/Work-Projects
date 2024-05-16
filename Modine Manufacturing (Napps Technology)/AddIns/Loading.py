""" Ancel Carson
    Napps Technology Comporation
    10/14/2020
    Loading.py: Loading display attempt
"""
import time as t

dotLoop = 3
loadLoop = 4

def main():
    global loadLoop
    for load in range(loadLoop):
        loading()
    print("Loading Complete")
    input("Press ENTER to close window...")

def loading():
    global dotLoop
    print("Loading", end = "")
    t.sleep(1)
    for dot in range(dotLoop):
        print(".", end = "")
        t.sleep(1)
    print()

if __name__ == "__main__":
    main()

