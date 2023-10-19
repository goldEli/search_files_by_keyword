import threading


def say_hello(name):
    for i in range(5):
        print(f"hello, {name}")


thread = threading.Thread(target=say_hello, args=("world",))

thread.start()

for i in range(5):
    print("Main thread is running...")

thread.join()

print("Main thread finished")
