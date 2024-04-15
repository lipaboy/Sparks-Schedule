# import time
# from multiprocessing import Process
#
# def print_word(word):
#     time.sleep(1)
#     print('hello,', word)
#
# if __name__ == '__main__':
#     p1 = Process(target=print_word, args=('bob',), daemon=True)
#     p2 = Process(target=print_word, args=('alice',), daemon=True)
#     p1.start()
#     p2.start()
#     p1.join()
#     p2.join()
#




# Import all the necessary libraries
from tkinter import *
import time
import threading

# Define the tkinter instance
win = Tk()

# Define the size of the tkinter frame
win.geometry("700x400")


# Define the function to start the thread
def thread_fun():
    label.config(text="You can Click the button or Wait")
    time.sleep(5)
    label.config(text="5 seconds Up!")


label = Label(win)
label.pack(pady=20)
# Create button
b1 = Button(win, text="Start", command=threading.Thread(target=thread_fun).start())
b1.pack(pady=20)

win.mainloop()