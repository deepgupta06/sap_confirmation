from tkinter import Tk
from tkinter import Label
from tkinter import Entry
from tkinter import Button
from cryptography.fernet import Fernet

root = Tk()

label_input = Label(root,text="Date:")
label_input.grid(row=0,column=0)

entry = Entry(root)
entry.grid(row=0, column=1)

def generate():
    key = Fernet.generate_key()
    fernet = Fernet(key)
            
    date_string  = entry.get()
    encmsg = fernet.encrypt(date_string.encode())
    
    with open("v.key", "wb") as keygen:
        keygen.write(key)
    print(key)
    with open("v.lnc", "wb") as file:
        file.write(encmsg)
        
    done_label = Label(root, text="Done!!;-)")
    done_label.grid(row=1,column=0)
        
btn  = Button(root, text = "Generate Key", command = generate)
btn.grid(row=0, column=2)

root.mainloop()