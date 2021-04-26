from tkinter import *
import tkinter as tk



def main():

    def authorizeUser(userName,password):
        print("hello")

    app = tk.Tk()
    screen_width = app.winfo_screenwidth()
    screen_height =app.winfo_screenheight()
    w = int(screen_width)
    h = int(screen_height)

    user = tk.StringVar()
    secret = tk.StringVar()

    app.geometry(f'{int(w / 4.5)}x{int(h / 4.5)}')
    app.title("LOGIN PAGE")
    app.configure(bg="#049fd9")


    name = Label(app, text="Username",fg="black",bg="#049fd9",font=(w/5)).place(x=int(w/20), y=int(h/20.5))

    password = Label(app, text="Password",fg="black",bg="#049fd9",font=(w/5)).place(x=int(w/20), y=int(h/10.5))


    e1 = Entry(app, textvariable=user).place(x=int(w/12), y=int(h/20))

    e2 = Entry(app, textvariable=secret).place(x=int(w/12), y=int(h/10.5))

    sbmitbtn = Button(app, text="Submit", activebackground="yellow", activeforeground="blue",width=int(w/200), height=int(h/500),command=lambda: authorizeUser(e1,e2)).place(x=int(w/12), y=int(h/8))






    app.mainloop()

if __name__ == "__main__":
    main()