import Tkinter as tk
import random


def main_menu(window):
    frame = tk.Frame(window).pack()
    tk.Label(frame, text="Welcome please choose a function below.").pack(pady=5)
    tk.Button(frame, text="LOLO Certificate Creator", command=lolo_creator).pack(pady=5)
    tk.Button(frame, text="Access Creator", command=access_creator).pack(pady=5)
    tk.Button(frame, text="Access Remover", command=access_remover).pack(pady=5)
    tk.Button(frame, text="Password Resetter", command=password_reset).pack(pady=5)
    tk.Button(frame, text="Reseller Onboarder (BETA)",command=reseller_onboard).pack(pady=5)
    tk.Button(frame, text="Quit", command=window.quit).pack(pady=5)

def lolo_creator():
    print("LOLO")

def access_creator():
    print("Access")

def access_remover():
    print("Remove")

def password_reset():
    print("Reset")

def reseller_onboard():
    print("Reseller")


window = tk.Tk()
window.title("IT Auto Scripts")
window.geometry("300x250+30+30")
main_menu(window)
window.mainloop()