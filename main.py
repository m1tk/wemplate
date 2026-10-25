import tkinter as tk

def main():
    root = tk.Tk()
    root.title("Hello World GUI")
    root.geometry("300x200")

    label = tk.Label(root, text="Hello, World!", font=("Arial", 16))
    label.pack(pady=50)

    button = tk.Button(root, text="Close", command=root.destroy)
    button.pack()

    root.mainloop()

if __name__ == "__main__":
    main()
