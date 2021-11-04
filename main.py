from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import threading
import convertapi
import os

# Constants #
FILE = ''
DATA = [('docx', '*.docx')]
FILE_TYPE = [("Pages Files", "*.pages")]
THREADS = []
index = 0


def convert():
    if FILE != '':
        select_button.place_forget()
        convert_button.place_forget()

        animate(loading_images)

        try:
            convertapi.api_secret = 'VcVYRHh3U43llHk0'
            convert_file = convertapi.convert('docx', {
                'File': FILE
            }, from_format='pages')
        except:
            messagebox.showerror('Error', 'Something went wrong, please try again.')
        else:
            while True:
                save_path = ask_save()
                try:
                    convert_file.save_files(save_path)
                except FileNotFoundError:
                    check_cancel = messagebox.askyesno('Cancel?', 'Are you sure you want to cancel?')
                    if check_cancel:
                        break
                else:
                    convert_file.save_files(save_path)
                    break
    else:
        messagebox.showerror("Error", "Please select a file to convert.")


def choose_file():
    global FILE
    file = filedialog.askopenfile(filetypes=FILE_TYPE)
    if file:
        FILE = file.name


def animate(images):
    global index
    global loading_image
    loading.pack()

    if loading_image is None:
        loading_image = loading.create_image(200, 300, image=images[index])
    else:
        loading.itemconfig(loading_image, image=images[index])

    if index == 3:
        index = 0
    else:
        index += 1

    for thread in THREADS:
        if threading.Thread.is_alive(thread):
            loading.after(200, lambda: animate(images))
        else:
            loading.pack_forget()
            THREADS.remove(thread)
            home_screen()


def ask_save():
    location = filedialog.asksaveasfilename(filetypes=DATA, defaultextension=('docx', '*.docx'),
                                            initialfile=os.path.split(FILE)[1].split('.')[0],
                                            initialdir=os.path.split(FILE)[0])
    return location


# Window Set Up
root = Tk()
root.title('Monkey\'s Converter')
root.minsize(400, 600)
root.config(bg='#FFFD95')
root.resizable(height=False, width=False)
root.iconbitmap('images/monkey.ico')


# Thread Set Up
def start_thread():
    thread = threading.Thread(target=convert)
    thread.start()
    THREADS.append(thread)


# Image Dependencies
button_images = [os.path.join('images', 'button_select.png'),
                 os.path.join('images', 'button_convert.png')]
loading_images = [PhotoImage(file=f"images/loading/{file}") for file in os.listdir('images/loading')]
select_file = PhotoImage(file=button_images[0])
convert_image = PhotoImage(file=button_images[1])


# GUI Elements
loading = Canvas(width=400, height=600, bg='#FFFD95', highlightthickness=0)
loading_image = None

select_button = Button(bg='#FFFD95', bd=0, activebackground='#FFFD95', text='Search', command=choose_file)
select_button.config(image=select_file)

convert_button = Button(bg='#FFFD95', bd=0, activebackground='#FFFD95', command=start_thread)
convert_button.config(image=convert_image)


def home_screen():
    select_button.place(x=115, y=200)
    convert_button.place(x=115, y=325)


if __name__ == '__main__':
    home_screen()
    root.mainloop()
