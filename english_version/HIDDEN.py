import random
import sqlite3
import webbrowser
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import win32com.client as win32
from cryptography.fernet import Fernet

# Creating connection with the database used
conn = sqlite3.connect("users.db")
cursor = conn.cursor()

# This string will take the place of the optional data when the user do not give one
EMPTY = "EMPTY"

# this string will make it explicit if the user has just entered the application
FIRST_ENTRY = True

# Parameters for meny functions
USERNAME = ""
PASSWORD = ""
EMAIL = ""
CODE = ""
SESSION_ID = ""
TREE_VIEW = ""

# Start the user interface program
root = Tk()
root.title("HIDDEN")
root.minsize(250, 0)
root.iconbitmap("padlock.ico")
root.resizable(width=False, height=False)

# Main labels
global page_label
global Menu_Frame


# Interface page-functions


def menu_options():
    global FIRST_ENTRY

    load_menu(FIRST_ENTRY)
    load_page(FIRST_ENTRY)
    FIRST_ENTRY = False

    login_button = Button(Menu_Frame, text="CREATE ACCOUNT", command=register_page)
    login_button.grid(row=0, column=0)

    login_button = Button(Menu_Frame, text="LOGIN", command=login_page)
    login_button.grid(row=0, column=1)

    # Menu main page
    message = """
    This app aims to save passwords securely! Any saved password will be encrypted 
    and  can only be decoded by the  application  after user  access with  username
    and password, giving access  only to the data that this user has previously saved. 
    in order that the application works correctly, it is important to:

         -Be running on the Windows operating system as it has not been tested on
          other systems;

         -Make sure that 'Microsoft Office Outlook' is installed the device used. 
          Otherwise, the attempting to use the application results in an error.
    """

    intro_label = Label(page_label, text=message, justify=LEFT)
    visit_pages_label = Label(page_label, justify=LEFT)
    github_button = Button(visit_pages_label, text="VISIT AUTHOR GITHUB",
                           command=lambda: webbrowser.open_new("https://github.com/Luerson/Passwords-Manager"))
    image_page_button = Button(visit_pages_label, text="Pixel perfect - Flaticon IMAGES",
                               command=lambda: webbrowser.open_new("https://www.flaticon.com/free-icons/password"))

    intro_label.pack()
    visit_pages_label.pack()
    github_button.grid(row=0, column=0, pady=5)
    image_page_button.grid(row=0, column=1, pady=5)


def register_page():
    load_page()

    # CREATING PAGE ITEMS

    # Creating user prompt
    request_username = Label(page_label, text="choose a username")
    request_password = Label(page_label, text="choose a password")
    request_confirmation = Label(page_label, text="confirm password")
    request_email = Label(page_label, text="contact email")

    # Creating user inputs
    get_username = Entry(page_label)
    get_password = Entry(page_label, show="*")
    get_confirmation = Entry(page_label, show="*")
    get_email = Entry(page_label)

    # Creating info buttons
    info_username = Button(page_label, text="?", command=lambda: info_function("username"))
    info_password = Button(page_label, text="?", command=lambda: info_function("password"))
    info_confirmation = Button(page_label, text="?", command=lambda: info_function("confirmation"))
    info_email = Button(page_label, text="?", command=lambda: info_function("email"))

    # Creating registration button
    register_button = Button(page_label, text="REGISTER",
                             command=lambda: valid_register_inputs(get_username.get(), get_password.get(),
                                                                   get_confirmation.get(), get_email.get()))

    # PLACING PAGE ITEMS

    # Placing prompt
    request_username.grid(row=0, column=0, sticky=W)
    request_password.grid(row=2, column=0, sticky=W)
    request_confirmation.grid(row=4, column=0, sticky=W)
    request_email.grid(row=6, column=0, sticky=W)

    # Placing user inputs
    get_username.grid(row=1, column=0)
    get_password.grid(row=3, column=0)
    get_confirmation.grid(row=5, column=0)
    get_email.grid(row=7, column=0)

    # Placing info buttons
    info_username.grid(row=1, column=1, padx=2)
    info_password.grid(row=3, column=1, padx=2)
    info_confirmation.grid(row=5, column=1, padx=2)
    info_email.grid(row=7, column=1, padx=2)

    # Placing registration button
    register_button.grid(row=8, column=0, columnspan=2, pady=30)


def login_page():
    load_page()

    # CREATING PAGE ITEMS

    # Creating user prompt
    request_username = Label(page_label, text="Type your username")
    request_password = Label(page_label, text="Type your password")

    # Creating user inputs
    get_username = Entry(page_label, width=26)
    get_password = Entry(page_label, show="*", width=26)

    # Creating login button and change password button
    login_button = Button(page_label, text="ENTER",
                          command=lambda: valid_login_inputs(get_username.get(), get_password.get()))
    change_password_button = Button(page_label, text="FORGOT PASSWORD",
                                    command=lambda: change_password(get_username.get()))

    # PLACING PAGE ITEMS

    # Placing prompt
    request_username.grid(row=0, column=0, columnspan=2, sticky=W)
    request_password.grid(row=2, column=0, columnspan=2, sticky=W)

    # Placing user inputs
    get_username.grid(row=1, column=0, columnspan=2)
    get_password.grid(row=3, column=0, columnspan=2)

    # Placing login button
    change_password_button.grid(row=4, column=0, pady=20)
    login_button.grid(row=4, column=1, pady=20)


# This function will delete the current page's content and create a label to receive the new page content
def load_page(first_entry=False):
    global page_label

    if not first_entry:
        page_label.pack_forget()

    page_label = Label(root)
    page_label.pack()

    return


# This function will delete the current menu's content and create a label to receive the new menu content
def load_menu(first_entry=False):
    global Menu_Frame

    if not first_entry:
        Menu_Frame.pack_forget()

    Menu_Frame = Label(root)
    Menu_Frame.pack()

    return


# Helper functions


# Now begins the part that refers to the registration process


# This function will take the user inputs for registration and will verify if they are valid
def valid_register_inputs(username, password, confirmation, email):
    if not username:
        messagebox.showerror("No user", "Type a username to proceed")
    elif username == "" or name_exists(username, "users") != 0:
        messagebox.showwarning("Unavailable", f"Username '{username}' is not available, please try another")
    elif not password:
        messagebox.showerror("No password", "Type a password to proceed")
    elif not confirmation:
        messagebox.showerror("No confirmation", "Confirm your password to proceed")
    elif not email:
        messagebox.showerror("No email", "Type an contact email")
    elif password != confirmation:
        messagebox.showerror("Incompatible", "Password and confirmation are not the same. Please try again")
    else:
        if not valid_email_format(email):
            messagebox.showerror("email not valid", "please, provide an email account in"
                                                    " the format 'exemplo@exemplo.com'")
        else:
            global USERNAME
            global PASSWORD
            global EMAIL
            global CODE
            USERNAME = username
            PASSWORD = password
            EMAIL = email

            CODE = str(random.randint(100000, 999999))
            title = "Validation Password"
            message = f"""<p>Hello, <b>{username}</b>,
                                      this email contains the security code to validate your email account</p>
                                      <p>Validation code: <b>{CODE}</b></p>
                                      <p>If you have no idea of what it is about, please ignore this message.</p>
                                      <p>Thanks for your attention, from Hidden Automatic Program!</p>"""

            read = messagebox.showinfo("Preparing code", f"A code is being sent to {EMAIL}!"
                                                         " This may take a while, please click 'ok'"
                                                         ", or close this window, and wait!")

            if read:
                if send_email(EMAIL, title, message):
                    received = messagebox.askyesno("Code sent", "Please, confirm if you received the code")

                    if received == 0:
                        messagebox.showwarning("Email not valid",
                                               "Please, check if the typed email is correct")
                    else:
                        type_code_window("register")


# This function will check if the typed email is in the following format: example@example.com
def valid_email_format(email):
    valid_format = True

    if '@' not in email or '.com' not in email:
        valid_format = False
    elif email.count('@') > 1 or email.count('.com') > 1:
        valid_format = False
    else:
        email_parts = email.split('@')
        email_parts[1] = email_parts[1].split('.com')

        if email_parts[0] == "" or email_parts[1][0] == "":
            valid_format = False

    return valid_format


# This function executes the user registration
def register(username, password, email, key):
    # Execute the registration
    cursor.execute("INSERT INTO users (username, password, email, key) VALUES (?, ?, ?, ?)",
                   (username, password, email, key))
    conn.commit()

    login_page()

    return


# Now begins the part that refers to the login


def valid_login_inputs(username, password):
    user_id = name_exists(username, "users")

    if username != "" and user_id != 0:
        valid_password = decrypt_function(user_id, "password", "users")
        valid_password = valid_password

        if password == valid_password:
            global EMAIL
            global SESSION_ID

            SESSION_ID = user_id

            EMAIL = decrypt_function(user_id, "email", "users")
            EMAIL = EMAIL

            login_main_page(get_data_list())
            return

    messagebox.showerror("Entry not valid", "username or password not valid")
    return


def login_main_page(passwords_list):
    load_menu()
    consult_page(passwords_list)

    # Login menu
    login_button = Button(Menu_Frame, text="LOGOUT", command=logout)
    login_button.grid(row=0, column=0)


def consult_page(passwords_list):
    global TREE_VIEW

    load_page()

    # search row defined
    search_label = Label(page_label)
    search_input = Entry(search_label, width=25)
    search_input.insert(0, "Type the password's title")
    search_button = Button(search_label, text="SEARCH", command=lambda: search_function(search_input.get()))
    see_all_button = Button(search_label, text="SEE ALL", command=lambda: consult_page(get_data_list()))

    # search row grid
    search_label.grid(row=0, column=0, pady=10, sticky=W)
    search_input.grid(row=0, column=0)
    search_button.grid(row=0, column=1)
    see_all_button.grid(row=0, column=2)

    # delete add row defined
    add_delete_label = Label(page_label)
    add_button = Button(add_delete_label, text="ADD PASSWORD", command=add_function)
    delete_button = Button(add_delete_label, text="DELETE PASSWORD", command=lambda: delete_function(TREE_VIEW))
    backup_button = Button(add_delete_label, text="BACKUP PASSWORDS TO EMAIL",
                           command=lambda: create_passwords_email(TREE_VIEW))

    # delete add row grid
    add_delete_label.grid(row=1, column=0, sticky=W)
    add_button.grid(row=0, column=0, padx=6)
    delete_button.grid(row=0, column=1, padx=5)
    backup_button.grid(row=0, column=2, padx=5)

    # data row
    data_label = Label(page_label)

    data_label.grid(row=2, column=0, columnspan=2)

    TREE_VIEW = ttk.Treeview(data_label, columns=("1", "2"), show="headings", height=5)
    TREE_VIEW.pack()

    TREE_VIEW.heading("1", text="TITLE")
    TREE_VIEW.heading("2", text="PASSWORD")

    for row in passwords_list:
        TREE_VIEW.insert('', 'end', values=row)

    return


def logout():
    global SESSION_ID

    SESSION_ID = ""
    menu_options()


def search_function(search_input):
    data_list = get_data_list()
    search_data_list = []

    for data in data_list:
        if search_input.lower() in data[0][0].lower():
            search_data_list.append(data)

    consult_page(search_data_list)


def add_function():
    top = Toplevel()
    top.title("HIDDEN")
    top.minsize(275, 0)
    top.iconbitmap("padlock.ico")
    top.resizable(width=False, height=False)

    title_label = Label(top, text="Choose a title for your password", justify=LEFT)
    password_label = Label(top, text="type the password you want to save", justify=LEFT)

    get_title = Entry(top, width=26)
    get_password = Entry(top, width=26)

    add_data_button = Button(top, text="SAVE", command=lambda: [save_data(get_title.get(), get_password.get())])

    title_label.pack()
    get_title.pack()
    password_label.pack()
    get_password.pack()
    add_data_button.pack(pady=5)
    return


def save_data(title, password):
    if name_exists(title, "passwords") != 0:
        messagebox.showerror("Not valid", "Different passwords can not have the same title")
    elif " " in title or " " in password:
        messagebox.showerror("Not valid", "The titles can not have white space")
    elif title == "" or password == "":
        messagebox.showerror("Not valid", "Title and password can not be empty")
    else:
        cursor.execute(f"SELECT key FROM users WHERE id = {SESSION_ID}")
        key = cursor.fetchall()
        key = key[0][0]

        encrypted_password = encrypt_function(password, key)

        cursor.execute(f"INSERT INTO passwords VALUES (?, ?, ?)", (SESSION_ID, title, encrypted_password))
        conn.commit()

        login_main_page(get_data_list())


def delete_function(tree_view):
    proceed = messagebox.askyesno("Delete?", "Are you sure you want to delete this password?")

    if proceed == 1:
        selected = tree_view.selection()

        if not selected:
            messagebox.showwarning("No selection",
                                   "Choose the password you want, then click in 'DELETE PASSWORD'!")
        else:
            values = tree_view.item(selected, "values")
            tree_view.delete(selected)
            cursor.execute("DELETE FROM passwords WHERE user_id = ? AND title = (?)", (SESSION_ID, values[0]))
            conn.commit()

    return


def change_password(username):
    if name_exists(username, "users") == 0:
        messagebox.showerror("Not valid", "please, type a valid username to proceed")
        return

    global USERNAME
    global EMAIL
    global CODE
    global SESSION_ID

    SESSION_ID = list(cursor.execute("SELECT id FROM users WHERE username = ?", (username,)))
    if not SESSION_ID:
        messagebox.showerror("No username", "type a username to proceed")
        return

    SESSION_ID = SESSION_ID[0][0]
    USERNAME = username
    EMAIL = decrypt_function(SESSION_ID, "email", "users")

    CODE = str(random.randint(100000, 999999))
    title = "Verification code"
    message = f"""<p>Hello, <b>{username}</b>,
                                          this message contains the security code to validate validate your account.</p>
                                          <p>verification code: <b>{CODE}</b></p>
                                          <p>If you have no idea what this is about, please ignore this message.</p>
                                          <p>Thanks for you attention, from Hidden Automatic Program!</p>"""

    read = messagebox.showinfo("sending code", f"a code is being sent to {EMAIL}!"
                                               " This may take a while, please click in 'ok'"
                                               " ,or close this window, and wait")

    if read:
        if send_email(EMAIL, title, message):
            received = messagebox.askyesno("Code sent", "Please, confirm if you received the code")

            if received == 0:
                messagebox.showerror("ERROR", "Ops! Something went wrong, please try again")
            else:
                type_code_window("change_password")

    return


def change_execute():
    load_page()

    password_label = Label(page_label, text="Type the new password")
    confirm_label = Label(page_label, text="Confirm the new password")

    get_new_password = Entry(page_label, show="*")
    get_confirmation = Entry(page_label, show="*")

    change_button = Button(page_label, text="CHANGE PASSWORD",
                           command=lambda: register_new_password(get_new_password.get()))

    password_label.pack()
    get_new_password.pack()
    confirm_label.pack()
    get_confirmation.pack()
    change_button.pack()
    return


def register_new_password(new_password):
    global SESSION_ID

    cursor.execute(f"SELECT key FROM users WHERE id = {SESSION_ID}")
    key = cursor.fetchall()
    key = key[0][0]

    new_password = encrypt_function(new_password, key)

    cursor.execute(f"UPDATE users SET password = ? WHERE id = ?", (new_password, SESSION_ID))
    conn.commit()

    login_page()
    return


def create_passwords_email(tree_view):
    global USERNAME
    global EMAIL

    tree_list = tree_view.get_children()

    message = """
    <style>
        table {
          font-family: arial, sans-serif;
          border-collapse: collapse;
          width: 100%;
        }

        td, th {
          border: 1px solid #dddddd;
          text-align: left;
          padding: 8px;
        }

        tr:nth-child(even) {
          background-color: #dddddd;
        }
        </style>
        """
    message += f"""
    <p>Hello {USERNAME}, here follows the list of the passwords saved in your database:</p>
    <table>
        <tr>
            <th>TITLE</th>
            <th>PASSWORD</th>
        </tr>
    """

    for row in tree_list:
        title = tree_view.item(row, "values")[0]
        password = tree_view.item(row, "values")[1]
        message += f"""
        <tr>
            <td>{title}</td>
            <td>{password}</td>
        </tr>
        """

    message += """</table>"""

    title = "passwords HIDDEN"

    messagebox.showinfo("Sending", f"Data is being prepared to be sent to {EMAIL}."
                                   " The process may take a while, please wait.")
    send_email(EMAIL, title, message)

    return


# Functions that are used through the entire program


def type_code_window(operation_type):
    top = Toplevel()
    top.title("HIDDEN")
    top.iconbitmap("padlock.ico")
    top.resizable(width=False, height=False)

    prompt = Label(top, text="Type the code", justify=LEFT)
    get_email_code = Entry(top)
    send_button = Button(top, text="SEND",
                         command=lambda: [check_code(get_email_code.get(), operation_type), top.destroy()])

    prompt.grid(row=0, column=0)
    get_email_code.grid(row=1, column=0, padx=5)
    send_button.grid(row=1, column=1, pady=10)


def check_code(typed_code, operation_type):
    if not typed_code.isdigit():
        messagebox.showwarning("Only numbers", "The code must have only numbers")
    elif int(typed_code) != int(CODE):
        messagebox.showerror("Wrong code", "The typed code is incorrect")
    elif operation_type == "register":
        proceed = messagebox.askyesno("Register?", "Your account is about to be registered."
                                                   " Please confirm if you want to proceed")

        if proceed:
            # Get the user key
            key = Fernet.generate_key()

            print(PASSWORD)

            encrypted_password = encrypt_function(PASSWORD, key)
            encrypted_email = encrypt_function(EMAIL, key)

            register(USERNAME, encrypted_password, encrypted_email, key)
    elif operation_type == "change_password":
        change_execute()


def get_data_list():
    global SESSION_ID

    cursor.execute(f"SELECT title FROM passwords WHERE user_id = {SESSION_ID} ORDER BY title")
    titles_list = cursor.fetchall()

    decrypted_passwords_list = decrypt_function(SESSION_ID, "password", "passwords")

    data_list = []

    for index in range(len(titles_list)):
        data_list.append([titles_list[index], decrypted_passwords_list[index]])

    return data_list


# This function send a message to an email
def send_email(recipient, title, message):
    try:
        outlook = win32.Dispatch("outlook.application")
    except:
        messagebox.showerror("Error outlook", "It was not possible to send the email. Please check if you have"
                                              " Microsoft Office Outlook installed on your device")
        return False

    email = outlook.CreateItem(0)

    email.To = f"{recipient}"
    email.subject = f"{title}"
    email.HTMLBody = f"{message}"

    email.Send()
    return True


# This function will show the required information in a messagebox
def info_function(info_case):
    if info_case == "username":
        messagebox.showinfo("Username", "You need to provide a user to be able to access the tools of the"
                                        " application")
    elif info_case == "password":
        messagebox.showinfo("Password?", "For security reasons, it is necessary that you have a password so that only "
                                         "you have access to the saved content.")
    elif info_case == "confirmation":
        messagebox.showinfo("Confirmation?", "To ensure that there was no typo in the previous entry, you need to "
                                             "re-enter the password you want.")
    elif info_case == "email":
        messagebox.showinfo("email?", "If you want to change your password, you will need an email address to make "
                                      "the transaction possible.")


# This function search for a specific username in the database
# If it is found, the function return the user_id, otherwise it will return zero
def name_exists(name, table):
    if table == "users":
        l_usernames = list(cursor.execute("SELECT username FROM users"))

        user_id = 1

        # Check if there is already another account with that username
        for username in l_usernames:

            # Username already exists
            if username[0] == name:
                return user_id

            user_id += 1

    elif table == "passwords":
        l_titles = list(cursor.execute(f"SELECT title FROM passwords WHERE user_id = {SESSION_ID}"))

        for title in l_titles:
            if title[0] == name:
                return title[0]

    # Available
    return 0


def encrypt_function(data, key):
    crypter = Fernet(key)

    encrypted_data = crypter.encrypt(bytes(data, 'utf-8'))

    return encrypted_data


def decrypt_function(user_id, column, table):
    list_key = list(cursor.execute(f"SELECT key FROM users WHERE id = {user_id}"))
    key = list_key[0][0]

    crypter = Fernet(key)

    if table == "users":
        cursor.execute(f"SELECT {column} FROM users WHERE id = {user_id}")
        encrypted_data = cursor.fetchall()

        binary_data = crypter.decrypt(encrypted_data[0][0])
        data = str(binary_data, 'utf-8')

        return data

    else:
        cursor.execute(f"SELECT {column} FROM passwords WHERE user_id = {user_id} ORDER BY title")
        list_data = cursor.fetchall()

        data_list = []

        for encrypted_data in list_data:
            binary_data = crypter.decrypt(encrypted_data[0])
            data = str(binary_data, 'utf-8')
            data_list.append(data)

        return data_list


menu_options()

root.mainloop()

conn.close()