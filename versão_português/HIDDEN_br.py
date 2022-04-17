import random
import sqlite3
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

    login_button = Button(Menu_Frame, text="CRIAR CONTA", command=register_page)
    login_button.grid(row=0, column=0)

    login_button = Button(Menu_Frame, text="FAZER LOGIN", command=login_page)
    login_button.grid(row=0, column=1)

    # Menu main page
    message = """
    Este aplicativo tem o objetivo de salvar senhas de maneira segura! Qualquer senha salva será
    criptografada e apenas poderá ser decodificada pelo aplicativo após o acesso com o usuário
    e senha, dando acesso apenas aos dados que este usuário salvou previamente. A fim de que o 
    aplicativo funcione corretamente, é importante que:

         -Esteja sendo executado no sistema operacional Windows, visto que não foi testado em
          outros sistemas;

         -O Microsoft Office Outlook esteja instalado no aparelho utilizado. Do contrário, pode
          ocorrer de a tentativa de usar o aplicativo resultar em erro.
    """

    intro_label = Label(page_label, text=message, justify=LEFT)
    intro_label.pack()


def register_page():
    load_page()

    # CREATING PAGE ITEMS

    # Creating user prompt
    request_username = Label(page_label, text="escolha um usuário")
    request_password = Label(page_label, text="escolha sua senha")
    request_confirmation = Label(page_label, text="confirme sua senha")
    request_email = Label(page_label, text="email para contato")

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
    register_button = Button(page_label, text="REGISTRAR",
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
    request_username = Label(page_label, text="Digite seu usuário")
    request_password = Label(page_label, text="Digite sua senha")

    # Creating user inputs
    get_username = Entry(page_label, width=26)
    get_password = Entry(page_label, show="*", width=26)

    # Creating login button and change password button
    login_button = Button(page_label, text="ENTRAR",
                          command=lambda: valid_login_inputs(get_username.get(), get_password.get()))
    change_password_button = Button(page_label, text="ESQUECI A SENHA",
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
        messagebox.showerror("Sem usuário", "Digite um usuário para prosseguir")
    elif username == "" or name_exists(username, "users") != 0:
        messagebox.showwarning("Indisponível", f"Usuário '{username}' entá indisponível, por favor tente outro")
    elif not password:
        messagebox.showerror("Sem senha", "Digite uma senha para prosseguir")
    elif not confirmation:
        messagebox.showerror("Sem confirmação", "Confirme sua senha para prosseguir")
    elif not email:
        messagebox.showerror("Sem email", "Digite um email para contato")
    elif password != confirmation:
        messagebox.showerror("incompatibilidade", "A senha e a confirmação não são iguais. Tente novamente")
    else:
        if not valid_email_format(email):
            messagebox.showerror("email inválido", "por favor forneça um email para contato no"
                                                   " formato 'exemplo@exemplo.com'")
        else:
            global USERNAME
            global PASSWORD
            global EMAIL
            global CODE
            USERNAME = username
            PASSWORD = password
            EMAIL = email

            CODE = str(random.randint(100000, 999999))
            title = "Senha de verificação"
            message = f"""<p>Olá, <b>{username}</b>,
                                      este email tem o código de segurança para validar seu email.</p>
                                      <p>código de validação: <b>{CODE}</b></p>
                                      <p>Se você não sabe do que isso se trata, por favor ignore esta mensagem.</p>
                                      <p>Obrigado por sua atenção, de Programa Automático Python!</p>"""

            read = messagebox.showinfo("Enviando código", f"Um código de verificação está sendo enviado para {EMAIL}!"
                                                          " Isso pode demorar um pouco, por favor clique em 'ok'"
                                                          " e aguarde!")

            if read:
                if send_email(EMAIL, title, message):
                    received = messagebox.askyesno("código enviado", "Por favor, confirme se recebeu o código")

                    if received == 0:
                        messagebox.showwarning("Email inválido", "Por favor, verifique se o email digitado está correto")
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

    messagebox.showerror("Entrada inválida", "usuário ou senha incorreta")
    return


def login_main_page(passwords_list):
    load_menu()
    consult_page(passwords_list)

    # Login menu
    login_button = Button(Menu_Frame, text="SAIR DA CONTA", command=logout)
    login_button.grid(row=0, column=0)


def consult_page(passwords_list):
    global TREE_VIEW

    load_page()

    # search row defined
    search_label = Label(page_label)
    search_input = Entry(search_label, width=25)
    search_input.insert(0, "Digite o título da senha")
    search_button = Button(search_label, text="PROCURAR", command=lambda: search_function(search_input.get()))
    see_all_button = Button(search_label, text="VER TODAS", command=lambda: consult_page(get_data_list()))

    # search row grid
    search_label.grid(row=0, column=0, pady=10, sticky=W)
    search_input.grid(row=0, column=0)
    search_button.grid(row=0, column=1)
    see_all_button.grid(row=0, column=2)

    # delete add row defined
    add_delete_label = Label(page_label)
    add_button = Button(add_delete_label, text="ADICIONAR SENHA", command=add_function)
    delete_button = Button(add_delete_label, text="DELETAR SENHA", command=lambda: delete_function(TREE_VIEW))
    backup_button = Button(add_delete_label, text="ENVIAR SENHAS PARA EMAIL", command=lambda: create_passwords_email(TREE_VIEW))

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

    TREE_VIEW.heading("1", text="TÍTULO")
    TREE_VIEW.heading("2", text="SENHA")

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

    title_label = Label(top, text="Escolha um título para a senha", justify=LEFT)
    password_label = Label(top, text="digite a senha que quer salvar", justify=LEFT)

    get_title = Entry(top, width=26)
    get_password = Entry(top, width=26)

    add_data_button = Button(top, text="SALVAR", command=lambda: [save_data(get_title.get(), get_password.get())])

    title_label.pack()
    get_title.pack()
    password_label.pack()
    get_password.pack()
    add_data_button.pack(pady=5)
    return


def save_data(title, password):
    if name_exists(title, "passwords") != 0:
        messagebox.showerror("Indisponível", "Você não pode ter duas senhas com o mesmo título")
    elif " " in title or " " in password:
        messagebox.showerror("Inválido", "os títulos não podem ter espaço em branco")
    elif title == "" or password == "":
        messagebox.showerror("Inválido", "título e senha não podem ser vazios!")
    else:
        cursor.execute(f"SELECT key FROM users WHERE id = {SESSION_ID}")
        key = cursor.fetchall()
        key = key[0][0]

        encrypted_password = encrypt_function(password, key)

        cursor.execute(f"INSERT INTO passwords VALUES (?, ?, ?)", (SESSION_ID, title, encrypted_password))
        conn.commit()

        login_main_page(get_data_list())


def delete_function(tree_view):

    proceed = messagebox.askyesno("Deletar?", "Tem certeza que deseja deletar essa senha?")

    if proceed == 1:
        selected = tree_view.selection()

        if not selected:
            messagebox.showwarning("Nada selecionado",
                                   "Escolha a senha que você quer e depois clique em 'DELETAR SENHA'!")
        else:
            values = tree_view.item(selected, "values")
            tree_view.delete(selected)
            cursor.execute("DELETE FROM passwords WHERE user_id = ? AND title = (?)", (SESSION_ID, values[0]))
            conn.commit()

    return


def change_password(username):
    if name_exists(username, "users") == 0:
        messagebox.showerror("Inválido", "por favor, insira um usuário válido para prosseguir")
        return

    global USERNAME
    global EMAIL
    global CODE
    global SESSION_ID

    SESSION_ID = list(cursor.execute("SELECT id FROM users WHERE username = ?", (username,)))
    if not SESSION_ID:
        messagebox.showerror("Sem usuário", "digite um usuário para prosseguir")
        return

    SESSION_ID = SESSION_ID[0][0]
    USERNAME = username
    EMAIL = decrypt_function(SESSION_ID, "email", "users")

    CODE = str(random.randint(100000, 999999))
    title = "Código de verificação"
    message = f"""<p>Olá, <b>{username}</b>,
                                          este email tem o código de segurança para validar seu email.</p>
                                          <p>código de validação: <b>{CODE}</b></p>
                                          <p>Se você não sabe do que isso se trata, por favor ignore esta mensagem.</p>
                                          <p>Obrigado por sua atenção, de Programa Automático Python!</p>"""

    read = messagebox.showinfo("Enviando código", f"Um código de verificação está sendo enviado para {EMAIL}!"
                                                  " Isso pode demorar um pouco, por favor clique em 'ok'"
                                                  " e aguarde!")

    if read:
        if send_email(EMAIL, title, message):
            received = messagebox.askyesno("código enviado", "Por favor, confirme se recebeu o código")

            if received == 0:
                messagebox.showerror("ERRO", "Ops! Algo deu errado, por favor tente novamente")
            else:
                type_code_window("change_password")

    return


def change_execute():
    load_page()

    password_label = Label(page_label, text="Digite a nova senha")
    confirm_label = Label(page_label, text="Confirme a nova senha")

    get_new_password = Entry(page_label, show="*")
    get_confirmation = Entry(page_label, show="*")

    change_button = Button(page_label, text="MUDAR SENHA",
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
    <p>Olá {USERNAME}, segue a lista com as senhas salvas em sua base de dados:</p>
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

    title = "senhas HIDDEN"

    messagebox.showinfo("Enviando", f"Os dados estão sendo preparados para serem enviados para {EMAIL}."
                                    " O processo pode demorar um pouco, por favor aguarde!")
    send_email(EMAIL, title, message)

    return


# Functions that are used through the entire program


def type_code_window(operation_type):
    top = Toplevel()
    top.title("HIDDEN")
    top.iconbitmap("padlock.ico")
    top.resizable(width=False, height=False)

    prompt = Label(top, text="Digite o código", justify=LEFT)
    get_email_code = Entry(top)
    send_button = Button(top, text="ENVIAR",
                         command=lambda: [check_code(get_email_code.get(), operation_type), top.destroy()])

    prompt.grid(row=0, column=0)
    get_email_code.grid(row=1, column=0, padx=5)
    send_button.grid(row=1, column=1, pady=10)


def check_code(typed_code, operation_type):
    if not typed_code.isdigit():
        messagebox.showwarning("Apenas números", "O código deve conter apenas números")
    elif int(typed_code) != int(CODE):
        messagebox.showerror("Código errado", "O código digitado está incorreto")
    elif operation_type == "register":
        proceed = messagebox.askyesno("Registrar?", "Sua conta está prestes a ser registrada."
                                                    " Por favor, confirme se quer prosseguir")

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
        messagebox.showerror("Erro outlook", "Não foi possível enviar o email. Por favor, verifique se você possui"
                                             "o Microsoft Office Outlook instalado no seu aparelho")
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
        messagebox.showinfo("Usuário", "Você precisa fornecer um usuário para poder ter acesso às ferramentas do"
                                       " aplicativo")
    elif info_case == "password":
        messagebox.showinfo("Senha?", "Por questão de segurança, é necessário que você possua uma senha para que"
                                      " apenas você tenha acesso ao conteúdo salvo")
    elif info_case == "confirmation":
        messagebox.showinfo("confirmação?", "Para garantir que não houve erro de digitação na entrada anterior,"
                                            " é preciso que você digite novamente a senha que você deseja")
    elif info_case == "email":
        messagebox.showinfo("email?", "Caso você queira mudar sua senha, será necessário um email para contato para"
                                      " possibilitar a transação.")


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