import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import threading
import tkinter as tk
from tkinter import filedialog, simpledialog

def enviar_mensagem_whatsapp(nome, telefone, mensagem, imagem):
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(60)
        seta = pyautogui.locateCenterOnScreen(imagem)
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
        return True
    except pyautogui.ImageNotFoundException:
        print(f'Não foi possível encontrar a imagem para clicar em {nome}')
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
        return False

def main():
    workbook = None
    pagina_clientes = None
    mensagem_padrao = None
    imagem = None
    continuar_execucao = True
    paused = False

    def worker():
        nonlocal paused
        nonlocal workbook
        nonlocal pagina_clientes
        nonlocal mensagem_padrao
        nonlocal imagem
        if workbook is None or pagina_clientes is None or mensagem_padrao is None or imagem is None:
            messagebox.showerror("Erro", "Por favor, selecione o arquivo, insira a mensagem e selecione a imagem antes de iniciar.")
            return
        for linha in pagina_clientes.iter_rows(min_row=2):
            if not continuar_execucao:
                break
            if paused:
                print("Programa pausado. Pressione 'c' para continuar ou 'q' para sair.")
                while paused:
                    sleep(1)
                    if not continuar_execucao:
                        break
                print("Continuando a execução...")
            nome = linha[0].value
            telefone = linha[1].value

            if nome is None or telefone is None:
                print("Encontrado valor None. Encerrando o programa.")
                break

            sucesso = enviar_mensagem_whatsapp(nome, telefone, mensagem_padrao, imagem)
            if not sucesso:
                print("Programa pausado. Digite 'c' para continuar ou 'q' para sair.")
                while paused:
                    sleep(1)
                    if not continuar_execucao:
                        break
                print("Continuando a execução...")

    thread = threading.Thread(target=worker)


    thread = threading.Thread(target=worker)

    def pause_program():
        nonlocal paused
        paused = True

    def continue_program():
        nonlocal paused
        paused = False

    def quit_program():
        nonlocal continuar_execucao
        continuar_execucao = False
        control_window.quit()

    def select_file():
        nonlocal workbook
        nonlocal pagina_clientes
        filename = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo",
                                              filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
        if filename:
            workbook = openpyxl.load_workbook(filename)
            pagina_clientes = workbook['Sheet1']

    def input_message():
        nonlocal mensagem_padrao
        dialog = tk.Toplevel(root)
        dialog.title("Digite a mensagem")
        text = tk.Text(dialog, width=50, height=20)  # Ajuste a largura e a altura conforme necessário
        text.pack()
        ok_button = tk.Button(dialog, text="OK", command=lambda: set_message(text.get("1.0", "end-1c"), dialog))
        ok_button.pack()

    def set_message(message, dialog):
        nonlocal mensagem_padrao
        mensagem_padrao = message
        dialog.destroy()

    def select_image():
        nonlocal imagem
        imagem = filedialog.askopenfilename(initialdir="/", title="Selecione a imagem",
                                            filetypes=(("png files", "*.png"), ("jpeg files", "*.jpg"), ("all files", "*.*")))

    def start_program():
        nonlocal thread
        if workbook is None or pagina_clientes is None or mensagem_padrao is None or imagem is None:
            messagebox.showerror("Erro", "Por favor, selecione o arquivo, insira a mensagem e selecione a imagem antes de iniciar.")
            return
        root.destroy()
        thread.start()
        control_window.mainloop()

    root = tk.Tk()
    root.title("Selecionar Arquivo, Mensagem e Imagem")

    select_file_button = tk.Button(root, text="Selecionar Arquivo", command=select_file)
    select_file_button.pack()

    input_message_button = tk.Button(root, text="Inserir Mensagem", command=input_message)
    input_message_button.pack()

    select_image_button = tk.Button(root, text="Selecionar Imagem", command=select_image)
    select_image_button.pack()

    start_button = tk.Button(root, text="Iniciar", command=start_program)
    start_button.pack()

    control_window = tk.Tk()
    control_window.title("Controle do Programa")

    pause_button = tk.Button(control_window, text="Pausar", command=pause_program)
    pause_button.pack()

    continue_button = tk.Button(control_window, text="Continuar", command=continue_program)
    continue_button.pack()

    quit_button = tk.Button(control_window, text="Sair", command=quit_program)
    quit_button.pack()

    root.mainloop()

    while continuar_execucao:
        if not paused:
            opcao = input("Digite 'p' para pausar: ").lower()
            if opcao == 'p':
                paused = True
        else:
            opcao = input("Programa pausado. Digite 'c' para continuar ou 'q' para sair: ").lower()
            if opcao == 'c':
                paused = False
            elif opcao == 'q':
                continuar_execucao = False
                paused = False
            else:
                print("Opção inválida. Por favor, digite 'c' ou 'q'.")

    thread.join()
    print("Programa encerrado.")

if __name__ == "__main__":
    main()
