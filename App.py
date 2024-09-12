import customtkinter as ctk
from tkinter import filedialog
import win32com.client as win32
from openpyxl import load_workbook
from docxtpl import DocxTemplate
from docx2pdf import convert

#Definindo aparencias padrões
ctk.set_appearance_mode('System')
ctk.set_default_color_theme('blue')

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.configuracao_layout()
        self.todo_sistema()

    def configuracao_layout(self):
        self.title('Gerador de certificado')
        self.geometry('500x250')

    def gerar(self):
        global send_email, entry_assunto, entry_corpo
        file = filedialog.askdirectory()
        file = file.replace('/', '\\')

        planilha = load_workbook('dados.xlsx')
        aba_ativa = planilha.active

        # Armazena a data, nome do palestrante, carga horaria e titulo da palestra nas variaveis
        data = aba_ativa['E2'].value
        palestra = aba_ativa['C2'].value
        carga = aba_ativa['D2'].value

        for celula in aba_ativa['A']:
            if celula.value != "nome":
                linha = celula.row
                nome = aba_ativa[f'A{linha}'].value
                email_pessoa = aba_ativa[f'B{linha}'].value

                # Código que troca o que está escrito no documento base para o que está na planilha utilizando {}
                doc = DocxTemplate("certificado.docx")
                context = {'nome': nome, 'palestra': palestra, 'data': data, 'carga': carga}
                doc.render(context)
                doc.save("replaced1.docx")

                # Converte o arquivo Reescrito em PDF e salva o nome do arquivo como: Certificado-NomeDoColaborador
                convert(r"replaced1.docx", fr"{file}\Certificado-{nome}.pdf")

                if send_email.get() == 1:
                    outlook = win32.Dispatch('outlook.application')
                    assunto = entry_assunto.get()
                    email_corpo = entry_corpo.get(0.0, ctk.END)
                    # criar um email
                    email = outlook.CreateItem(0)

                    # configurar as informações do seu e-mail
                    email.To = f'{email_pessoa}'
                    email.Subject = f"{assunto}"
                    email.HTMLBody = f"""
                            <p>{email_corpo}</p>

                            <p>Att,</p>
                            <p>Paulo</p>
                            """

                    anexo = rf"{file}\Certificado-{nome}.pdf"
                    email.Attachments.Add(anexo)

                    email.Send()


    def todo_sistema(self):
        global send_email, entry_assunto, entry_corpo
        def hide_element():
            if send_email.get() == 1:
                self.geometry('500x400')
                frame1.place(x=170, y=200)
                label_assunto.pack()
                entry_assunto.pack()
                label_corpo.pack()
                entry_corpo.pack()
            else:
                self.geometry('500x250')
                label_assunto.pack_forget()
                label_corpo.pack_forget()
                entry_corpo.pack_forget()
                entry_assunto.pack_forget()

        send_email = ctk.IntVar()

        #Checkbox e-mail
        checkbox_email = ctk.CTkCheckBox(self, text='Enviar certificados para email', border_width=1, hover_color='#555', variable=send_email, onvalue=1, offvalue=0, command=hide_element)
        checkbox_email.place(x=150, y=150)

        #Frames
        frame1 = ctk.CTkFrame(self, width=500, border_width=0, bg_color='transparent', fg_color='transparent')
        frame_top = ctk.CTkFrame(self, width=500, height=50, border_width=0, fg_color='#70DAE5', corner_radius=0)
        frame_top.place(x=0, y=0)

        #Labels
        label_assunto = ctk.CTkLabel(frame1, text='Assunto do email')
        label_corpo = ctk.CTkLabel(frame1, text='Corpo do email')
        label_titulo = ctk.CTkLabel(frame_top, text='Gerador de certificado', bg_color='transparent', fg_color='transparent', font=('Arial', 24))
        label_titulo.place(x=140, y=10)

        #Entrys
        entry_assunto = ctk.CTkEntry(frame1)
        entry_corpo = ctk.CTkTextbox(frame1, width=200, height=100)

        #Botao
        botao1 = ctk.CTkButton(self, text='Gerar certificados', command=self.gerar, height=40, font=('Arial', 18))
        botao1.place(x=180, y=80)

if __name__ == '__main__':
    app = App()
    app.mainloop()
