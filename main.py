from customtkinter import CTk
from dashboard import DashboardEmpresarial


if __name__ == "__main__":
    root = CTk()
    # Define o ícone para a janela principal
    root.iconbitmap("icone.ico")  # Substitua "icone.ico" pelo caminho do arquivo do ícone
    app = DashboardEmpresarial(root)
    root.mainloop()