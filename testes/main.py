import tkinter as tk
from tkinter import ttk, messagebox
from ttkthemes import ThemedTk
from PIL import Image, ImageTk
from threading import Thread

# Funções importadas dos módulos de lógica
from aba1 import (
    desativar_selecao, criar_arquivo, cadastrar_renovar_chave,
    on_checkbox_clicked, update_project_dropdown
)
from base import on_browse, on_select_image_dir, limpar_imagens

# --- Funções da Interface ---

def handle_checkbox_click():
    """Lida com o clique na checkbox em uma thread separada para não travar a UI."""
    def task():
        if on_checkbox_clicked(checkbox_var):
            # A função on_checkbox_clicked já mostra erros,
            # então apenas atualizamos o dropdown em caso de sucesso.
            update_project_dropdown(checkbox_var, project_dropdown)
        else:
            # Garante que o dropdown seja limpo se a validação falhar
            project_dropdown.set('')
            project_dropdown.configure(values=[])
            
    Thread(target=task, daemon=True).start()

def iniciar_geracao():
    """Inicia a geração do documento em uma thread para não congelar a janela."""
    status_label.config(text="Gerando documento...")
    btn_ativar.config(state='disabled')
    
    # Coleta os valores dos campos da UI
    args_para_criar_arquivo = (
        entries["Arquivo padrão"].get(),
        entries["Diretório de Imagens"].get(),
        entries["ID/Nome do cenário de teste"].get(),
        entries["Nome do Tester"].get(),
        ambiente_dropdown.get(),
        entries["Perfil"].get(),
        entries["Bugs"].get(),
        project_dropdown.get(),
        checkbox_var
    )

    gerar_thread = Thread(target=criar_arquivo, args=args_para_criar_arquivo, daemon=True)
    gerar_thread.start()
    janela.after(100, lambda: verificar_thread(gerar_thread))

def verificar_thread(thread):
    """Verifica periodicamente se a thread de geração terminou."""
    if thread.is_alive():
        janela.after(100, lambda: verificar_thread(thread))
    else:
        btn_ativar.config(state='normal')
        status_label.config(text="")

# --- Configuração da Janela Principal ---
janela = ThemedTk(theme="arc")
janela.geometry("650x580")
janela.title("Assistente de Testes")
try:
    janela.iconbitmap("ey.ico")
except tk.TclError:
    print("Ícone 'ey.ico' não encontrado. O programa continuará sem ícone.")

# --- Estilos ---
style = ttk.Style()
style.configure("TEntry", padding=(5, 4), relief="flat")
style.configure("TCombobox", padding=(5, 4), relief="flat")
style.configure("TButton", padding=(10, 5), relief="flat")

# --- Abas (Notebook) ---
notebook = ttk.Notebook(janela)
aba1 = ttk.Frame(notebook)
notebook.add(aba1, text="Documentos")
notebook.pack(expand=True, fill='both', padx=10, pady=10)

# --- Widgets da Aba 1 ---

# Dicionário para armazenar os campos de entrada para fácil acesso
entries = {}
# O campo 'Produto' foi removido pois não era utilizado na interface
labels_text = ["Arquivo padrão", "Diretório de Imagens", "Nome do Tester", "ID/Nome do cenário de teste", "Perfil", "Bugs"]

# Criação de Labels e Campos de Entrada
for i, text in enumerate(labels_text):
    ttk.Label(aba1, text=text).grid(row=i, column=0, padx=20, pady=10, sticky='w')
    entry = ttk.Entry(aba1, width=35)
    entry.grid(row=i, column=1, padx=10, pady=10, sticky='w')
    entries[text] = entry

# Botões de "Procurar"
btn_browse_arquivo = ttk.Button(aba1, text="Procurar", command=lambda: on_browse(entries["Arquivo padrão"]))
btn_browse_arquivo.grid(row=0, column=2, padx=5, pady=10, sticky='w')

btn_browse_imagens = ttk.Button(aba1, text="Procurar", command=lambda: on_select_image_dir(entries["Diretório de Imagens"]))
btn_browse_imagens.grid(row=1, column=2, padx=5, pady=10, sticky='w')

btn_clear_images = ttk.Button(aba1, text="Limpar imagens", command=lambda: limpar_imagens(entries["Diretório de Imagens"]))
btn_clear_images.grid(row=1, column=3, padx=5, pady=10, sticky='w')

# Dropdown de Ambiente
ttk.Label(aba1, text="Ambiente").grid(row=len(labels_text), column=0, padx=20, pady=10, sticky='w')
amb_var = tk.StringVar(value='HML')
ambiente_dropdown = ttk.Combobox(aba1, textvariable=amb_var, values=["DEV", "HML", "PRD"], state="readonly", width=33)
ambiente_dropdown.grid(row=len(labels_text), column=1, padx=10, pady=10, sticky='w')

# Dropdown de Projetos DevOps
ttk.Label(aba1, text="Produto").grid(row=len(labels_text) + 1, column=0, padx=20, pady=10, sticky='w')
project_dropdown = ttk.Combobox(aba1, state="readonly", width=33)
project_dropdown.grid(row=len(labels_text) + 1, column=1, padx=10, pady=10, sticky='w')

# Checkbox e Chave DevOps
checkbox_var = tk.BooleanVar(value=False)
checkbox = ttk.Checkbutton(aba1, text="Integrar com DevOps", variable=checkbox_var, command=handle_checkbox_click)
checkbox.grid(row=3, column=2, columnspan=2, padx=(20, 0), pady=10, sticky='w')

try:
    image = Image.open("chave.png").resize((16, 16))
    photo = ImageTk.PhotoImage(image)
    btn_chave_devops = ttk.Button(aba1, image=photo, command=lambda: cadastrar_renovar_chave(janela))
    btn_chave_devops.grid(row=3, column=3, padx=(20, 0), pady=10, sticky='e')
except Exception as e:
    messagebox.showerror("Erro de Ícone", f"Não foi possível carregar o ícone 'chave.png': {e}")

# Bindings para limpar seleção (evita seleções conflitantes)
project_dropdown.bind("<<ComboboxSelected>>", lambda event: desativar_selecao(ambiente_dropdown, project_dropdown))
ambiente_dropdown.bind("<<ComboboxSelected>>", lambda event: desativar_selecao(ambiente_dropdown, project_dropdown))

# Botão de Gerar e Status
status_label = ttk.Label(aba1, text="")
status_label.grid(row=len(labels_text) + 3, column=0, columnspan=4, pady=10)

btn_ativar = ttk.Button(aba1, text="Gerar Documento", command=iniciar_geracao, style="Accent.TButton")
btn_ativar.grid(row=len(labels_text) + 2, column=1, columnspan=2, padx=20, pady=20, sticky='ew')

janela.mainloop()