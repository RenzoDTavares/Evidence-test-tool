import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import base64

def on_browse(entry_widget):
    """Abre um diálogo para selecionar um arquivo Word (.docx) e insere o caminho no campo."""
    filepath = filedialog.askopenfilename(
        title="Selecione o arquivo padrão",
        filetypes=[("Documentos Word", "*.docx"), ("Todos os arquivos", "*.*")]
    )
    if filepath:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

def on_select_image_dir(entry_widget):
    """Abre um diálogo para selecionar um diretório de imagens e insere o caminho no campo."""
    directory = filedialog.askdirectory(title="Selecione o diretório com as imagens de evidência")
    if directory:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, directory)

def limpar_imagens(campo_diretorio_imagens):
    """Remove todos os arquivos de imagem de um diretório especificado."""
    diretorio_imagens = campo_diretorio_imagens.get()
    if not diretorio_imagens:
        messagebox.showwarning("Atenção", "Por favor, selecione um diretório de imagens primeiro.")
        return

    diretorio = Path(diretorio_imagens)
    if not diretorio.is_dir():
        messagebox.showerror("Erro", f"O diretório '{diretorio_imagens}' não existe.")
        return

    extensoes_imagem = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
    arquivos_removidos = 0
    erros = []

    for arquivo in diretorio.iterdir():
        if arquivo.is_file() and arquivo.suffix.lower() in extensoes_imagem:
            try:
                arquivo.unlink()
                arquivos_removidos += 1
            except Exception as e:
                erros.append(f"Falha ao remover {arquivo.name}: {e}")

    if erros:
        messagebox.showerror("Erro na Limpeza", "\n".join(erros))
    
    messagebox.showinfo("Limpeza Concluída", f"{arquivos_removidos} arquivo(s) de imagem foram removido(s) com sucesso.")

def decodificar_chave():
    """Lê e decodifica a chave de acesso do Azure DevOps do arquivo local."""
    try:
        with open("chave_devops.txt", "r") as file:
            chave_codificada = file.read().strip()
            return base64.b64decode(chave_codificada.encode('utf-8')).decode('utf-8')
    except FileNotFoundError:
        # Não mostra um erro, pois o arquivo pode não existir ainda.
        # A validação da chave lidará com o retorno None.
        return None
    except (base64.binascii.Error, UnicodeDecodeError):
        messagebox.showerror("Erro de Chave", "A chave armazenada está em um formato inválido. Por favor, cadastre-a novamente.")
        return None