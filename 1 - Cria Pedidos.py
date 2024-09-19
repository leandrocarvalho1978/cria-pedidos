import os
import shutil
from tkinter import Tk, Label, Entry, Button, Text, END, messagebox
from win32com.client import Dispatch

# Função para criar um atalho
def criar_atalho(nome, destino):
    shell = Dispatch('WScript.Shell')
    atalho = shell.CreateShortcut(nome)
    atalho.TargetPath = destino
    atalho.Save()

# Função para resolver o caminho de um atalho
def resolver_atalho(caminho_atalho):
    shell = Dispatch('WScript.Shell')
    atalho = shell.CreateShortcut(caminho_atalho)
    return atalho.TargetPath

# Função para processar os atalhos dentro da pasta de análise
# Não remove atalhos antigos, apenas adiciona novos, se necessário
def processar_atalhos(pasta_analise):
    atalhos = [os.path.join(pasta_analise, item) for item in os.listdir(pasta_analise) if item.endswith('.lnk')]
    
    for caminho_atalho in atalhos:
        caminho_destino = resolver_atalho(caminho_atalho)

        if os.path.isdir(caminho_destino):
            subpastas = os.listdir(caminho_destino)
            
            for subpasta in subpastas:
                subpasta_caminho = os.path.join(caminho_destino, subpasta)
                if os.path.isdir(subpasta_caminho):
                    nome_novo_atalho = f"{os.path.splitext(os.path.basename(caminho_atalho))[0]} {subpasta}.lnk"
                    caminho_novo_atalho = os.path.join(pasta_analise, nome_novo_atalho)

                    # Verificar se o atalho já existe antes de criar
                    if not os.path.exists(caminho_novo_atalho):
                        criar_atalho(caminho_novo_atalho, subpasta_caminho)
                        print(f"Criado atalho: {nome_novo_atalho}")
                    else:
                        print(f"Atalho {nome_novo_atalho} já existe e não foi recriado.")

# Função principal para processar as pastas e criar os arquivos
def processar():
    nome_nova_pasta = entrada_nome.get().strip()
    if not nome_nova_pasta:
        messagebox.showerror("Erro", "Por favor, digite o nome da nova pasta do cliente.")
        return

    caminho_nova_pasta = os.path.join("C:\\Users\\leand\\3D Objects", nome_nova_pasta)
    os.makedirs(caminho_nova_pasta, exist_ok=True)  # Cria a pasta, se não existir

    # Captura as palavras-chave inseridas no Text, divididas por linha
    palavras_chave_produtos = entrada_palavras.get("1.0", END).strip()
    if not palavras_chave_produtos:
        messagebox.showerror("Erro", "Por favor, insira as palavras-chave dos produtos.")
        return

    # Cada linha será considerada uma palavra-chave
    lista_palavras_chave_produtos = [palavra.strip() for palavra in palavras_chave_produtos.splitlines() if palavra.strip()]

    caminho_principal = "D:\\1 - Kellen Cortadores\\2 - Catálogo\\1 - Biscoitos"
    caminho_svg_modelo = "C:\\Users\\leand\\AppData\\Roaming\\inkscape\\templates\\default.svg"

    for palavra_chave in lista_palavras_chave_produtos:
        print(f"Buscando pastas com a palavra-chave '{palavra_chave}'...")

        pastas_encontradas = [os.path.join(root, dir_name)
                              for root, dirs, _ in os.walk(caminho_principal)
                              for dir_name in dirs
                              if palavra_chave.lower() in dir_name.lower()]

        if pastas_encontradas:
            for pasta_produto in pastas_encontradas:
                nome_atalho = os.path.join(caminho_nova_pasta, os.path.basename(pasta_produto) + '.lnk')
                
                # Verificar se o atalho já existe
                if not os.path.exists(nome_atalho):
                    criar_atalho(nome_atalho, pasta_produto)
                    print(f"Criado atalho para pasta: {os.path.basename(pasta_produto)}")
                else:
                    print(f"Atalho para pasta '{os.path.basename(pasta_produto)}' já existe. Não foi recriado.")
        else:
            nova_pasta = os.path.join(caminho_nova_pasta, palavra_chave)
            os.makedirs(nova_pasta, exist_ok=True)  # Cria a pasta, se não existir
            print(f"Nenhuma pasta encontrada com a palavra-chave '{palavra_chave}'. Nova pasta criada.")

            for extensao in ['.svg', '.png', '.sldprt', ' M.stl', ' C.stl']:
                nome_arquivo = os.path.join(nova_pasta, f"{palavra_chave}{extensao}".strip())
                
                # Verificar se o arquivo já existe
                if not os.path.exists(nome_arquivo):
                    if extensao == '.svg':
                        shutil.copyfile(caminho_svg_modelo, nome_arquivo)
                    else:
                        open(nome_arquivo, 'a').close()
                    print(f"Arquivo {nome_arquivo} criado.")
                else:
                    print(f"Arquivo {nome_arquivo} já existe. Não foi recriado.")

    # Agora processamos os atalhos, mas sem remover os antigos
    processar_atalhos(caminho_nova_pasta)
    os.startfile(caminho_nova_pasta)
    
    janela.destroy()

# Configuração da interface gráfica
janela = Tk()
janela.title("Gerenciador de Pastas e Arquivos")

Label(janela, text="Nome da Nova Pasta do Cliente:").grid(row=0, column=0, sticky="w")
entrada_nome = Entry(janela, width=50)
entrada_nome.grid(row=0, column=1)

# Alterando o rótulo para refletir a nova funcionalidade
Label(janela, text="Palavras-chave dos Produtos (uma por linha):").grid(row=1, column=0, sticky="w")
entrada_palavras = Text(janela, width=50, height=6)  # Aumentando um pouco a altura para facilitar a entrada
entrada_palavras.grid(row=1, column=1)

Button(janela, text="Processar", command=processar).grid(row=2, columnspan=2)

janela.mainloop()
