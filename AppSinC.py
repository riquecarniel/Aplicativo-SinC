import os
import math
import webbrowser
import pandas as pd
import customtkinter as ctk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox, Label

# Define o tema para a janela criada
ctk.set_appearance_mode('Dark')
ctk.set_default_color_theme('blue')

# Cria a interface gráfica
janela = ctk.CTk()
janela.title('SinC 1.2')
janela.geometry('325x280')
janela.rowconfigure(0, weight=1)
janela.columnconfigure([0, 1], weight=1)
janela.resizable(width=False, height=False)

# Função para criar os imputs
def criar_entry(janela, placeholder, x, y):
    campo = ctk.CTkEntry(
        janela,
        placeholder_text=placeholder,
        placeholder_text_color='white',
        text_color='white',
        border_width=3,
        corner_radius=10,
        width=305)
    campo.place(x=x, y=y)
    return campo

# Texto da tela
texto_label = Label(janela, text='Aplicativo SinC', font=('Helvetica', 25, 'bold'), background=janela.cget('bg'), fg='white')
texto_label.place(x=40, y=2)

# Função para abrir o drive(Google Drive)
def abrir_drive():
    url_do_site = 'https://drive.google.com/drive/folders/1ExaULcWX7ytNRkJ0UbqExkQw_3g60frI?usp=drive_link'
    webbrowser.open(url_do_site)

# Função para calcular a apótema do octógono
def calcular_apotema(lado_octogono):
    n_lados = 8
    apotema = lado_octogono / (2 * math.tan(math.pi / n_lados))
    resultado_apotema = apotema * 2
    return resultado_apotema
# Função para calcular a altura do triângulo equilátero
def calcular_triangulo(lado_triangulo):
    altura = (math.sqrt(3) / 2) * lado_triangulo
    return altura
# Função para calcular a diagonal do quadrado
def calcular_quadrado(lado_quadrado):
    resultado_quadrado = lado_quadrado * math.sqrt(2)
    return resultado_quadrado

# Campos de entrada
nome_excel = criar_entry(janela, 'Informe o nome do Excel a ser gerado', 10, 130)
buraco_excel = criar_entry(janela, 'Informe a profundidade da Fundação', 10, 165)
limite_excel = criar_entry(janela, 'Informe o limite mínimo de Suporte', 10, 200)

# Variável para armazenar o caminho do arquivo Excel selecionado
caminho_excel = None

# Função para selecionar o arquivo Excel
def importar_excel():
    global caminho_excel
    caminho_excel = askopenfilename(title='Selecione o Excel!')

# Função principal para verificar os dados
def verificar():
    if not caminho_excel:
        messagebox.showwarning('Aviso', 'Por favor, selecione um arquivo Excel.')
        return

    if not nome_excel.get():
        messagebox.showwarning('Aviso', 'Por favor, informe um nome para o arquivo Excel.')
        return
    
    if not buraco_excel.get():
        messagebox.showwarning('Aviso', 'Por favor, informe a profundidade da Fundação.')
        return
    else:
        try:
            buraco = float(buraco_excel.get().replace(',', '.'))
        except ValueError:
            messagebox.showwarning('Aviso', 'Por favor, informe um valor numérico para a profundidade da fundação.')
            return

    if not limite_excel.get():
        messagebox.showwarning('Aviso', 'Por favor, informe o limite mínimo de suporte.')
        return
    else:
        try:
            limite_suporte = float(limite_excel.get().replace(',', '.'))
        except ValueError:
            messagebox.showwarning('Aviso', 'Por favor, informe um valor numérico para o limite mínimo de suporte.')
            return

    nome_arquivo = nome_excel.get() + '.xlsx'
    buraco = float(buraco_excel.get().replace(',', '.'))
    limite_suporte = float(limite_excel.get().replace(',', '.'))

    df_placas_origem = pd.read_excel(caminho_excel, sheet_name='Placas')
    df_suportes_origem = pd.read_excel(caminho_excel, sheet_name='Suportes')

    df_placas_selecionadas = df_placas_origem.iloc[3:, [3, 5, 6, 7, 14, 15, 13, 11, 12]]
    df_suportes_selecionados = df_suportes_origem.iloc[3:, [3, 5, 6, 7, 10, 11, 12]]

    with pd.ExcelWriter(nome_arquivo) as writer:
        df_placas_selecionadas.to_excel(writer, sheet_name='Placas', index=False, header=False)
        df_suportes_selecionados.to_excel(writer, sheet_name='Suportes', index=False, header=False)

    df_final_placas = pd.read_excel(nome_arquivo, sheet_name='Placas')
    df_final_suportes = pd.read_excel(nome_arquivo, sheet_name='Suportes')

    df = df_final_placas.merge(df_final_suportes, how='left')

    df_ordenada = df.sort_values('Código', ascending=True)

    df_ordenada = df_ordenada.drop_duplicates()

    df_ordenada.to_excel(nome_arquivo, sheet_name='Dados', index=False)

    # Carregar a planilha para fazer o calculo da fundação
    planilha = pd.read_excel(nome_arquivo, sheet_name='Dados')

    planilha['Fundação'] = buraco
    planilha['Altura Placa'] = planilha['Dimensão']
    planilha['Altura Suporte'] = planilha['Altura']

    remover = ['L', '=', 'm', 'Ø']
    for char in remover:
        planilha['Altura Placa'] = planilha['Altura Placa'].str.replace(char, '')

    remover = ['m']
    for char in remover:
        planilha['Altura Suporte'] = planilha['Altura Suporte'].str.replace(char, '')

    planilha.loc[planilha['Altura Placa'].str.contains('x', na=False), 'Altura Placa'] = \
    planilha.loc[planilha['Altura Placa'].str.contains('x', na=False), 'Altura Placa'].str.split('x').str[1]

    planilha['Altura Placa'] = planilha['Altura Placa'].str.replace(',', '.').str.strip()
    planilha['Altura Suporte'] = planilha['Altura Suporte'].str.replace(',', '.').str.strip()

    for indice, linha in planilha.iterrows():
        if linha['Código'] == 'R-1':
            altura_placa = float(linha['Altura Placa'])
            lado_octogono = altura_placa
            resultado_apotema = calcular_apotema(lado_octogono)
            planilha.at[indice, 'Altura Placa'] = resultado_apotema
        elif linha['Código'] == 'R-2':
            altura_placa = float(linha['Altura Placa'])
            lado_triangulo = altura_placa
            resultado_triangulo = calcular_triangulo(lado_triangulo)
            planilha.at[indice, 'Altura Placa'] = resultado_triangulo
        elif linha['Código'].startswith('A-'):
            altura_placa = float(linha['Altura Placa'])
            lado_quadrado = altura_placa
            resultado_quadrado = calcular_quadrado(lado_quadrado)
            planilha.at[indice, 'Altura Placa'] = resultado_quadrado
     
    planilha['Resultado'] = pd.to_numeric(planilha['Altura Placa']) + pd.to_numeric(planilha['Altura Suporte']) + buraco

    planilha['Limite Suporte'] = limite_suporte

    planilha['Altura Total'] = planilha['Resultado']

    # Substituir valores NaN por 0 na coluna 'Resultado'
    planilha['Altura Total'].fillna(0, inplace=True)

    # Arredonda os valores na coluna 'Altura Total' para o próximo múltiplo de 0.5
    planilha['Altura Total'] = planilha['Altura Total'].apply(lambda x: math.ceil(x * 2) / 2)

    # Verifica se o valor é menor que o limite de suporte(limite_suporte) e ajusta se necessário
    planilha.loc[planilha['Altura Total'] < limite_suporte, 'Altura Total'] = limite_suporte

    # Formata os valores na coluna 'Altura Total'
    planilha['Altura Total'] = planilha['Altura Total'].apply(lambda x: '{:.2f}'.format(x))
    planilha['Altura Total'] = planilha['Altura Total'].str.replace('nan', '')
    planilha['Altura Total'] = planilha['Altura Total'].str.replace('.', ',').str.strip()
    planilha['Altura Total'] = planilha['Altura Total'].str.replace('0,00m', '')

    # Verifica e remover dados na coluna 'Altura Total' se a coluna 'Altura Suporte' estiver vazia
    planilha.loc[planilha['Altura Suporte'].isnull(), 'Altura Total'] = ''

    # Formata os valores na coluna 'Altura Total' para terem duas casas decimais
    planilha['Resultado'] = planilha['Resultado'].apply(lambda x: '{:.2f}'.format(float(x)))

    # Substituir vírgulas por pontos e remover espaços em branco extras
    planilha['Resultado'] = planilha['Resultado'].str.replace(',', '.').str.strip()

    # Verificar se algum valor na coluna 'Resultado' é 3.5, 4.5, 5.5, ..., e substituir na coluna 'Altura Total'
    valores_substituir = ['3.50', '4.50', '5.50', '6.50', '7.50', '8.50', '9.50', '10.50', '11.50', '12.50', 13.5, 14.5, 15.5, 16.5, 17.5, 18.5, 19.5, 20.5]

    # Substituir os valores
    for valor_substituir in valores_substituir:
        planilha.loc[planilha['Resultado'] == str(valor_substituir), 'Altura Total'] = str(valor_substituir)

    planilha['Resultado'] = planilha['Resultado'].str.replace('nan', '')
    planilha['Altura Total'] = planilha['Altura Total'].str.replace('.', ',').str.strip()
    planilha['Altura Total'] = planilha['Altura Total'].astype(str) + 'm'
    planilha.loc[planilha['Altura Suporte'].isnull(), 'Altura Total'] = ''

    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        planilha.to_excel(writer, sheet_name='Dados', index=False)

    planilha = pd.read_excel(nome_arquivo, sheet_name='Dados')
    dados_iguais = pd.DataFrame(columns=planilha.columns)

    # Verifica os dados iguais
    for indice, linha in planilha.iterrows():
        dados_comparar = planilha[(planilha['Eixo'] == linha['Eixo']) & (planilha['Estaca/KM'] == linha['Estaca/KM']) & (planilha['Lado'] == linha['Lado'])]
        if len(dados_comparar) > 1:
            dados_iguais = pd.concat([dados_iguais, dados_comparar], ignore_index=True)

    # Remove os dados duplicados
    dados_iguais = dados_iguais.drop_duplicates()

    nome_arquivo_salvar = os.path.join(os.getcwd(), nome_arquivo)

    # Salva os dados iguais em uma nova aba(Eixo-Estaca-Lado Iguais)
    with pd.ExcelWriter(nome_arquivo_salvar, mode='a', engine='openpyxl') as writer:
        dados_iguais.to_excel(writer, sheet_name='Eixo-Estaca-Lado Iguais', index=False)

    # Carrega a planilha
    resumoplaca = pd.read_excel(nome_arquivo_salvar, sheet_name='Dados')
    # Verificar e copiar valores únicos para a planilha Resumo
    df_resumo = resumoplaca.groupby(['Código', 'Dimensão', 'Área']).size().reset_index(name='Quantidade')
    # Converter a coluna 'Área' para o tipo float
    df_resumo['Área'] = df_resumo['Área'].astype(str).str.replace(',', '.').astype(float)
    # Multiplicar os valores da coluna 'Área' pela coluna 'Quantidade'
    df_resumo['Área Total'] = df_resumo['Área'] * df_resumo['Quantidade']

    df_resumo.rename(columns={'Área': 'Área Unitária'}, inplace=True)

    with pd.ExcelWriter(nome_arquivo_salvar, mode='a', engine='openpyxl') as writer:
        df_resumo.to_excel(writer, sheet_name='Resumo Placas', index=False)

    resumosuportes = pd.read_excel(nome_arquivo_salvar, sheet_name='Dados')
    # Verificar e copiar valores únicos para a planilha Resumo Suportes
    df_resumo_suportes = resumosuportes.groupby(['Tipo', 'Material', 'Altura Total']).size().reset_index(name='Quantidade')

    # Multiplicar por 2 a quantidade quando o tipo de suporte for 'Coluna Dupla'
    df_resumo_suportes.loc[df_resumo_suportes['Tipo'] == 'Coluna Dupla', 'Quantidade'] *= 2

    # Renomear a coluna 'Tipo' para 'Tipo Suporte'
    df_resumo_suportes.rename(columns={'Tipo': 'Tipo Suporte'}, inplace=True)

    # Salvar na planilha Resumo Suportes
    with pd.ExcelWriter(nome_arquivo_salvar, mode='a', engine='openpyxl') as writer:
        df_resumo_suportes.to_excel(writer, sheet_name='Resumo Suportes', index=False)

    dados_faltando = planilha[planilha.isnull().any(axis=1)]

    if not dados_faltando.empty:
        messagebox.showwarning('Aviso', 'Planilha Gerada, mas existem dados faltando.')
    else:
        messagebox.showinfo('SinC', 'Planilha Gerada.')

    janela.destroy()

# Função para criar botões
def criar_button(janela, texto, comando, x, y, fg_cor, hover_cor, height, width, corner_radius):
    botao = ctk.CTkButton(
        janela,
        text=texto,
        font=('Helvetica', 17, 'bold'),
        command=comando,
        height=height,
        width=width,
        corner_radius=corner_radius,
        fg_color=fg_cor,
        hover_color=hover_cor)
    botao.place(x=x, y=y)
    return botao

# Botões
but_arquivos = criar_button(janela, 'Planilha SinC', abrir_drive, 10, 50, None, None, height=35, width=305, corner_radius=200)
but_abrir = criar_button(janela, 'Importar Excel', importar_excel, 10, 90, None, None, height=35, width=305, corner_radius=200)
but_verificar = criar_button(janela, 'Processar', verificar, 10, 235, None, None, height=35, width=305, corner_radius=200)

janela.mainloop()