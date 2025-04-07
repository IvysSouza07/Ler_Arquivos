#Nome: Ivys Emanoel Cordeiro de Souza Lima
#Matricula: 2024012961
#Importar Bibliotecas: pip install PyPDF2 | pip install python-docx

#Importando Bibliotecas.
import PyPDF2
import docx
import statistics
import os

#Definindo o caminho ate o arquivo.
diretorio = "documentos"
nome_arquivo = "dados"
extensoes = [".pdf",".docx"]

#Função para encontrar o arquivo .pdf ou .docx e verificar sua existencia.
def encontrar_arquivos():
    arquivos_encotrados = [] #Metodo para saber se existe mais de um arquivo com o nome dados.
    for extensao in extensoes:
        caminho_arquivo = os.path.join(diretorio, nome_arquivo + extensao)
        if os.path.exists(caminho_arquivo): 
            arquivos_encotrados.append(caminho_arquivo)
    return arquivos_encotrados

#Função para ler o arquivo e retornar uma lista de números(float) extraídos de cada linha, caso seja .docx.
def ler_docx(caminho_arq):
    valores = []
    documento = docx.Document(caminho_arq) 
    for paragrafo in documento.paragraphs:
        try:
            valor = float(paragrafo.text.strip()) 
            valores.append(valor)
        except ValueError:
            continue #Ignora linhas que não contem numeros.
    return valores

#Função para ler o arquivo e retornar uma lista de números(float) extraídos de cada linha, caso seja .pdf.
def ler_pdf(caminho_arq):
    valores = []
    with open(caminho_arq, 'rb') as f:
        leitor = PyPDF2.PdfReader(f)
        for pagina in leitor.pages:
            texto = pagina.extract_text()
            if texto:
                for linha in texto.splitlines():
                    try:
                        valor = float(linha.strip())
                        valores.append(valor)
                    except ValueError:
                        continue
    return valores

#Função para calcular as estatisticas conforme foi pedido na atividade.
def calcular_estatisticas(valores):
    print(f"Média: {statistics.mean(valores)}")
    print(f"Mediana: {statistics.median(valores)}")
    print(f"Somatório: {sum(valores)}")
    print(f"Maior Valor: {max(valores)}")
    print(f"Menor Valor: {min(valores)}")
    
#Função Principal para rodar o codigo.
def main():
    try:
        #Tratamento de exceção caso exista o arquivo ou se tem mais de um com o mesmo nome.
        arquivos = encontrar_arquivos()
        if not arquivos: 
            print(f"Arquivo '{nome_arquivo}.pdf' ou '{nome_arquivo}.docx' não encontrado no diretorio '{diretorio}'.")
            return
        if len(arquivos) > 1:
            print(f"Conflito: Existem dois arquivos com o nome {nome_arquivo}. Remova um deles para continuar!")
            return
        
        #Acessando o arquivo.
        caminho_arquivo = arquivos[0]
        
        print(f"Arquivo Encontrado: {caminho_arquivo}")
        
        #Aplicando a função adequada para o tipo de arquivo.
        if caminho_arquivo.endswith(".docx"):
            valores = ler_docx(caminho_arquivo)
        elif caminho_arquivo.endswith(".pdf"):
            valores = ler_pdf(caminho_arquivo)
        #Tratamento para caso o arquivo seja em outro formato.
        else: 
            print("Formato de arquivo não suportado.")
            return
        
        #Tratamento para caso o arquivo esteja vazio.
        if not valores:
            print("O arquivo está vazio ou não contém números válidos.")
            return
        
        #Chamada da função que calcula os valores dentro do arquivo.
        calcular_estatisticas(valores)
    #Tratamento de exceções.
    except Exception as ex:
        print(f"Ocorreu um erro inesperado: {ex}")

#Iniciando a função main() ou principal.        
if __name__ == "__main__":
    main()