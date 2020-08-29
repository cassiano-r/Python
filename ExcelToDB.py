#Biblioteca para usar Excel
import xlrd
import xlwt

"""
#Biblioteca para acessar DB
import pyodbc

conn = pyodbc.connect()

cursor = conn.cursor()
cursor.execute('SELECT * FROM TABLE')

for row in cursor:
    print(row)


cursor.close()
conn.close()

"""

#Biblioteca para capturar requesitos e caminhos do os
import os

#Caminho do Diretório
directory = "\\\\XXX.XXX.XXX.XXX\\\20200330\\"

#Verifica se Diretório existe, senão cria
if not os.path.exists(directory + "temp"):
    os.makedirs(directory + "temp")

#Função de manipulacao do arquivo
#---------------------------------------------------------------------------------------------------------
def manipulacao(nomeArquivo):
    
    #Cria uma planilha
    wb = xlwt.Workbook()

    #Cria Aba - Futuros
    futuros = wb.add_sheet('futuros')

    #Cria Aba - Outros Fundos
    outrosfundos = wb.add_sheet('outrosfundos')

    #Cria Aba - RendaFixa
    rendafixa = wb.add_sheet('rendafixa')

    #Cria Aba - Contas
    contas = wb.add_sheet('contas')

    #Cria Aba - Tesouraria
    tesouraria = wb.add_sheet('tesouraria')

    #Cria Aba - Patrimonio
    patrimonio = wb.add_sheet('patrimonio')

    #Cria Aba - Rentabilidade
    rentabilidade = wb.add_sheet('rentabilidade')

    #arquivo Excel
    book = xlrd.open_workbook(directory + nomeArquivo)

    #Obtem a 1ª Aba
    sh = book.sheet_by_index(0)

    #Variavel auxiliar
    aux = "null"
    count = 0

    for line in range(sh.nrows):

        #Escreve os valores nas devidas abas
        if(aux == "futuros"):      
            futuros.write(count, 0, sh.cell_value(rowx=line, colx=0))
            futuros.write(count, 1, sh.cell_value(rowx=line, colx=1))
            futuros.write(count, 2, sh.cell_value(rowx=line, colx=2))
            futuros.write(count, 3, sh.cell_value(rowx=line, colx=3))
            futuros.write(count, 4, sh.cell_value(rowx=line, colx=4))
            futuros.write(count, 5, sh.cell_value(rowx=line, colx=5))
            futuros.write(count, 6, sh.cell_value(rowx=line, colx=6))
            futuros.write(count, 7, sh.cell_value(rowx=line, colx=7))
            futuros.write(count, 8, sh.cell_value(rowx=line, colx=8))
            futuros.write(count, 9, sh.cell_value(rowx=line, colx=9))
            count +=1;
        elif(aux == "outrosfundos"):      
            outrosfundos.write(count, 0, sh.cell_value(rowx=line, colx=0))
            outrosfundos.write(count, 1, sh.cell_value(rowx=line, colx=1))
            outrosfundos.write(count, 2, sh.cell_value(rowx=line, colx=2))
            outrosfundos.write(count, 3, sh.cell_value(rowx=line, colx=3))
            outrosfundos.write(count, 4, sh.cell_value(rowx=line, colx=4))
            outrosfundos.write(count, 5, sh.cell_value(rowx=line, colx=5))
            outrosfundos.write(count, 6, sh.cell_value(rowx=line, colx=6))
            outrosfundos.write(count, 7, sh.cell_value(rowx=line, colx=7))
            outrosfundos.write(count, 8, sh.cell_value(rowx=line, colx=8))
            outrosfundos.write(count, 9, sh.cell_value(rowx=line, colx=9))
            outrosfundos.write(count, 10, sh.cell_value(rowx=line, colx=10))
            outrosfundos.write(count, 11, sh.cell_value(rowx=line, colx=11))
            count +=1;
        elif(aux == "rendafixa"):
            rendafixa.write(count, 0, sh.cell_value(rowx=line, colx=0))
            rendafixa.write(count, 1, sh.cell_value(rowx=line, colx=1))
            rendafixa.write(count, 2, sh.cell_value(rowx=line, colx=2))
            rendafixa.write(count, 3, sh.cell_value(rowx=line, colx=3))
            rendafixa.write(count, 4, sh.cell_value(rowx=line, colx=4))
            rendafixa.write(count, 5, sh.cell_value(rowx=line, colx=5))
            rendafixa.write(count, 6, sh.cell_value(rowx=line, colx=6))
            rendafixa.write(count, 7, sh.cell_value(rowx=line, colx=7))
            rendafixa.write(count, 8, sh.cell_value(rowx=line, colx=8))
            rendafixa.write(count, 9, sh.cell_value(rowx=line, colx=9))
            rendafixa.write(count, 10, sh.cell_value(rowx=line, colx=10))
            rendafixa.write(count, 11, sh.cell_value(rowx=line, colx=11))
            rendafixa.write(count, 12, sh.cell_value(rowx=line, colx=12))
            rendafixa.write(count, 13, sh.cell_value(rowx=line, colx=13))
            rendafixa.write(count, 14, sh.cell_value(rowx=line, colx=14))
            rendafixa.write(count, 15, sh.cell_value(rowx=line, colx=15))
            rendafixa.write(count, 16, sh.cell_value(rowx=line, colx=16))
            count +=1;
        elif(aux == "contas"):
            contas.write(count, 0, sh.cell_value(rowx=line, colx=0))
            contas.write(count, 1, sh.cell_value(rowx=line, colx=1))
            contas.write(count, 2, sh.cell_value(rowx=line, colx=2))
            contas.write(count, 3, sh.cell_value(rowx=line, colx=3))
            count +=1;
        elif(aux == "tesouraria"):
            tesouraria.write(count, 0, sh.cell_value(rowx=line, colx=0))
            tesouraria.write(count, 1, sh.cell_value(rowx=line, colx=1))
            tesouraria.write(count, 2, sh.cell_value(rowx=line, colx=2))
            tesouraria.write(count, 3, sh.cell_value(rowx=line, colx=3))
            count +=1;
        elif(aux == "patrimonio"):
            patrimonio.write(count, 0, sh.cell_value(rowx=line, colx=0))
            patrimonio.write(count, 1, sh.cell_value(rowx=line, colx=1))
            count +=1;
        elif(aux == "rentabilidade"):
            rentabilidade.write(count, 0, sh.cell_value(rowx=line, colx=0))
            rentabilidade.write(count, 1, sh.cell_value(rowx=line, colx=1))
            rentabilidade.write(count, 2, sh.cell_value(rowx=line, colx=2))
            rentabilidade.write(count, 3, sh.cell_value(rowx=line, colx=3))
            rentabilidade.write(count, 4, sh.cell_value(rowx=line, colx=4))
            rentabilidade.write(count, 5, sh.cell_value(rowx=line, colx=5))
            rentabilidade.write(count, 6, sh.cell_value(rowx=line, colx=6))
            rentabilidade.write(count, 7, sh.cell_value(rowx=line, colx=7))
            count +=1;

        #Validação para encontrar as tabelas
        if (sh.cell_value(rowx=line, colx=0) == "Futuros" and sh.cell_value(rowx=line+1, colx=0) == "Ativo"):
            aux = "futuros"
        elif (sh.cell_value(rowx=line, colx=0) == "Fundos de Investimentos - Outros Fundos" and sh.cell_value(rowx=line+1, colx=0) == "Código"):
            aux = "outrosfundos"
        elif (sh.cell_value(rowx=line, colx=0) == "Renda Fixa" and sh.cell_value(rowx=line+1, colx=0) == "Código"):
            aux = "rendafixa"
        elif (sh.cell_value(rowx=line, colx=0) == "Contas a Pagar/Receber" and sh.cell_value(rowx=line+1, colx=0) == "Descrição"):
            aux = "contas"
        elif (sh.cell_value(rowx=line, colx=0) == "Tesouraria" and sh.cell_value(rowx=line+1, colx=0) == "Descrição"):
            aux = "tesouraria"
        elif (sh.cell_value(rowx=line, colx=0) == "Patrimônio" and sh.cell_value(rowx=line+1, colx=0) == "Total do Patrimônio"):
            aux = "patrimonio"
        elif (sh.cell_value(rowx=line, colx=0) == "Rentabilidade Acumulada" and sh.cell_value(rowx=line+1, colx=0) == "Indexador"):
            aux = "rentabilidade"
        elif (sh.cell_value(rowx=line, colx=0) == ""):
            aux = "null"
            count = 0


    #Salva Arquivo
    wb.save(directory + '\\temp\\' + nomeArquivo)
#---------------------------------------------------------------------------------------------------------

#Percorre os arquivos e faz as manipulacoes
for filename in os.listdir(directory):
    if filename.endswith(".xls"): 
        print(os.path.join(directory, filename))
        manipulacao(filename)


