import csv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
# Módulos requeridos: python-docx, jinja2
import os
import subprocess
import datetime
import win32com.client



def docx2pdf(docxCaminho, WordApp):
    """
    Dado o caminho absoluto para um arquivo .docx, e uma instância de um COMObject para o MS Word, essa função salvará
    uma cópia do arquivo original em formato pdf na mesma pasta em que estava o arquivo original.
    Args:
        docxCaminho: string contendo o caminho absoluto para o arquivo .docx a ser convertido.
        WordApp: COMObject Word Application (use win32com.client.DispatchEx('Word.Application'))
    """
    saida = docxCaminho[:-4] + 'pdf'  # Deriva o caminho e nome do arquivo PDF do arquivo .docx original
    doc = WordApp.Documents.Open(docxCaminho, ReadOnly=1)  # Abre o arquivo .docx no Word em modo de leitura
    doc.SaveAs(saida, FileFormat=17)  # Número 17 equivale ao formato PDF
    doc.Close()  # Fecha o arquivo gerado


def csvParaFichasDeAvaliacao(arquivoCsv):
    """
    Dado o caminho absoluto para um arquivo .csv contendo as informações necessárias, essa função gerará as fichas de
    avaliação de voo para todos os voos constantes no arquivo csv.
    A função está configurada de forma a ler um arquivo csv produzido pelo MS Excel com codificação UTF-8
    (opção salvar como > CSV UTF-8 (Delimitado por vírgulas) (*.csv);

    Args:
        arquivoCsv: string contendo o caminho absoluto para o arquivo csv a ser usado como base.

        Para a função ser efetiva, o arquivo csv deve obedecer aos seguintes padrões:
        Na primeira linha:
            A célula A1 deve conter o nome do curso onde as fichas serão utilizadas (ex.:PP, PC, PC/IFR);
            A célula B1 deve conter a etapa e a fase descrita no arquivo CSV (idealmente uma por arquivo csv)
                (ex.: Fase 1, Etapa 1 Fase 1);
            As demais células da primeira linha devem corresponder ao nome que descreve o voo (ex.: 1, 2, N1, X1, X2...)
                cada célula na primeira linha deve conter apenas um voo;

        Nas N células da primeira coluna:
            A primeira linha/célula (A1) deve ter o conteúdo expresso acima;
            As demais linhas/células devem ser preenchidas com os exercícios a serem avaliados nessa fase do curso;

        Nas N células da segunda coluna:
            A primeira linha/célula (B1) deve ter o conteúdo expresso acima;
            As demais linhas/células devem ser preenchidas com os conteúdos a serem estutados antes da execução do
                exercício, separados apenas por uma vírgula e um espaço (", ")
                (ex.: "operação da aeronave, regras de tráfego aéreo, circuito de tráfego")
                (Os conteúdos a serem previamente estudados podem ser repetidos - o ideal é que cada critério conte
                com o seu estudo prévio, mesmo que os itens a serem estudados já tenham sido repetidos em outros
                exercícios - o script se encarregará de não repetí-los em um único voo, desde que tenham sido digitados
                de forma identica)

        Nas N células das colunas que correspondem a cada voo:
            A primeira linha ou célula (de C1 até a coluna necessária para compreender todos os voos) devem corresponder
                ao nome que descreve o voo (ex.: 1, 2, N1, X1, X2...)
                (cada célula na primeira linha deve conter apenas um voo);

            As demais linhas/células devem ser preenchidas com o nível de aprendizado requerido para cada exercício ou
                deixado em branco, caso o exercício não seja avaliado naquele voo em específico.

            Ao final dessas colunas, devem ser adicionadas mais cinco linhas, que corresponderão, respectivamente ao(s):
                1. Aeródromos envolvidos nessa missão, separados por ponto e vírgula espaço (ex.: SBGO;SBCN);
                2. Rota a ser executada no exercício, tal qual o campo de rota do PVC (ex.: "SWNS SBR606 SWNS")
                3. O tipo de voo, DC (para duplo comando) ou Solo.
                     (O script apenas diferenciará dc ou DC, os demais serão considerados como Solo.)
                4. Duração do voo, podendo ser apenas a quantidade de horas, no caso de horas inteiras (ex.: 1, 2, 3...)
                     ou, no caso de horas incompletas, preencher no formato h:mm (ex.: 1:30, 2:30...);
                5. Número de pousos previstos durante a execução deste voo.

    """
    with open(arquivoCsv, encoding='utf-8-sig', newline='') as arquivoCsv:  # Abre o arquivo CSV
        leitorCsv = csv.reader(arquivoCsv, dialect='excel', delimiter=';')  # Gera um objeto leitor de CSV
        dados = [row for row in leitorCsv]  # Extrai os dados do arquivo CSV para a lista "dados"

    # Itera entre as células da segunda coluna, transformando a string estudos prévios em lista de tópicos
    # a serem estudados
    for exercicio in dados[1:-5]:
        if ',' in exercicio[1]:
            exercicio[1] = exercicio[1].split(', ')
        else:
            exercicio[1] = [exercicio[1]]

    voos = {}  # Cria um dicionário de voos - as chaves correspondem ao número do voo dentro da fase de treinamento
    for i in range(len(dados[0][2:])):
        vooNum = dados[0][i+2]
        vooTipo = dados[-3][i+2]
        voos[vooNum] = {}

        # Cria um dicionário para cada voo.
        # Cada dicionário contém informações necessárias para o preenchimento de uma ficha de voo:
        #       O curso sendo desenvolvido, a etapa e a fase do curso a que pertence o voo, o número do voo,
        #       Os aeródromos envolvidos na operação, o tipo de voo, a instrução teórica a ser dada no briefing,
        #       A rota, o tempo previsto de voo, o número previsto de pousos, os conteúdos a serem introduzidos (Diff)
        #       e a lista de dicionários contendo os exercícios e o nível de aprendizagem requerido para cada um.

        voos[vooNum]['course'] = dados[0][0]
        voos[vooNum]['etapaFase'] = dados[0][1]
        voos[vooNum]['FNum'] = 'Voo ' + vooNum
        voos[vooNum]['ADIcao'] = dados[-5][i+2]
        voos[vooNum]['TypeOfFlight'] = 'Duplo Comando' if vooTipo.upper() == 'DC' else 'Solo'
        voos[vooNum]['TheoricalInstruction'] = 'Durante o briefing, o instrutor deverá conferir a documentação ' \
                                               'de porte obrigatório, juntamente com o aluno, verificar se este está ' \
                                               'apto ao voo, repassar as rotinas operacionais esperadas e descrever, ' \
                                               'brevemente, as manobras a serem executadas durante o voo.'

        if (23 + len(voos[vooNum]['TheoricalInstruction']) * 1.05) % 75 > 0:
            voos[vooNum]['TheoricalInstructionExt'] = (23 + len(voos[vooNum]['TheoricalInstruction'])*1.05)//75+1
        else:
            voos[vooNum]['TheoricalInstructionExt'] = (23 + len(voos[vooNum]['TheoricalInstruction'])*1.05)/75

        voos[vooNum]['Route'] = dados[-4][i+2]

        try:
            voos[vooNum]['PNLands'] = ('/' + str(int(dados[-1][i+2])))
        except ValueError:
            voos[vooNum]['PNLands'] = ''

        if len(dados[-2][i+2]) == 1:
            voos[vooNum]['PTime'] = (dados[-2][i + 2] + ':00')
        elif len(dados[-2][i+2]) == 2 and dados[-2][i+2][0] == '0':
            voos[vooNum]['PTime'] = (dados[-2][i + 2][1] + ':00')
        elif len(dados[-2][i+2]) == 5 and dados[-2][i+2][0] == '0':
            voos[vooNum]['PTime'] = (dados[-2][i + 2][1:])
        else:
            voos[vooNum]['PTime'] = dados[-2][i+2]

        voos[vooNum]['Diff'] = []

    chavesVoos = [x for x in voos.keys()]  # Gera uma lista com as chaves do dicionário voos

    # Itera entre as colunas da tabela, captando informações de estudos prévios necessários à execução da lição de voo,
    # bem como os exercícios e seus níveis de aprendizagem requeridos para cada voo e anexando-os ao dicionário
    # apropriado a cada voo.

    for x in range(len(dados[0])-2):
        listaE = []
        listaExercicios = []
        listaEstudosPrevios = []

        for y in range(len(dados)-6):
            if dados[(1+y)][(2+x)] != '':
                listaE.append({'Exercise': dados[1+y][0].capitalize(), 'KL': dados[(1+y)][(2+x)].upper()})
                listaExercicios.append(dados[1+y][0])

                for element in dados[1+y][1]:
                    if element not in listaEstudosPrevios:
                        listaEstudosPrevios.append(element)

        voos[dados[0][(x+2)]]['ExList'] = listaE.copy()
        voos[dados[0][(x+2)]]['EList'] = sorted(listaExercicios)
        ps = ', '.join(sorted(listaEstudosPrevios)) + '.'
        voos[dados[0][(x+2)]]['PreviousStudy'] = ps[:ps.rfind(', ')] + ' e ' + ps[(ps.rfind(', ')+2):]

        if ((19 + len(voos[dados[0][(x + 2)]]['PreviousStudy'])*1.05) % 75) > 0:
            voos[dados[0][(x + 2)]]['PreviousStudyExt'] = (19 + len(voos[dados[0][(x+2)]]['PreviousStudy'])*1.05)//75 \
                                                          + 1
        else:
            voos[dados[0][(x + 2)]]['PreviousStudyExt'] = (19 + len(voos[dados[0][(x+2)]]['PreviousStudy'])*1.05)/75

    # Para cada voo, verifica se há exercícios novos (não presentes em lições anteriores) sendo introduzidos
    for i in range(1, len(chavesVoos)):
        vooAnterior = []
        p = i
        while p > 0:
            p -= 1
            vooAnterior.extend(voos[chavesVoos[p]]['EList'])

        novoNoVooAtual = voos[chavesVoos[i]]['EList']
        for exerc in vooAnterior:
            try:
                novoNoVooAtual.remove(exerc)
            except ValueError:
                continue

        if novoNoVooAtual:
            voos[chavesVoos[i]]['Diff'] = novoNoVooAtual

    # Gera o objetivo da missão baseado no nome do voo e na existência ou não de conteúdos novos sendo introduzidos
    for i in range(len(chavesVoos)):
        if i == 0:
            voos[chavesVoos[i]]['Objectives'] = 'introduzir o aluno à nova fase de voos, iniciando a partir das manobras e ' \
                                           'procedimentos listados nas seções a seguir.'
        elif 'X' in voos[chavesVoos[i]]['FNum'].upper():
            voos[chavesVoos[i]]['Objectives'] = 'verificar o aprendizado do aluno em relação às manobras e exercícios ' \
                                           'introduzidos e treinados ao longo dessa fase do treinamento.'
        elif 'N' in voos[chavesVoos[i]]['FNum'].upper() and 'NV' not in voos[chavesVoos[i]]['FNum'].upper():
            voos[chavesVoos[i]]['Objectives'] = 'promover a prática dos exercícios treinados ao longo desta fase ' \
                                           'no período noturno.'
        elif voos[chavesVoos[i]]['Diff']:
            voos[chavesVoos[i]]['Objectives'] = ',além de praticar os procedimentos aprendidos em lições passadas, ' \
                                          'introduzir o aluno à execução de procedimentos relativos a ' + \
                                           ', '.join(sorted(voos[chavesVoos[i]]['Diff'])).lower() + '.'

            lci = voos[chavesVoos[i]]['Objectives'].rfind(', ')
            voos[chavesVoos[i]]['Objectives'] = voos[chavesVoos[i]]['Objectives'][:lci] + ' e ' + \
                                          voos[chavesVoos[i]]['Objectives'][(lci+2):]
        else:
            voos[chavesVoos[i]]['Objectives'] = 'promover a prática dos exercícios treinados ao longo desta fase do ' \
                                          'treinamento.'

        # Calcula a quantidade de linhas a serem utilizadas pelo objetivo da missão na ficha de avaliação de voo
        if ((29 + len(voos[chavesVoos[i]]['Diff'])*1.05) % 75) > 0:
            voos[chavesVoos[i]]['DiffExt'] = (29 + len(voos[chavesVoos[i]]['Diff'])*1.05)//75 + 1
        else:
            voos[chavesVoos[i]]['DiffExt'] = (29 + len(voos[chavesVoos[i]]['Diff']*1.05)) / 75

    for voo in chavesVoos:
        # Calcula o número de linhas utilizadas pelos campos objetivo, estudos prévios e instrução teórica
        linhasExistentes = voos[voo]['DiffExt'] + voos[voo]['PreviousStudyExt'] + voos[voo]['TheoricalInstructionExt']
        linhasDispPrimPag = 25 - linhasExistentes
        linhasNecessarias = len(voos[voo]['ExList'])

        # Determina o modelo que oferece melhor racionalização do espaço baseado no número de linhas disponíveis na
        # primeira página e o número de linhas necessárias para comportar a lista de exercícios a serem avaliados.
        if (linhasNecessarias <= linhasDispPrimPag) or ((linhasNecessarias // 44) > 0
                                                        and (linhasNecessarias % 44) <= linhasDispPrimPag):
            docModelo = 'Modelo0.docx'
        else:
            docModelo = 'Modelo1.docx'

        # Determina o nome do arquivo a ser gerado
        nomeDoArquivo = f"Ficha de avaliação de Voo - {voos[voo]['course']} - {voos[voo]['etapaFase']} -" \
                        f" {voos[voo]['FNum']}.docx"
        #documento = MailMerge(docModelo)  # Determina o modelo a ser utilizado
        # Passa o dicionário do voo, com os valores a serem utilizados para o módulo que o preencherá
        #documento.merge(**voos[voo])
        # Passa a lista de exercícios do voo para o preenchimento da lista de exercícios no modelo
        #documento.merge_rows('Exercise', voos[voo]['ExList'])
        #documento.write(nomeDoArquivo)  # Executa o preenchimento
        # print(f'{nomeDoArquivo} gerada')  # Mostra o progresso do loop
        documento = DocxTemplate(docModelo)
        img = InlineImage(documento, (os.getcwd()+'\\Logo.jpg'), width=Mm(37), height=Mm(36))
        voos[voo]['Logo'] = img
        documento.render(voos[voo])
        documento.save(nomeDoArquivo)


# Caso o script seja executado diretamente do terminal, ao invés de ser importado como um módulo, ele irá presumir que
# todos os arquivos com extensão .csv dentro da pasta estão no formato desejado para sua execução, e tentará gerar
# fichas de avaliação de voo a partir deles.

if __name__ == '__main__':
    docInicio = datetime.datetime.now()  # Registra o tempo de início da execução do script
    diretorio = os.listdir(os.getcwd())  # Gera uma lista com todos os arquivos dentro do diretório

    for arquivo in diretorio:  # Para todos os arquivos csv no diretório, executa a função csvParaFichasDeAvaliação
        if arquivo[-4:].lower() == '.csv' and 'example' not in arquivo and '#' not in arquivo:
            csvParaFichasDeAvaliacao(arquivo)

    deltaT = str(datetime.datetime.now() - docInicio).split(':')  # Calcula o tempo gasto para gerar os arquivos .docx
    contagem = 0

    for arquivo in os.listdir(os.getcwd()):  # Conta o número de arquivos gerados
        if arquivo[-5:] == '.docx' and 'model' not in arquivo.lower():
            contagem += 1

    print('\n\n')
    # Mostra o tempo gasto na geração dos arquivos .docx
    print(f"Geração de {contagem} fichas de avaliação de voo terminada em "
          f"{f'{int(deltaT[1])} minuto' if int(deltaT[1]) > 0 else ''}"
          f"{'s' if int(deltaT[1]) > 1 else ''}{' e ' if int(deltaT[1]) > 0 and round(float(deltaT[2]), 1) > 0 else ''}"
          f"{f'{round(float(deltaT[2]), 1)} segundo' if float(deltaT[2]) > 0 else ''}"
          f"{'s' if round(float(deltaT[2]), 1) > 1 else ''}")

    #Conversão para PDF
    try:
        diretorio = os.listdir(os.getcwd())
        # Gera o COMObject Word.Application necessário à função docx2pdf
        msWord = win32com.client.DispatchEx('Word.Application')
        contagemFichas = 0
        conversaoPDF = datetime.datetime.now()  # Registra o início da conversão dos arquivos
        for docx in diretorio:  # Para cada arquivo .docx no diretório que não os modelos, executa a função docx2pdf
            if docx[-5:].lower() == '.docx' and 'model' not in docx.lower():
                docx2pdf(f'{os.getcwd()}\\{docx}', msWord)
                contagemFichas += 1
                # print(f'{docx[:-5]}.pdf gerada')  # Mostra o progresso da conversão
        deltaT = str(datetime.datetime.now() - conversaoPDF).split(':')  # Calcula o tempo necessário para a conversão
        print(f"Conversão para .pdf terminada em {f'{int(deltaT[1])} minuto' if int(deltaT[1]) > 0 else ''}"
              f"{'s' if int(deltaT[1]) > 1 else ''}"
              f"{' e ' if int(deltaT[1]) > 0 and round(float(deltaT[2]), 1) > 0 else ''}"
              f"{f'{round(float(deltaT[2]), 1)} segundo' if float(deltaT[2]) > 0 else ''}"
              f"{'s' if round(float(deltaT[2]), 1) > 1 else ''}")  # Mostra a duração do processo de conversão

    except:
        print(f"Não foi possível fazer a conversão de {docx} para PDF")
