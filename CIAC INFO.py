import requests
import datetime
import csv

indexFile = input("caminho para o arquivo de índice: ")
outputFile = f'{input("caminho e nome do arquivo sem extensão: ")}.csv'

def getCanacsFromIndex(indexFileOrPath):
    """
    Dado um arquivo html contendo o índice de instituições de ensino, retirado do site
    www.anac.gov.br/educator/index2.aspx, esta função retornará uma lista contendo o código ANAC de todas
    as instituições de ensino constantes no referido índice.

    :param indexFileOrPath: Nome do .HTML contendo o índice de instituições, se na mesma pasta que este script
                            ou caminho deste arquivo .HTML.
    :return: Lista contendo o código ANAC (CANAC) de todas as instituições presentes no índice dado como parâmetro.
    """
    with open(indexFileOrPath, 'r', encoding='utf8') as fileHandle:
        data = fileHandle.readlines()

    CANACsList = []

    for line in data:
        if 'entidade.Asp?cod' in line:
            line = line.split("'")
            for part in line:
                if 'entidade.Asp?cod=' in part:
                    CANACsList.append(part.strip('entidade.Asp?cod='))

    return CANACsList


def genFirstLine():
    """
    Esta função gera a primeira linha do arquivo csv, correspondente às colunas da tabela.
    Neste formato, cada linha representa uma coluna.

    :return:Tupple contendo a primeira linha do arquivo .CSV, correspondente às colunas do mesmo.
    """
    return ("CANAC",
            "Nome Fantasia",
            "Razão Social",
            "CNPJ",
            "Situação",
            "Endereço",
            "Bairro",
            "Cidade",
            "UF",
            "Telefone",
            "Fax",
            "Site",
            "EMail",
            "Presidencia/Diretoria",
            "Coordenador",
            "Pedagogo",
            "Psicólogo",
            "Aeródromo",
            "Autorizado a funcionar até",
            "N DOU",
            "N DOU Renovação",
            "Primeiro curso a vencer",
            "Certificado de Piloto Aerodesportivo (prático)",
            "Comissário de voo (Semipresencial - Teórico/prático)",
            "Comissário de voo (Teórico/Prático)",
            "Despachante Operacional de Voo (IS-141-003A Teórico/Prático)",
            "Despachante Operacional de Voo (Semipresencial - Teórico/Prático)",
            "Despachante Operacional de Voo (Teórico/Prático)",
            "Instrutor Certificado de Piloto Aerodesportivo (Teórico)",
            "Instrutor de Voo de Avião (EAD - Teórico)",
            "Instrutor de Voo de Avião (Teórico)",
            "Instrutor de Voo de Avião (Prático)",
            "Instrutor de Voo de Helicóptero (EAD - Teórico)",
            "Instrutor de Voo de Helicóptero (Teórico)",
            "Instrutor de Voo de Helicóptero (Prático)",
            "Instrutor de Voo de Planador (Teórico/Prático)",
            "Instrutor de Voo de Planador (Teórico)",
            "Instrutor de Voo de Planador (Prático)",
            "Mecânico de Manutenção aeronáutica - Célula (Semipresencial - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Célula (IS 141-002B - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Célula (Semipresencial - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Célula (Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo Motopropulsor (Semipresencial - IS 141-002B - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo Motopropulsor (IS 141-002B - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo Motopropulsor (Semipresencial - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo Motopropulsor (Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo aviônicos (Semipresencial - IS 141-002B - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo aviônicos (IS 141-002B - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo aviônicos (Semipresencial - Teórico/Prático)",
            "Mecânico de Manutenção aeronáutica - Grupo aviônicos (Teórico/Prático)",
            "Piloto Agrícola de Aviões (Teórico/Prático)",
            "Piloto Agrícola de Helicópteros (Teórico/Prático)",
            "Piloto Comercial de Aviões (Prático)",
            "Piloto Comercial de Aviões + IFR (EAD - Teórico)",
            "Piloto Comercial de Helicópteros (EAD - Teórcio)",
            "Piloto Comercial de Helicópteros (Teórico)",
            "Piloto Comercial de Helicópteros (Prático)",
            "Piloto de Linha Aérea de Aviões (EAD - Teórico)",
            "Piloto de Linha Aérea de Aviões (Teórico)",
            "Piloto de Linha Aérea de Helicópteros (EAD - Teórico)",
            "Piloto de Linha Aérea de Helicópteros (Teórico)",
            "Piloto de Planador (Teórico)",
            "Piloto de Planador (Prático)",
            "Piloto de Recreio de Ultraleve (Teórico/Prático)",
            "Piloto de Recreio de Ultraleve (Teórico)",
            "Piloto de Recreio de Ultraleve (Prático)",
            "Piloto Desportivo de Ultraleve (Teórico/Prático)",
            "Piloto Desportivo de Ultraleve (Teórico)",
            "Piloto Desportivo de Ultraleve (Prático)",
            "Piloto Privado de Aviões (EAD - Teórico)",
            "Piloto Privado de Aviões (Teórico)",
            "Piloto Privado de Aviões (Prático)",
            "Piloto Privado de Helicópteros (EAD - Teórico)",
            "Piloto Privado de Helicópteros (Teórico)",
            "Piloto Privado de Helicópteros (Prático)",
            "Treinamento de Solo ATR42 (Teórico)",
            "Treinamento de Solo ATR42 - CMV (Teórico)",
            "TREINAMENTO DE SOLO SCHWEIZER 300  HU-30 (Teórico)",
            "Voo por Instrumentos (Sob capota - Prático)",
            "Voo por Instrumentos (EAD- Teórico)",
            "Voo por Instrumentos (Teórico)",
            "Voo por Instrumentos (IS 61-002D - Prático)",
            "Voo por Instrumentos Helicóptero (IS 61-002C - Prático)",
            "Voo por Instrumentos Helicóptero (IS 61-002D - Prático)")


def parseInfoByCANAC(canac):
    """
    Dado o código ANAC (CANAC) de um centro de instrução, esta função irá baixar as informações relativas a este CIAC
    do site da ANAC, os mapeará para variáveis e retornará uma Tupple no formato adequado à composição do arquivo .CSV
    conforme as colunas configuradas na função genFirstLine.

    :param canac: Código ANAC do CIAC que se deseja buscar.
    :return: Tupple  contendo todas as informações que a a ANAC disponibiliza acerca do CIAC
    """

    # Faz o download da página com as informações do CIAC
    r = requests.get(f'https://sistemas.anac.gov.br/educator/entidade.Asp?cod={canac}', verify=False)
    t = r.text.splitlines()

    # Inicializa todas as variáveis
    CVDates = []

    CNPJ = ''
    CANAC = ''
    nomeFantasia = ''
    situacao = ''
    razao = ''
    endereco = ''
    bairro = ''
    cidade = ''
    UF = ''
    telefone = ''
    fax = ''
    site = ''
    email = ''
    pres = ''
    coord = ''
    pedagogo = ''
    psicologo = ''
    ad = ''
    afa = ''
    ndou = ''
    ndour = ''
    CPAp = ''
    COMtp = ''
    COMSPtp = ''
    DOVIS141tp = ''
    DOVSPtp = ''
    DOVtp = ''
    ICPAt = ''
    INVAEADt = ''
    INVAt = ''
    INVAp = ''
    INVHEADt = ''
    INVHt = ''
    INVHp = ''
    INVPLtp = ''
    INVPLt = ''
    INVPLp = ''
    MMACSP141tp = ''
    MMAC141tp = ''
    MMACSPtp = ''
    MMACtp = ''
    MMAGMPSP141tp = ''
    MMAGMP141tp = ''
    MMAGMPSPtp = ''
    MMAGMPtp = ''
    MMAASP141tp = ''
    MMAA141tp = ''
    MMAASPtp = ''
    MMAAtp = ''
    PAGAtp = ''
    PAGHtp = ''
    PCAp = ''
    PCAIFREADt = ''
    PCHEADt = ''
    PCHt = ''
    PCHp = ''
    PLAEADt = ''
    PLAt = ''
    PLAHEADt = ''
    PLAHt = ''
    PPLt = ''
    PPLp = ''
    PRULtp = ''
    PRULt = ''
    PRULp = ''
    PDULtp = ''
    PDULt = ''
    PDULp = ''
    PPAEADt = ''
    PPAt = ''
    PPAp = ''
    PPHEADt = ''
    PPHt = ''
    PPHp = ''
    ATR42t = ''
    ATR42CMVt = ''
    schweizer300t = ''
    IFRCapp = ''
    IFREADt = ''
    IFRt = ''
    IFRA61p = ''
    IFRH61Cp = ''
    IFRH61Dp = ''

    # Mapeia todas as informações disponíveis no site para as variáveis adequadas
    for lineIndex in range(len(t)):
        line = t[lineIndex]
        if 'CNPJ:&nbsp;' in line:
            CNPJ = line[((line.index('<b>') + 3) + 3):line.index('</B>')]
        elif 'Código ANAC:&nbsp;<b>' in line:
            CANAC = line[(line.index('<b>') + 3):line.index('</B>')]
        elif 'NomeFantasia:&nbsp;<B>' in line:
            nomeFantasia = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Situação:&nbsp;' in line:
            a = t[(lineIndex + 1)]
            situacao = a[(a.index('<b>') + 3):a.index('</b>')]
            del a
        elif 'Razão Social:&nbsp;<B>' in line:
            razao = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Endereço:&nbsp;' in line:
            endereco = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Bairro:&nbsp;' in line:
            bairro = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Cidade:&nbsp;' in line:
            cidade = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'UF:&nbsp;' in line:
            UF = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Telefone:&nbsp;' in line:
            telefone = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Fax:&nbsp;' in line:
            fax = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Internet:&nbsp;' in line:
            site = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'E-mail:&nbsp;' in line:
            email = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Presidente/Diretor:&nbsp;' in line:
            pres = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Coordenador:&nbsp;' in line:
            coord = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Pedagogo:&nbsp;' in line:
            pedagogo = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Psicólogo:&nbsp;' in line:
            psicologo = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Aeródromo:&nbsp;' in line:
            ad = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Autorizada a funcionar Até:&nbsp;' in line:
            afa = datetime.datetime.strptime(line[(line.index('<b>') + 3):line.index('</B>')], '%d/%m/%Y')
        elif 'Nº do DOU:&nbsp;' in line:
            ndou = line[(line.index('<B>') + 3):line.index('</B>')]
        elif 'Nº do DOU (Renovação):&nbsp;' in line:
            ndour = line[(line.index('<B>') + 3):line.index('</B>')]

        elif 'CERTIFICADO DE PILOTO AERODESPORTIVO' in line:
            if 'Prático' in t[(lineIndex + 1)]:
                try:
                    CPAp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(CPAp)
                except IndexError:
                    CPAp = t[(lineIndex + 4)]

        elif 'COMISSÁRIO DE VOO - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    COMSPtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(COMSPtp)
                except IndexError:
                    COMSPtp = t[(lineIndex + 4)]

        elif 'COMISSÁRIO DE VOO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    COMtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(COMtp)
                except IndexError:
                    COMtp = t[(lineIndex + 4)]

        elif 'DESPACHANTE OPERACIONAL DE VOO (IS 141-003A)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    DOVIS141tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(DOVIS141tp)
                except IndexError:
                    DOVIS141tp = t[(lineIndex + 4)]

        elif 'DESPACHANTE OPERACIONAL DE VOO - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    DOVSPtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(DOVSPtp)
                except IndexError:
                    DOVSPtp = t[(lineIndex + 4)]

        elif 'DESPACHANTE OPERACIONAL DE VOO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    DOVtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(DOVtp)
                except IndexError:
                    DOVtp = t[(lineIndex + 4)]

        elif 'INSTRUTOR CERTIFICADO DE PILOTO AERODESPORTIVO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    ICPAt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(ICPAt)
                except IndexError:
                    ICPAt = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO AVIÃO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    INVAEADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVAEADt)
                except IndexError:
                    INVAEADt = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO AVIÃO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    INVAt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVAt)
                except IndexError:
                    INVAt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    INVAp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVAp)
                except IndexError:
                    INVAp = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    INVHEADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVHEADt)
                except IndexError:
                    INVHEADt = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex + 1)]:
                try:
                    INVHt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVHt)
                except IndexError:
                    INVHt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex + 1)]:
                try:
                    INVHp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVHp)
                except IndexError:
                    INVHp = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO PLANADOR' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    INVPLtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVPLtp)
                except IndexError:
                    INVPLtp = t[(lineIndex + 4)]

            elif 'Teórico' in t[(lineIndex+1)]:
                try:
                    INVPLt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVPLt)
                except IndexError:
                    INVPLt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    INVPLp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(INVPLp)
                except IndexError:
                    INVPLp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO CÉLULA - MODALIDADE SEMIPRESENCIAL (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMACSP141tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMACSP141tp)
                except IndexError:
                    MMACSP141tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO CÉLULA - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMACSPtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMACSPtp)
                except IndexError:
                    MMACSPtp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Célula (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAC141tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAC141tp)
                except IndexError:
                    MMAC141tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Célula' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMACtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMACtp)
                except IndexError:
                    MMACtp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO MOTOPROPULSOR - MODALIDADE SEMIPRESENCIAL (IS 141-002B)'\
                in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAGMPSP141tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAGMPSP141tp)
                except IndexError:
                    MMAGMPSP141tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO MOTOPROPULSOR - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAGMPSPtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAGMPSPtp)
                except IndexError:
                    MMAGMPSPtp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Motopropulsor (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAGMP141tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAGMP141tp)
                except IndexError:
                    MMAGMP141tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Motopropulsor' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAGMPtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAGMPtp)
                except IndexError:
                    MMAGMPtp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - MÓDULO AVIÔNICOS - MODALIDADE SEMIPRESENCIAL (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAASP141tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAASP141tp)
                except IndexError:
                    MMAASP141tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - MÓDULO AVIÔNICOS - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAASPtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAASPtp)
                except IndexError:
                    MMAASPtp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Aviônicos (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAA141tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAA141tp)
                except IndexError:
                    MMAA141tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Aviônicos' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    MMAAtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(MMAAtp)
                except IndexError:
                    MMAAtp = t[(lineIndex + 4)]

        elif 'PILOTO AGRÍCOLA DE AVIÃO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    PAGAtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PAGAtp)
                except IndexError:
                    PAGAtp = t[(lineIndex + 4)]

        elif 'PILOTO AGRÍCOLA DE HELICÓPTERO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    PAGHtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PAGHtp)
                except IndexError:
                    PAGHtp = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL (AVIÃO)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    PCAp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PCAp)
                except IndexError:
                    PCAp = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL DE AVIÃO/IFR - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PCAIFREADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PCAIFREADt)
                except IndexError:
                    PCAIFREADt = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL DE HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PCHEADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PCHEADt)
                except IndexError:
                    PCHEADt = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL DE HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PCHt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PCHt)
                except IndexError:
                    PCHt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    PCHp = t[(lineIndex+4)].split()[2]
                    CVDates.append(PCHp)
                except IndexError:
                    PCHp = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE AVIÃO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PLAEADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PLAEADt)
                except IndexError:
                    PLAEADt = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE AVIÃO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PLAt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PLAt)
                except IndexError:
                    PLAt = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PLAHEADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PLAHEADt)
                except IndexError:
                    PLAHEADt = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PLAHt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PLAHt)
                except IndexError:
                    PLAHt = t[(lineIndex + 4)]

        elif 'PILOTO DE PLANADOR' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PPLt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPLt)
                except IndexError:
                    PPLt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    PPLp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPLp)
                except IndexError:
                    PPLp = t[(lineIndex + 4)]

        elif 'PILOTO DE RECREIO DE ULTRALEVE' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    PRULtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PRULtp)
                except IndexError:
                    PRULtp = t[(lineIndex + 4)]

            elif 'Teórico' in t[(lineIndex+1)]:
                try:
                    PRULt = t[(lineIndex+4)].split()[2]
                    CVDates.append(PRULt)
                except IndexError:
                    PRULt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    PRULp = t[(lineIndex+4)].split()[2]
                    CVDates.append(PRULp)
                except IndexError:
                    PRULp = t[(lineIndex + 4)]

        elif 'PILOTO DESPORTIVO DE ULTRALEVE' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    PDULtp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PDULtp)
                except IndexError:
                    PDULtp = t[(lineIndex + 4)]

            elif 'Teórico' in t[(lineIndex+1)]:
                try:
                    PDULt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PDULt)
                except IndexError:
                    PDULt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    PDULp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PDULp)
                except IndexError:
                    PDULp = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE AVIÃO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PPAEADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPAEADt)
                except IndexError:
                    PPAEADt = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE AVIÃO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PPAt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPAt)
                except IndexError:
                    PPAt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    PPAp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPAp)
                except IndexError:
                    PPAp = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PPHEADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPHEADt)
                except IndexError:
                    PPHEADt = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    PPHt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPHt)
                except IndexError:
                    PPHt = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    PPHp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(PPHp)
                except IndexError:
                    PPHp = t[(lineIndex + 4)]

        elif 'TREINAMENTO DE SOLO ATR-42 - CMV' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    ATR42CMVt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(ATR42CMVt)
                except IndexError:
                    ATR42CMVt = t[(lineIndex + 4)]

        elif 'TREINAMENTO DE SOLO ATR-42' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    ATR42t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(ATR42CMVt)
                except IndexError:
                    ATR42t = t[(lineIndex + 4)]

        elif 'TREINAMENTO DE SOLO SCHWEIZER 300' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    schweizer300t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(schweizer300t)
                except IndexError:
                    schweizer300t = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS (SOB CAPOTA)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    IFRCapp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(IFRCapp)
                except IndexError:
                    IFRCapp = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    IFREADt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(IFREADt)
                except IndexError:
                    IFREADt = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS AVIÃO (IS 61-002D)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    IFRA61p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(IFRA61p)
                except IndexError:
                    IFRA61p = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS HELICÓPTERO (IS 61-002C)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    IFRH61Cp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(IFRH61Cp)
                except IndexError:
                    IFRH61Cp = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS HELICÓPTERO (IS 61-002D)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    IFRH61Dp = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(IFRH61Dp)
                except IndexError:
                    IFRH61Dp = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS' in line:
            if 'Teórico' in t[(lineIndex + 1)]:
                try:
                    IFRt = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    CVDates.append(IFRt)
                except IndexError:
                    IFRt = t[(lineIndex + 4)]

    # Verifica qual curso está mais próximo de vencer.
    IDates = []
    for i in CVDates:
        if type(i) == datetime.datetime:
            IDates.append(i)

    try:
        PrCaV = min(IDates)
    except ValueError:
        PrCaV = ''

    # Retorna uma string no formato especificado em genFirstLine, correspondente a cada CIAC

    return (CANAC,
            nomeFantasia,
            razao,
            CNPJ,
            situacao,
            endereco,
            bairro,
            cidade,
            UF,
            telefone,
            fax,
            site,
            email,
            pres,
            coord,
            pedagogo,
            psicologo,
            ad,
            afa,
            ndou,
            ndour,
            PrCaV,
            CPAp,
            COMSPtp,
            COMtp,
            DOVIS141tp,
            DOVSPtp,
            DOVtp,
            ICPAt,
            INVAEADt,
            INVAt,
            INVAp,
            INVHEADt,
            INVHt,
            INVHp,
            INVPLtp,
            INVPLt,
            INVPLp,
            MMACSP141tp,
            MMAC141tp,
            MMACSPtp,
            MMACtp,
            MMAGMPSP141tp,
            MMAGMP141tp,
            MMAGMPSPtp,
            MMAGMPtp,
            MMAASP141tp,
            MMAA141tp,
            MMAASPtp,
            MMAAtp,
            PAGAtp,
            PAGHtp,
            PCAp,
            PCAIFREADt,
            PCHEADt,
            PCHt,
            PCHp,
            PLAEADt,
            PLAt,
            PLAHEADt,
            PLAHt,
            PPLt,
            PPLp,
            PRULtp,
            PRULt,
            PRULp,
            PDULtp,
            PDULt,
            PDULp,
            PPAEADt,
            PPAt,
            PPAp,
            PPHEADt,
            PPHt,
            PPHp,
            ATR42t,
            ATR42CMVt,
            schweizer300t,
            IFRCapp,
            IFREADt,
            IFRt,
            IFRA61p,
            IFRH61Cp,
            IFRH61Dp)


CANACs = getCanacsFromIndex(indexFile)
print(f'{len(CANACs)} CANACs retirados do arquivo de índice.')


with open(outputFile, 'w', newline='', encoding='utf8') as fileHandle:
    csvWriter = csv.writer(fileHandle, dialect='excel')
    csvWriter.writerow(genFirstLine())
    for CIACCANAC in CANACs:
        csvWriter.writerow(parseInfoByCANAC(CIACCANAC))
