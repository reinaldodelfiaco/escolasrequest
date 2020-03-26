import requests
import datetime
import csv
from bs4 import BeautifulSoup
import re


def get_canacs_from_index():
    """
    Esta função se comunicará com o site www.anac.gov.br/educator/index2.aspx, solicitará as informações de todas as
     escolas de aviação cadastradas no índice de educadores e retornará uma lista contendo o código ANAC de todas
    as instituições de ensino constantes no referido índice.

    :param
    :return: Lista contendo o código ANAC (CANAC) de todas as instituições presentes no índice de educadores da ANAC.
    """
    print(f'\n[{datetime.datetime.now().strftime("%H:%M:%S")}] Recuperando CANACs do índice de educadores')
    process_start = datetime.datetime.now()
    requests_session = requests.Session()
    request = requests_session.get('https://sistemas.anac.gov.br/educator/Index2.aspx')
    parsed_html_get_response = BeautifulSoup(request.text, 'html.parser')

    view_state = parsed_html_get_response.find(id='__VIEWSTATE').attrs['value']
    view_state_generator = parsed_html_get_response.find(id='__VIEWSTATEGENERATOR').attrs['value']
    event_validation = parsed_html_get_response.find(id='__EVENTVALIDATION').attrs['value']

    data_to_post = {'__VIEWSTATE': view_state,
                    '__VIEWSTATEGENERATOR': view_state_generator,
                    '__EVENTVALIDATION': event_validation,
                    'ctl00$ContentPlaceHolder1$Codigo': '',
                    'ctl00$ContentPlaceHolder$CNPJ': '',
                    'ctl00$ContentPlaceHolder$NomeEscola': '',
                    'ctl00$ContentPlaceHolder$Cidade': '',
                    'ctl00$ContentPlaceHolder$UF': 'Todos',
                    'ctl00$ContentPlaceHolder$ListaCurso': 'Todos',
                    'ctl00$ContentPlaceHolder$Button1': 'Buscar', }

    post_response = requests_session.post('https://sistemas.anac.gov.br/educator/Index2.aspx', data=data_to_post)
    parsed_html_post_response = BeautifulSoup(post_response.text, 'html.parser')

    appropriate_wildcard = re.compile('ContentPlaceHolder1_DataList1_Cod_EscolaLabel_')
    canac_list = [tag.contents[0] for tag in parsed_html_post_response.findAll("span", {'id': appropriate_wildcard})]

    delta_t = str(datetime.datetime.now() - process_start).split(':')
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] - Esse processo encontrou {len(canac_list)} escolas e levou "
          f"{f'{int(delta_t[1])} minuto' if int(delta_t[1]) > 0 else ''}"
          f"{'s' if int(delta_t[1]) > 1 else ''}"
          f"{' e ' if int(delta_t[1]) > 0 and round(float(delta_t[2]), 1) > 0 else ''}"
          f"{f'{round(float(delta_t[2]), 1)} segundo' if float(delta_t[2]) > 0 else ''}"
          f"{'s' if round(float(delta_t[2]), 1) > 1 else ''}\n")

    return canac_list


def gen_first_line():
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


def parse_info_by_canac(canac):
    """
    Dado o código ANAC (CANAC) de um centro de instrução, esta função irá baixar as informações relativas a este CIAC
    do site da ANAC, os mapeará para variáveis e retornará uma Tupple no formato adequado à composição do arquivo .CSV
    conforme as colunas configuradas na função genFirstLine.

    :param canac: Código ANAC do CIAC que se deseja buscar.
    :return: Tupple  contendo todas as informações que a a ANAC disponibiliza acerca do CIAC
    """

    # Faz o download da página com as informações do CIAC
    r = requests.get(f'https://sistemas.anac.gov.br/educator/entidade.Asp?cod={canac}')
    t = r.text.splitlines()

    # Inicializa todas as variáveis
    cv_dates = []

    cnpj = ''
    canac = ''
    nome_fantasia = ''
    situacao = ''
    razao = ''
    endereco = ''
    bairro = ''
    cidade = ''
    uf = ''
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
    cpa_p = ''
    com_tp = ''
    comsp_tp = ''
    dovis141_tp = ''
    dovsp_tp = ''
    dov_tp = ''
    icpa_t = ''
    invaead_t = ''
    inva_t = ''
    inva_p = ''
    invhead_t = ''
    invh_t = ''
    invh_p = ''
    invpl_tp = ''
    invpl_t = ''
    invpl_p = ''
    mmacsp141_tp = ''
    mmac141_tp = ''
    mmacsp_tp = ''
    mmac_tp = ''
    mmagmpsp141_tp = ''
    mmagmp141_tp = ''
    mmagmpsp_tp = ''
    mmagmp_tp = ''
    mmaasp141_tp = ''
    mmaa141_tp = ''
    mmaasp_tp = ''
    mmaa_tp = ''
    paga_tp = ''
    pagh_tp = ''
    pca_p = ''
    pcaifread_t = ''
    pchead_t = ''
    pch_t = ''
    pch_p = ''
    plaead_t = ''
    pla_t = ''
    plahead_t = ''
    plah_t = ''
    ppl_t = ''
    ppl_p = ''
    prul_tp = ''
    prul_t = ''
    prul_p = ''
    pdul_tp = ''
    pdul_t = ''
    pdul_p = ''
    ppaead_t = ''
    ppa_t = ''
    ppa_p = ''
    pphead_t = ''
    pph_t = ''
    pph_p = ''
    atr42_t = ''
    atr42_cmv_t = ''
    schweizer300_t = ''
    ifr_cap_p = ''
    ifread_t = ''
    ifr_t = ''
    ifra61_p = ''
    ifrh61_c_p = ''
    ifrh61_d_p = ''

    # Mapeia todas as informações disponíveis no site para as variáveis adequadas
    for lineIndex in range(len(t)):
        line = t[lineIndex]
        if 'CNPJ:&nbsp;' in line:
            cnpj = line[(line.index('<b>') + 3):line.index('</B>')]
        elif 'Código ANAC:&nbsp;<b>' in line:
            canac = line[(line.index('<b>') + 3):line.index('</B>')]
        elif 'NomeFantasia:&nbsp;<B>' in line:
            nome_fantasia = line[(line.index('<B>') + 3):line.index('</B>')]
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
            uf = line[(line.index('<B>') + 3):line.index('</B>')]
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
                    cpa_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(cpa_p)
                except IndexError:
                    cpa_p = t[(lineIndex + 4)]

        elif 'COMISSÁRIO DE VOO - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    comsp_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(comsp_tp)
                except IndexError:
                    comsp_tp = t[(lineIndex + 4)]

        elif 'COMISSÁRIO DE VOO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    com_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(com_tp)
                except IndexError:
                    com_tp = t[(lineIndex + 4)]

        elif 'DESPACHANTE OPERACIONAL DE VOO (IS 141-003A)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    dovis141_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(dovis141_tp)
                except IndexError:
                    dovis141_tp = t[(lineIndex + 4)]

        elif 'DESPACHANTE OPERACIONAL DE VOO - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    dovsp_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(dovsp_tp)
                except IndexError:
                    dovsp_tp = t[(lineIndex + 4)]

        elif 'DESPACHANTE OPERACIONAL DE VOO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    dov_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(dov_tp)
                except IndexError:
                    dov_tp = t[(lineIndex + 4)]

        elif 'INSTRUTOR CERTIFICADO DE PILOTO AERODESPORTIVO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    icpa_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(icpa_t)
                except IndexError:
                    icpa_t = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO AVIÃO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    invaead_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(invaead_t)
                except IndexError:
                    invaead_t = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO AVIÃO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    inva_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(inva_t)
                except IndexError:
                    inva_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    inva_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(inva_p)
                except IndexError:
                    inva_p = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    invhead_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(invhead_t)
                except IndexError:
                    invhead_t = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex + 1)]:
                try:
                    invh_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(invh_t)
                except IndexError:
                    invh_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex + 1)]:
                try:
                    invh_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(invh_p)
                except IndexError:
                    invh_p = t[(lineIndex + 4)]

        elif 'INSTRUTOR DE VOO PLANADOR' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    invpl_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(invpl_tp)
                except IndexError:
                    invpl_tp = t[(lineIndex + 4)]

            elif 'Teórico' in t[(lineIndex+1)]:
                try:
                    invpl_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(invpl_t)
                except IndexError:
                    invpl_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    invpl_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(invpl_p)
                except IndexError:
                    invpl_p = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO CÉLULA - MODALIDADE SEMIPRESENCIAL (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmacsp141_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmacsp141_tp)
                except IndexError:
                    mmacsp141_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO CÉLULA - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmacsp_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmacsp_tp)
                except IndexError:
                    mmacsp_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Célula (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmac141_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmac141_tp)
                except IndexError:
                    mmac141_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Célula' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmac_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmac_tp)
                except IndexError:
                    mmac_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO MOTOPROPULSOR - MODALIDADE SEMIPRESENCIAL (IS 141-002B)'\
                in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmagmpsp141_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmagmpsp141_tp)
                except IndexError:
                    mmagmpsp141_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - GRUPO MOTOPROPULSOR - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmagmpsp_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmagmpsp_tp)
                except IndexError:
                    mmagmpsp_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Motopropulsor (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmagmp141_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmagmp141_tp)
                except IndexError:
                    mmagmp141_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Motopropulsor' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmagmp_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmagmp_tp)
                except IndexError:
                    mmagmp_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - MÓDULO AVIÔNICOS - MODALIDADE SEMIPRESENCIAL (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmaasp141_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmaasp141_tp)
                except IndexError:
                    mmaasp141_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - MÓDULO AVIÔNICOS - MODALIDADE SEMIPRESENCIAL' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmaasp_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmaasp_tp)
                except IndexError:
                    mmaasp_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Aviônicos (IS 141-002B)' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmaa141_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmaa141_tp)
                except IndexError:
                    mmaa141_tp = t[(lineIndex + 4)]

        elif 'MECÂNICO DE MANUTENÇÃO AERONÁUTICA - Grupo Aviônicos' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    mmaa_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(mmaa_tp)
                except IndexError:
                    mmaa_tp = t[(lineIndex + 4)]

        elif 'PILOTO AGRÍCOLA DE AVIÃO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    paga_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(paga_tp)
                except IndexError:
                    paga_tp = t[(lineIndex + 4)]

        elif 'PILOTO AGRÍCOLA DE HELICÓPTERO' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    pagh_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pagh_tp)
                except IndexError:
                    pagh_tp = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL (AVIÃO)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    pca_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pca_p)
                except IndexError:
                    pca_p = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL DE AVIÃO/IFR - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    pcaifread_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pcaifread_t)
                except IndexError:
                    pcaifread_t = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL DE HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    pchead_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pchead_t)
                except IndexError:
                    pchead_t = t[(lineIndex + 4)]

        elif 'PILOTO COMERCIAL DE HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    pch_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pch_t)
                except IndexError:
                    pch_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    pch_p = t[(lineIndex+4)].split()[2]
                    cv_dates.append(pch_p)
                except IndexError:
                    pch_p = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE AVIÃO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    plaead_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(plaead_t)
                except IndexError:
                    plaead_t = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE AVIÃO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    pla_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pla_t)
                except IndexError:
                    pla_t = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    plahead_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(plahead_t)
                except IndexError:
                    plahead_t = t[(lineIndex + 4)]

        elif 'PILOTO DE LINHA AÉREA DE HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    plah_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(plah_t)
                except IndexError:
                    plah_t = t[(lineIndex + 4)]

        elif 'PILOTO DE PLANADOR' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    ppl_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ppl_t)
                except IndexError:
                    ppl_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    ppl_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ppl_p)
                except IndexError:
                    ppl_p = t[(lineIndex + 4)]

        elif 'PILOTO DE RECREIO DE ULTRALEVE' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    prul_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(prul_tp)
                except IndexError:
                    prul_tp = t[(lineIndex + 4)]

            elif 'Teórico' in t[(lineIndex+1)]:
                try:
                    prul_t = t[(lineIndex+4)].split()[2]
                    cv_dates.append(prul_t)
                except IndexError:
                    prul_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    prul_p = t[(lineIndex+4)].split()[2]
                    cv_dates.append(prul_p)
                except IndexError:
                    prul_p = t[(lineIndex + 4)]

        elif 'PILOTO DESPORTIVO DE ULTRALEVE' in line:
            if 'Teórico/Prático' in t[(lineIndex+1)]:
                try:
                    pdul_tp = datetime.datetime.strptime(t[(lineIndex+4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pdul_tp)
                except IndexError:
                    pdul_tp = t[(lineIndex + 4)]

            elif 'Teórico' in t[(lineIndex+1)]:
                try:
                    pdul_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pdul_t)
                except IndexError:
                    pdul_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    pdul_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pdul_p)
                except IndexError:
                    pdul_p = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE AVIÃO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    ppaead_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ppaead_t)
                except IndexError:
                    ppaead_t = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE AVIÃO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    ppa_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ppa_t)
                except IndexError:
                    ppa_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    ppa_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ppa_p)
                except IndexError:
                    ppa_p = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE HELICÓPTERO - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    pphead_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pphead_t)
                except IndexError:
                    pphead_t = t[(lineIndex + 4)]

        elif 'PILOTO PRIVADO DE HELICÓPTERO' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    pph_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pph_t)
                except IndexError:
                    pph_t = t[(lineIndex + 4)]

            elif 'Prático' in t[(lineIndex+1)]:
                try:
                    pph_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(pph_p)
                except IndexError:
                    pph_p = t[(lineIndex + 4)]

        elif 'TREINAMENTO DE SOLO ATR-42 - CMV' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    atr42_cmv_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(atr42_cmv_t)
                except IndexError:
                    atr42_cmv_t = t[(lineIndex + 4)]

        elif 'TREINAMENTO DE SOLO ATR-42' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    atr42_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(atr42_cmv_t)
                except IndexError:
                    atr42_t = t[(lineIndex + 4)]

        elif 'TREINAMENTO DE SOLO SCHWEIZER 300' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    schweizer300_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(schweizer300_t)
                except IndexError:
                    schweizer300_t = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS (SOB CAPOTA)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    ifr_cap_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ifr_cap_p)
                except IndexError:
                    ifr_cap_p = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS - MODALIDADE EAD' in line:
            if 'Teórico' in t[(lineIndex+1)]:
                try:
                    ifread_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ifread_t)
                except IndexError:
                    ifread_t = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS AVIÃO (IS 61-002D)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    ifra61_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ifra61_p)
                except IndexError:
                    ifra61_p = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS HELICÓPTERO (IS 61-002C)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    ifrh61_c_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ifrh61_c_p)
                except IndexError:
                    ifrh61_c_p = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS HELICÓPTERO (IS 61-002D)' in line:
            if 'Prático' in t[(lineIndex+1)]:
                try:
                    ifrh61_d_p = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ifrh61_d_p)
                except IndexError:
                    ifrh61_d_p = t[(lineIndex + 4)]

        elif 'VOO POR INSTRUMENTOS' in line:
            if 'Teórico' in t[(lineIndex + 1)]:
                try:
                    ifr_t = datetime.datetime.strptime(t[(lineIndex + 4)].split()[2], '%d/%m/%Y')
                    cv_dates.append(ifr_t)
                except IndexError:
                    ifr_t = t[(lineIndex + 4)]

    # Verifica qual curso está mais próximo de vencer.
    i_dates = []
    for i in cv_dates:
        if type(i) == datetime.datetime:
            i_dates.append(i)

    try:
        pr_ca_v = min(i_dates)
    except ValueError:
        pr_ca_v = ''

    # Retorna uma string no formato especificado em genFirstLine, correspondente a cada CIAC

    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Informações mapeadas para o CANAC {canac} - {nome_fantasia}")

    return (canac,
            nome_fantasia,
            razao,
            cnpj,
            situacao,
            endereco,
            bairro,
            cidade,
            uf,
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
            pr_ca_v,
            cpa_p,
            comsp_tp,
            com_tp,
            dovis141_tp,
            dovsp_tp,
            dov_tp,
            icpa_t,
            invaead_t,
            inva_t,
            inva_p,
            invhead_t,
            invh_t,
            invh_p,
            invpl_tp,
            invpl_t,
            invpl_p,
            mmacsp141_tp,
            mmac141_tp,
            mmacsp_tp,
            mmac_tp,
            mmagmpsp141_tp,
            mmagmp141_tp,
            mmagmpsp_tp,
            mmagmp_tp,
            mmaasp141_tp,
            mmaa141_tp,
            mmaasp_tp,
            mmaa_tp,
            paga_tp,
            pagh_tp,
            pca_p,
            pcaifread_t,
            pchead_t,
            pch_t,
            pch_p,
            plaead_t,
            pla_t,
            plahead_t,
            plah_t,
            ppl_t,
            ppl_p,
            prul_tp,
            prul_t,
            prul_p,
            pdul_tp,
            pdul_t,
            pdul_p,
            ppaead_t,
            ppa_t,
            ppa_p,
            pphead_t,
            pph_t,
            pph_p,
            atr42_t,
            atr42_cmv_t,
            schweizer300_t,
            ifr_cap_p,
            ifread_t,
            ifr_t,
            ifra61_p,
            ifrh61_c_p,
            ifrh61_d_p)


if __name__ == '__main__':
    outputFile = f'{input("caminho e nome do arquivo sem extensão: ")}.csv'
    process_start = datetime.datetime.now()
    CANACs = get_canacs_from_index()

    print(f'[{datetime.datetime.now().strftime("%H:%M:%S")}] Buscando e transcrevendo informações para {outputFile}\n')
    with open(outputFile, 'w', newline='', encoding='utf-8-sig') as fileHandle:
        csvWriter = csv.writer(fileHandle, dialect='excel')
        csvWriter.writerow(gen_first_line())
        for CIACCANAC in CANACs:
            csvWriter.writerow(parse_info_by_canac(CIACCANAC))

    delta_t = str(datetime.datetime.now() - process_start).split(':')
    print(f"\n[{datetime.datetime.now().strftime('%H:%M:%S')}] processo terminado em "
          f"{f'{int(delta_t[1])} minuto' if int(delta_t[1]) > 0 else ''}"
          f"{'s' if int(delta_t[1]) > 1 else ''}"
          f"{' e ' if int(delta_t[1]) > 0 and round(float(delta_t[2]), 1) > 0 else ''}"
          f"{f'{round(float(delta_t[2]), 1)} segundo' if float(delta_t[2]) > 0 else ''}"
          f"{'s' if round(float(delta_t[2]), 1) > 1 else ''}\n")
