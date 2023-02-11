import datetime
import time
import os.path
from selenium.webdriver.common.by import By
import os
from msedge.selenium_tools import Edge, EdgeOptions
import PyPDF2
import re
import xlsxwriter
import pathlib

if not os.path.isdir(os.path.join("Downloads")):
    os.mkdir(os.path.join('Downloads'))
full_path = str(pathlib.Path().resolve()) + '\Downloads'
processos = []
log = []


def delete_arq(arq):
    dir = os.listdir(full_path)
    for file in dir:
        if file == arq:
            os.remove(file)


def gera_log():
    for item in processos:
        if len(list(filter(lambda i: i[0] == item[0], log))) == 0:
            lista_d = list(filter(lambda i: (item[0] == i[0]), processos))
            if len(lista_d) > 1:
                removed = [x for n, x in enumerate(lista_d) if x not in lista_d[:n]]
                if len(removed) > 1:
                    for x in removed:
                        log.append([x[0], x[1]])
    workbook = xlsxwriter.Workbook(full_path + '/' + f'TST log_duplicatas.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'N° Processo')
    worksheet.write('B1', 'Data')
    linha = 2
    for item in log:
        worksheet.write('A' + str(linha), item[0])
        worksheet.write('B' + str(linha), item[1])
        linha += 1
    workbook.close()


def gera_planilha(dt_ini, dt_fim):
    controle = datetime.datetime.strptime(dt_ini, "%d/%m/%Y")
    while controle < datetime.datetime.strptime(dt_fim, "%d/%m/%Y"):
        lista_dia = list(filter(lambda i: i[1] == controle.strftime("%d/%m/%Y"), processos))
        if len(lista_dia) != 0:
            controle = controle + datetime.timedelta(days=1)
            workbook = xlsxwriter.Workbook(full_path + '/' + f'TST {controle.strftime("%d-%m-%Y")}.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.write('A1', 'N° Processo')
            worksheet.write('B1', 'Data')
            linha = 2
            for item in lista_dia:
                worksheet.write('A' + str(linha), item[0])
                worksheet.write('B' + str(linha), item[1])
                linha += 1
            workbook.close()
        else:
            controle = controle + datetime.timedelta(days=1)


def pdf_reader(arq, data):
    path = '/' + arq + '.pdf'
    pdf_file = open(full_path + path, 'rb')
    read_pdf = PyPDF2.PdfReader(pdf_file)
    numpages = len(read_pdf.pages)
    count_page = 0
    count_page_r = 0
    while count_page < numpages:
        for page in read_pdf.pages:
            rext = page.extract_text()
            res_search = re.search("PROCESSO Nº TST", rext)
            if res_search:
                ini = res_search.end() + 1
                end = rext.find("\n", ini)
                nprocess = rext[ini:end].replace(" ", "").replace("\n", "")
                data = data.replace("_", "/")

                processos.append([nprocess, data])
                count_page_r += 1
                print(f"Número de processos encontrados : {count_page_r} e número de páginas analisadas: {count_page} ")
            count_page += 1
            print(f"Número de processos encontrados : {count_page_r} e número de páginas analisadas: {count_page} ")
    pdf_file.close()


def latest_download_file():
    os.chdir(full_path)
    files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
    newest = files[-1]

    return newest


def check_download():
    fileends = "crdownload"
    while "crdownload" == fileends:
        time.sleep(1)
        newest_file = latest_download_file()
        if "crdownload" in newest_file:
            fileends = "crdownload"
        else:
            fileends = "none"
            return "Pronto"


def calcula_ultima_semana():
    data = datetime.datetime.today()
    data_ini = data
    data_fim = data
    while data_ini.weekday() != 6 and data_fim.weekday() != 5:
        if data.weekday() == 6:
            data_ini = data - datetime.timedelta(days=7)
            data_fim = data_ini + datetime.timedelta(days=6)
        data = data - datetime.timedelta(days=1)
    return data_ini.strftime("%d/%m/%Y"), data_fim.strftime("%d/%m/%Y")


def init():
    options = EdgeOptions()
    options.use_chromium = True
    options.add_argument("headless")
    options.add_argument("disable-gpu")
    options.add_experimental_option("prefs", {"download.default_directory": full_path})
    navegador = Edge(executable_path=r"msedgedriver.exe", options=options)
    navegador.get("https://dejt.jt.jus.br/dejt/f/n/diariocon")
    time.sleep(2)
    dias = calcula_ultima_semana()

    # setup de datas : INICIAL
    dataini = navegador.find_element(value="corpo:formulario:dataIni")
    dataini.clear()
    dataini.send_keys(dias[0])

    # setup de datas : FINAL
    datafim = navegador.find_element(value="corpo:formulario:dataFim")
    datafim.clear()
    datafim.send_keys(dias[1])

    # filtro : TST
    dropdown = navegador.find_element(value="corpo:formulario:tribunal")
    opcoes = dropdown.find_elements(By.TAG_NAME, "option")
    opcoes[1].click()

    # Click no botão de pesquisa
    time.sleep(1)
    navegador.find_element(value="corpo:formulario:botaoAcaoPesquisar").click()

    # obtendo lista de cadernos:
    table = navegador.find_element(By.XPATH, '//*[@id="diarioCon"]/fieldset/table')
    rows = table.find_elements(By.TAG_NAME, 'tr')

    nav = navegador.find_element(By.ID, "diarioInferiorNav")
    tds = nav.find_elements(By.TAG_NAME, 'td')
    if len(tds) > 2:
        total_cad = tds[3].text.split(" ")
    else:
        total_cad = tds[1].text.split(" ")
    total_cad = int(total_cad[4])
    count = 0
    while count != total_cad:
        for row in rows:
            if row.find_elements(By.TAG_NAME, 'td'):
                name = row.find_elements(By.TAG_NAME, 'td')[1].text.replace("/", "_")
                date = row.find_elements(By.TAG_NAME, 'td')[0].text.replace("/", "_")
                row.find_element(By.TAG_NAME, "button").click()
                check_download()
                arq = "Diario_" + name.split(" ")[1].split("_")[0] + "__" + str(int(date.split("_")[0])) + "_" + \
                      str(int(date.split("_")[1])) + "_" + date.split("_")[2]
                print(f"iniciando leitura do pdf: {name}")
                pdf_reader(arq, date)
                count += 1
                delete_arq(arq + '.pdf')

        # Paginação:
        if len(rows) >= 31:
            tdf = nav.find_element(By.TAG_NAME, "table")
            botoes = tdf.find_elements(By.TAG_NAME, 'button')
            for bot in botoes:
                try:
                    att = bot.get_attribute("onclick")
                    if str(count + 1) in att:
                        bot.click()
                        time.sleep(1)
                        table = navegador.find_element(By.XPATH, '//*[@id="diarioCon"]/fieldset/table')
                        rows = table.find_elements(By.TAG_NAME, 'tr')

                except:
                    continue
    gera_planilha(dias[0], dias[1])
    gera_log()
    navegador.close()


if __name__ == '__main__':
    init()
