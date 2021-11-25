# import do pacote que fará a interface gráfica
from PySimpleGUI import PySimpleGUI as sg
import sys
import pandas as pd

# Layout
# definindo o tema que será usado em nosso Layout
sg.theme('Reddit')

if len(sys.argv) == 1:
    fname = sg.Window('Escolher tabela',
                      [[sg.Text('Localizar tabela')],
                       [sg.In(), sg.FileBrowse()],
                       [sg.Open(), sg.Cancel()]]).read(close=True)[1][0]

else:
    fname = sys.argv[1]

planilha = pd.read_excel(rf"{fname}")
#planilha = pd.read_excel(f"Pasta1.xlsx")

qtde_linhas = planilha.shape[0]

lista = []

for i in range(0, qtde_linhas):
    evento = []
    evento.append(planilha["Índice"][i])
    evento.append(planilha["Horário Divulgação"][i])
    evento.append(planilha["Nome"][i])
    evento.append(planilha["Cargo"][i])
    evento.append(planilha["Horário"][i])
    evento.append(planilha["Telefone"][i])
    evento.append(planilha["Skype"][i])
    evento.append(planilha["Tel Assessoria"][i])
    lista.append(evento)

conteudo = ''

for i in lista:
    dentro = ''
    indice = str(i[0])
    horario_divulgacao = str(i[1])
    nome = str(i[2])
    cargo = str(i[3])
    horario = str(i[4])
    telefone = str(i[5])
    skype = str(i[6])
    tel_assessoria = str(i[7])

    if skype != 'nan':
        dentro += f"""<span face="Arial, sans-serif" color="#5b5c5f" style="margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; line-height: 16px; vertical-align: baseline;"><strong>Skype:&nbsp;</strong>{skype}</span><br></span><strong style="color: #73bfe8; font-family: Arial, sans-serif; font-size: 12px;">As entrevistas via Skype devem ser agendadas com a assessoria de imprensa pelo n&uacute;mero {tel_assessoria}.</strong>"""

    conteudo += f"""<table width="600" height="228" border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="#ffffff">
<tbody>
<tr>
<td height="10" style="width: 97.6875px;"></td>
<td colspan="2" style="width: 396.594px;"><span face="Arial, sans-serif" color="#002D4D" style="text-align: left; line-height: 20px; font-size: 18px; color: #002d4d; font-family: Arial, sans-serif;"><strong style="color: #002d4d; font-family: Arial, sans-serif; font-size: 18px;">{indice}</strong></span></td>
<td style="width: 97.7188px;"></td>
</tr>
<tr>
<td height="10" style="width: 97.6875px;"></td>
<td height="10" style="width: 29.75px;"></td>
<td height="10" style="width: 364.844px;"></td>
<td height="10" style="width: 97.7188px;"></td>
</tr>
<tr>
<td height="10" style="width: 97.6875px;"></td>
<td style="width: 29.75px;"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/aviso_de_pauta/imagens/icone_relogiov2.png" width="auto" height="auto" alt="Hor&aacute;rio" style="display: block; border: 0;"></td>
<td style="width: 364.844px;"><span face="Arial, sans-serif" color="#5b5c5f" style="text-align: left; line-height: 24px; font-size: 15px; color: #5b5c5f; font-family: Arial, sans-serif;">Divulga&ccedil;&atilde;o &agrave;s <strong>{horario_divulgacao}h</strong>&nbsp;no Portal IBRE.</span></td>
<td style="width: 97.7188px;"></td>
</tr>
<tr>
<td height="10" style="width: 97.6875px;"></td>
<td height="10" style="width: 29.75px;"></td>
<td height="10" style="width: 364.844px;"></td>
<td height="10" style="width: 97.7188px;"></td>
</tr>
<tr>
<td height="10" style="width: 97.6875px;"></td>
<td style="width: 29.75px;"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/aviso_de_pauta/imagens/icone_imprensav2.png" width="auto" height="auto" alt="Imprensa" style="display: block; border: 0; vertical-align: top;"></td>
<td style="width: 364.844px;"><span face="Arial, sans-serif" color="#5b5c5f" style="text-align: left; line-height: 20px; font-size: 15px; color: #5b5c5f; font-family: Arial, sans-serif;">Atendimento &agrave; Imprensa:</span></td>
<td style="width: 97.7188px;"></td>
</tr>
<tr>
<td height="10" style="width: 97.6875px;"></td>
<td style="width: 29.75px;"></td>
<td style="width: 364.844px;">
<p>
    <span face="Arial, sans-serif" color="#5b5c5f" style="text-align: left; line-height: 16px; font-size: 12px; color: #5b5c5f; font-family: Arial, sans-serif;"><strong>{nome}</strong>, {cargo}</span> 
    <br> 
    <span face="Arial, sans-serif" color="#5b5c5f" style="text-align: left; line-height: 16px; font-size: 12px; color: #5b5c5f; font-family: Arial, sans-serif;"><strong>Hor&aacute;rio:</strong>&nbsp;a partir das {horario}h</span>
    <br>
    <span face="Arial, sans-serif" color="#5b5c5f" style="margin: 0px; padding: 0px; border: 0px; font-variant-numeric: inherit; font-variant-east-asian: inherit; font-stretch: inherit; font-size: 12px; line-height: 16px; font-family: Arial, sans-serif; vertical-align: baseline; color: #5b5c5f;"><span face="Arial, sans-serif" color="#5b5c5f" style="margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; line-height: 16px; vertical-align: baseline;"><strong>Telefone:&nbsp;</strong>{telefone}</span>
    <br style="color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 14px;">

    {dentro}
</p>
</td>
<td style="width: 95.7812px;"></td>
</tr>
<tr>
<td height="30" style="width: 108px;"></td>
<td height="30" style="width: 29px;"></td>
<td height="30" style="width: 398.219px;"></td>
</tr>
</tbody>
</table> """

# criando um layou com 4 linhas
layout = [
    [sg.Text('Mês'), sg.Input(key='mes', size=(20, 1))],
    [sg.Text('Cabeçalho'), sg.Input(key='cabecalho', size=(20, 10))],
    [sg.Button('Converter')],
    [sg.Text('Saída:'), sg.Text(size=(20, 4), key='-OUTPUT-')],
]

# Janela
janela = sg.Window('Conversor HTML', layout)

# Ler os Eventos
while True:
    eventos, valores = janela.read()
    if eventos == sg.WINDOW_CLOSED:
        break
    if eventos == 'Converter':
        mes = str(valores['mes'])
        cabecalho = str(valores['cabecalho'])

        resposta = f"""<!DOCTYPE html>
<html>
<head>
<title>FGV IBRE - Aviso de Pauta</title>

</head>
<body bgcolor="#F4F4F4" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table id="Barra IBRE" width="600" border="0" cellpadding="0" cellspacing="0" align="center">
<tbody>
<tr>
<td width="auto" height="auto" bgcolor="#ffffff"><a href="https://ibre.fgv.br/" target="_blank" rel="noopener"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/barra_ibre_01.jpg" alt="IBRE" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://www.linkedin.com/company/fgv-ibre/" target="_blank" rel="noopener"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/barra_ibre_02.jpg" alt="Linkedin" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://www.facebook.com/FGV.IBRE/" target="_blank" rel="noopener"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/barra_ibre_03.jpg" alt="Facebook" style="display: block; border: 0px;"></a></td>
<td width="auto" height="auto" bgcolor="#ffffff"><a href="https://twitter.com/FGVIBRE" target="_blank" rel="noopener"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/barra_ibre_04.jpg" alt="Twitter" style="display: block; border: 0px;"></a></td>
<td width="auto" height="auto" bgcolor="#ffffff"><a href="https://www.youtube.com/playlist?list=PLspVbtJ_9_Hp7nMkkHO9ZmEoLzcUMn_lB" target="_blank" rel="noopener"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/barra_ibre_05.jpg" alt="Youtube" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/barra_ibre_06.jpg" alt="FGV IBRE" style="display: block; border: 0px;"></td>
</tr>
</tbody>
</table>
<table width="600" height="auto" border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="#ffffff">
<tbody>
<tr>
<td><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/aviso_de_pauta/imagens/header.png" width="600" height="auto" alt="FGV IBRE - Aviso de Pauta" style="display: block; border: 0;"></td>
</tr>
</tbody>
</table>
<table width="600" height="auto" border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="#ffffff">
<tbody>
<tr>
<td width="50" height="20"></td>
<td width="500" height="20"></td>
<td width="50" height="20"></td>
</tr>
<tr>
<td width="50" height="20"></td>
<td width="500" height="20" style="font-family: arial, sans-serif; font-size: 11px; color: #002d4d; font-weight: bold; text-align: left;">{mes}</td>
<td width="50" height="20"></td>
</tr>
<tr>
<td width="50" height="10"></td>
<td width="500" height="10"></td>
<td width="50" height="10"></td>
</tr>
<tr>
<td width="50"></td>
<td width="500"><span face="Arial, sans-serif" color="#5b5c5f" style="text-align: left; line-height: 21px; font-size: 15px; color: #5b5c5f; font-family: Arial, sans-serif;">{cabecalho}</span></td>
<td width="50"></td>
</tr>
<tr>
<td width="50" height="30"></td>
<td width="500" height="30"></td>
<td width="50" height="30"></td>
</tr>
</tbody>
</table>   

{conteudo}

<table width="600" height="auto" border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="#ffffff">
<tbody>
<tr>
<td width="50" height="20"></td>
<td width="500" height="20"></td>
<td width="50" height="20"></td>
</tr>
<tr>
<td width="50"></td>
<td width="500" style="text-align: center;"><span style="font-size: 8pt;"><span style="background-color: #ffffff; font-family: arial, helvetica, sans-serif;"><span style="box-sizing: border-box;">Para mais informa&ccedil;&otilde;es, entre em contato com a equipe da Insight Comunica&ccedil;&atilde;o pelos telefones (21) 2509-5399 e (21) 99121-3771 ou pelo e-mail:<a href="mailto:assessoria.fgv@insightnet.com.br" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" data-linkindex="3" style="box-sizing: border-box; color: #3699ff; text-decoration-line: none; transition: color 0.15s ease 0s, background-color 0.15s ease 0s, border-color 0.15s ease 0s, box-shadow 0.15s ease 0s; margin: 0px; padding: 0px; border: 0px; font-style: inherit; font-variant: inherit; font-weight: inherit; font-stretch: inherit; line-height: inherit; vertical-align: baseline; outline: 0px !important; background-color: #ffffff;">assessoria.fgv@<wbr style="box-sizing: border-box;">insightnet.com.br</a>.</span></span><strong><em><span style="background-color: #ffffff; font-family: arial, helvetica, sans-serif;"><span style="box-sizing: border-box; text-align: start; margin: 0px; padding: 0px; border: 0px; font-variant-numeric: inherit; font-variant-east-asian: inherit; font-stretch: inherit; line-height: inherit; vertical-align: baseline; color: #201f1e; background-color: #ffffff;">&nbsp;</span></span></em></strong></span></td>
<td width="50"></td>
</tr>
</tbody>
</table>
<table width="600" height="auto" border="0" cellpadding="0" cellspacing="0" align="center" bgcolor="#ffffff">
<tbody>
<tr>
<td height="20"></td>
</tr>
<tr>
<td><a href="https://portalibre.fgv.br/" target="_blank" rel="noopener"><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/footer_ibre_01.jpg" width="auto" height="auto" alt="FGV IBRE - Blog da Conjuntura Econ&ocirc;mica" style="display: block; border: 0;"></a></td>
</tr>
<tr>
<td><img src="http://www.fgv.br/mailing/2021/ibre/novos_templates/icomex/imagens/footer_ibre_02.png" width="auto" height="auto" alt="FGV IBRE - Blog da Conjuntura Econ&ocirc;mica" style="display: block; border: 0;"></td>
</tr>
</tbody>
</table>
<!-- BARRA DICOM -->
<table width="600" height="53" border="0" cellpadding="0" cellspacing="0" align="center">
<tbody>
<tr>
<td valign="top">
<table width="600" height="14" border="0" align="center" cellpadding="0" cellspacing="0" id="Tabela_15">
<tbody>
<tr>
<td colspan="4"><img src="https://www.fgv.br/mailing/2021/dicom/43719_barra_dicom_2021/600/pt/imagens/barra_DICOM_2020_600px_1.png" width="600" height="auto" alt="linha" style="display: block;" border="0"></td>
</tr>
</tbody>
</table>
<table id="Tabela_16" width="600" height="53" border="0" cellpadding="0" cellspacing="0" align="center">
<tbody>
<tr>
<td width="auto" height="auto" bgcolor="#ffffff"><a href="https://portal.fgv.br/" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_01.jpg" width="390" height="53" alt="DICOM" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://www.linkedin.com/company/fgv" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_02.jpg" width="24" height="53" alt="LinkedIn" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://www.facebook.com/FGV/" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_03.jpg" width="24" height="53" alt="LinkedIn" style="display: block; border: 0px;"></a></td>
<td width="auto" height="auto" bgcolor="#ffffff"><a href="https://www.instagram.com/FGV.OFICIAL/" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_04.jpg" width="24" height="53" alt="Facebook" style="display: block; border: 0px;"></a></td>
<td width="auto" height="auto" bgcolor="#ffffff"><a href="https://twitter.com/fgv" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_05.jpg" width="24" height="53" alt="Instagram" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://www.tiktok.com/@fgv.oficial" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_06.jpg" width="24" height="52" alt="Twitter" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://www.youtube.com/fgv" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_07.jpg" width="24" height="52" alt="YouTube" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://portal.fgv.br/redes-sociais" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_08.jpg" width="17" height="53" alt="Mais" style="display: block; border: 0px;"></a></td>
<td bgcolor="#ffffff"><a href="https://portal.fgv.br/redes-sociais" target="_blank" rel="noopener"><img src="https://www.fgv.br/mailing/_BARRA_DICOM/2021/43719_barra_dicom_2021/600/pt/imagens/barra_dicom_600px_pt_09.jpg" width="50" height="53" alt="Mais" style="display: block; border: 0px;"></a></td>
</tr>
</tbody>
</table>
</td>
</tr>
<!-- BARRA DICOM --></tbody>
</table>
<!-- End Save for Web Slices -->
</body>
</html>
"""
    print(resposta)
    # limpando o arquivo
    open('montado.txt', 'w').close()

    # criando o arquivo, que receberá append(adições)
    with open('montado.txt', 'a', encoding="utf-8") as arquivo:
        print(resposta, file=arquivo)

    valores['-IN-'] = fname + 'montado.txt'
    janela['-OUTPUT-'].update(valores['-IN-'])