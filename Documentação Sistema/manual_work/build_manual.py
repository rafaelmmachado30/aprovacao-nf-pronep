#!/usr/bin/env python3
"""Gera o XML do conteudo do manual e injeta no document.xml do timbrado."""

# Helpers pra construir paragrafos com formatacao Pronep
def p_titulo_h1(texto):
    """Heading 1: Arial 12pt bold azul-escuro #1F4E79"""
    return f'''<w:p><w:pPr><w:spacing w:before="240" w:after="120"/><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">{texto}</w:t></w:r></w:p>'''

def p_titulo_h2(texto):
    """Heading 2: Arial 11pt bold vermelho-escuro #C00000"""
    return f'''<w:p><w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="C00000"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">{texto}</w:t></w:r></w:p>'''

def p_subtitulo(texto):
    """Subtitulo: Arial 11pt bold azul-claro #27AAE1"""
    return f'''<w:p><w:pPr><w:spacing w:before="160" w:after="60"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="27AAE1"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">{texto}</w:t></w:r></w:p>'''

def p_texto(texto, justify=True):
    """Paragrafo normal: Arial 10pt"""
    jus = '<w:jc w:val="both"/>' if justify else ''
    return f'''<w:p><w:pPr><w:spacing w:after="100" w:line="276" w:lineRule="auto"/>{jus}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">{texto}</w:t></w:r></w:p>'''

def p_texto_negrito(texto):
    """Paragrafo com mix de bold (usa marcadores **texto**)"""
    runs = ''
    parts = texto.split('**')
    for i, part in enumerate(parts):
        if not part:
            continue
        bold = '<w:b/>' if i % 2 == 1 else ''
        # escape XML chars
        safe = part.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        runs += f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>{bold}<w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">{safe}</w:t></w:r>'
    return f'<w:p><w:pPr><w:spacing w:after="100" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>{runs}</w:p>'

def p_passo(num, texto):
    """Item numerado de passo a passo"""
    runs = ''
    # primeiro: numero em destaque
    runs += f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="1F4E79"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">{num}.  </w:t></w:r>'
    # texto com bold parsing
    parts = texto.split('**')
    for i, part in enumerate(parts):
        if not part: continue
        bold = '<w:b/>' if i % 2 == 1 else ''
        safe = part.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        runs += f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>{bold}<w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">{safe}</w:t></w:r>'
    return f'<w:p><w:pPr><w:spacing w:after="80" w:line="276" w:lineRule="auto"/><w:ind w:left="360" w:hanging="360"/></w:pPr>{runs}</w:p>'

def p_dica(texto):
    """Box de dica destacado em azul"""
    parts = texto.split('**')
    runs = ''
    for i, part in enumerate(parts):
        if not part: continue
        bold = '<w:b/>' if i % 2 == 1 else ''
        safe = part.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        runs += f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>{bold}<w:color w:val="1F4E79"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">{safe}</w:t></w:r>'
    return f'''<w:p><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:space="6" w:color="27AAE1"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="EDF2F9"/><w:spacing w:before="120" w:after="120" w:line="260" w:lineRule="auto"/><w:ind w:left="240" w:right="240"/></w:pPr>{runs}</w:p>'''

def p_separador():
    """Linha de espaco"""
    return '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>'

# ============================================================
# MONTA O CONTEUDO DO MANUAL
# ============================================================
content = []

# Capa
content.append(p_titulo_h1('Sistema de Aprovação de Notas Fiscais'))
content.append(p_subtitulo('Manual de Instalação e Uso'))
content.append(p_texto('Este manual ensina, passo a passo, como instalar o Sistema de Aprovação de NF Pronep no seu computador (Windows ou Mac) e no seu celular (Android ou iPhone), e como ativar as notificações que avisam quando uma nota fiscal chega para você aprovar.', justify=True))
content.append(p_texto('O sistema é uma aplicação web moderna que pode ser instalada como um aplicativo nativo no seu dispositivo. Isso significa que você terá um ícone do Pronep direto na tela inicial / Menu Iniciar, e o app abre em tela cheia, sem a barra do navegador.', justify=True))
content.append(p_separador())

# Acesso pelo browser
content.append(p_titulo_h2('1. Acessando o sistema pelo navegador'))
content.append(p_texto_negrito('Endereço de acesso: **https://purple-forest-09588fe10.7.azurestaticapps.net**'))
content.append(p_texto_negrito('Use o **Microsoft Edge** ou o **Google Chrome** para a melhor experiência. Você fará login automaticamente com sua conta Microsoft @pronep.com.br — não precisa criar nova senha.'))
content.append(p_separador())

# Instalacao Windows
content.append(p_titulo_h2('2. Instalando como aplicativo no computador (Windows)'))
content.append(p_subtitulo('Opção A — Instalação pelo navegador (recomendada)'))
content.append(p_passo(1, 'Abra o sistema no **Microsoft Edge** ou **Google Chrome**'))
content.append(p_passo(2, 'Na barra de endereço, no canto direito, você verá um ícone de **monitor com uma seta para baixo** (ou um sinal de mais "+")'))
content.append(p_passo(3, 'Clique no ícone — vai aparecer um pop-up perguntando se você quer instalar o aplicativo'))
content.append(p_passo(4, 'Clique em **Instalar**'))
content.append(p_passo(5, 'Pronto. O aplicativo agora aparece no Menu Iniciar e pode ser fixado na barra de tarefas'))
content.append(p_separador())
content.append(p_subtitulo('Opção B — Instalador automático (pacote enviado pelo TI)'))
content.append(p_texto_negrito('Se você recebeu uma pasta com os arquivos **Instalar Sistema NF Pronep.bat**, **Sistema-NF-Pronep.bat** e **Pronep-NF.ico**:'))
content.append(p_passo(1, 'Clique duas vezes no arquivo **"Instalar Sistema NF Pronep.bat"**'))
content.append(p_passo(2, 'O Windows pode mostrar um aviso de segurança ("Windows protegeu seu computador"). Clique em **Mais informações** -> **Executar assim mesmo**'))
content.append(p_passo(3, 'Aguarde os 4 passos da instalação (leva 5 segundos)'))
content.append(p_passo(4, 'Um atalho **"Sistema NF Pronep"** aparece no seu Desktop e no Menu Iniciar, com o ícone Pronep'))
content.append(p_passo(5, 'O sistema abre automaticamente em modo aplicativo (sem barra do navegador)'))
content.append(p_separador())

# Instalacao Android
content.append(p_titulo_h2('3. Instalando como aplicativo no celular Android'))
content.append(p_passo(1, 'Abra o **Google Chrome** no celular'))
content.append(p_passo(2, 'Acesse **https://purple-forest-09588fe10.7.azurestaticapps.net**'))
content.append(p_passo(3, 'Faça login com sua conta Microsoft @pronep.com.br'))
content.append(p_passo(4, 'Toque nos **três pontinhos** no canto superior direito do Chrome'))
content.append(p_passo(5, 'Selecione **Instalar app** ou **Adicionar à tela inicial**'))
content.append(p_passo(6, 'Confirme. O ícone Pronep aparece na tela inicial do celular como um app normal'))
content.append(p_separador())

# Instalacao iPhone
content.append(p_titulo_h2('4. Instalando como aplicativo no iPhone'))
content.append(p_texto_negrito('**Importante**: no iPhone, use o **Safari** (não funciona pelo Chrome do iPhone).'))
content.append(p_passo(1, 'Abra o **Safari** no iPhone'))
content.append(p_passo(2, 'Acesse **https://purple-forest-09588fe10.7.azurestaticapps.net**'))
content.append(p_passo(3, 'Faça login com sua conta Microsoft @pronep.com.br'))
content.append(p_passo(4, 'Toque no ícone de **Compartilhar** (quadrado com uma seta para cima, na barra inferior)'))
content.append(p_passo(5, 'Role para baixo e toque em **Adicionar à Tela de Início**'))
content.append(p_passo(6, 'Confirme o nome (pode deixar "NF Pronep") e toque em **Adicionar**'))
content.append(p_passo(7, 'O ícone Pronep aparece na tela inicial. Abra por ele para ter a experiência completa de app'))
content.append(p_separador())

# Notificacoes
content.append(p_titulo_h2('5. Ativando notificações no celular e desktop'))
content.append(p_texto_negrito('Você pode receber **alertas no celular ou desktop** quando uma NF cair na sua fila de aprovação — funciona como notificação do WhatsApp, Gmail, etc.'))
content.append(p_texto_negrito('**Antes de ativar**: no iPhone, é obrigatório ter instalado o app pela Tela de Início (passo 4 acima). Em outros dispositivos, recomenda-se também ter instalado.'))
content.append(p_passo(1, 'Abra o sistema (pelo app instalado ou pelo navegador)'))
content.append(p_passo(2, 'No menu lateral, toque em **Configurações** (engrenagem no rodapé)'))
content.append(p_passo(3, 'Role a página até encontrar o card **"Notificações no celular / desktop"**'))
content.append(p_passo(4, 'Toque no botão **Ativar notificações**'))
content.append(p_passo(5, 'O navegador vai pedir permissão — toque em **Permitir**'))
content.append(p_passo(6, 'Quando o status mudar para **"Notificações ativadas neste dispositivo"**, você pode tocar em **Testar notificação** para confirmar'))
content.append(p_dica('**Dica**: você precisa ativar as notificações em cada dispositivo onde quiser receber (celular pessoal, celular do trabalho, desktop, etc). Cada dispositivo gera uma assinatura separada — pode ativar em todos sem problema.'))
content.append(p_separador())

# Visao geral do sistema
content.append(p_titulo_h2('6. Visão geral do sistema'))
content.append(p_subtitulo('Dashboard'))
content.append(p_texto('Tela inicial com indicadores do mês: valor total aprovado, NFs aprovadas, pendentes, duplicidades evitadas. Tem filtro por unidade (SP/RJ/ES/Todas) e gráficos por diretoria e evolução mensal.'))
content.append(p_subtitulo('Fila de Aprovação'))
content.append(p_texto('Lista de NFs pendentes da sua aprovação. Cada linha tem botões de aprovar (✓) ou rejeitar (×). NFs próximas do vencimento aparecem em vermelho. NFs acima do valor limite passam por aprovação em dois níveis.'))
content.append(p_subtitulo('Notas Aprovadas'))
content.append(p_texto('Histórico das NFs aprovadas. Tem filtros por unidade, diretoria, datas, status de processamento e busca. O setor financeiro pode marcar como "Processado" depois de integrar no sistema fiscal.'))
content.append(p_subtitulo('Notas Rejeitadas'))
content.append(p_texto('NFs que foram rejeitadas por algum aprovador, com o motivo. Quem lançou recebeu notificação por e-mail e Teams. A NF rejeitada pode ser corrigida e reenviada sem bloquear duplicidade.'))
content.append(p_subtitulo('Lançamento de NF'))
content.append(p_texto('Tela para subir uma nova NF: selecionar fornecedor, valor, vencimento, unidade e anexar o PDF. O sistema valida automaticamente CNPJ, duplicidade (hash do arquivo + CNPJ+número), prazo de vencimento (D+5 dias úteis) e roteia para o aprovador correto.'))
content.append(p_subtitulo('Fornecedores'))
content.append(p_texto('Cadastro de fornecedores com integração à Brasilapi para consulta automática de CNPJ. Permite importação em massa via planilha XLSX/CSV.'))
content.append(p_subtitulo('Configurações (admin)'))
content.append(p_texto('Configurações do sistema: aprovação multi-nível por valor, gestor master, integração com grupos do Microsoft Entra ID, ativação de notificações e (para administradores) limpeza da base.'))
content.append(p_separador())

# Suporte
content.append(p_titulo_h2('7. Dúvidas e suporte'))
content.append(p_texto_negrito('Em caso de dúvida, erro no sistema ou sugestão de melhoria, fale com **Rafael Machado** ou com o time de TI da Pronep.'))
content.append(p_texto_negrito('E-mail do sistema (origem das notificações): **datanalytics@pronep.com.br**'))
content.append(p_separador())

xml_content = '\n    '.join(content)

# Le o document.xml original
with open('unpacked/word/document.xml', 'r', encoding='utf-8') as f:
    doc = f.read()

# Substitui o paragrafo placeholder pelo nosso conteudo
placeholder = '<w:p w14:paraId="4DDBEF00" w14:textId="77777777" w:rsidR="007B111B" w:rsidRPr="00692E53" w:rsidRDefault="007B111B" w:rsidP="00692E53"/>\n    <w:sectPr'
replacement = xml_content + '\n    <w:sectPr'

if placeholder not in doc:
    print('ERRO: placeholder nao encontrado no document.xml')
    print('Procurando uma variante...')
    import re
    m = re.search(r'<w:p[^/]*paraId="4DDBEF00"[^/]*/>', doc)
    if m:
        print('Match alternativo encontrado:', m.group()[:80])
    raise SystemExit(1)

doc_new = doc.replace(placeholder, replacement)

with open('unpacked/word/document.xml', 'w', encoding='utf-8') as f:
    f.write(doc_new)

print('OK: document.xml atualizado')
print('Tamanho final:', len(doc_new), 'chars')
