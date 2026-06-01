/**
 * sanManual.js — Conhecimento detalhado do sistema embutido na SAN.
 *
 * Conteudo extraido de Manual_Usabilidade_NF_Pronep.pdf (Versao 1.0 · Maio 2026).
 * Cada secao corresponde a um perfil de usuario do sistema.
 *
 * A SAN consulta o trecho relevante baseado na view atual (via buildSystemPrompt).
 * Isso transforma a SAN num especialista do sistema — ela responde com a mesma
 * profundidade do manual oficial, mantendo coerencia com a documentacao escrita.
 */

const MANUAL_SUBMETEDOR = `
=== MANUAL DO SUBMETEDOR (CONHECIMENTO COMPLETO) ===

Quem eh o Submetedor: colaborador que recebe a NF do fornecedor (por e-mail, sistema ou fisica) e a lanca no sistema para iniciar o fluxo de aprovacao. Qualquer colaborador Pronep pode atuar como submetedor.

LANCAMENTO DE NF — passo a passo:
Acesse o menu lateral -> Lancamento de NF. O formulario pede:

1. FORNECEDOR: comece a digitar o nome ou CNPJ. O sistema autocompleta. Quando voce escolhe, o sistema preenche automaticamente Unidade (se atende uma so), Diretoria e Categoria (se nao for multi-diretoria). Se o fornecedor nao aparecer na busca, ele NAO ESTA CADASTRADO — peca ao Financeiro pra cadastrar antes (sem cadastro, o sistema nao permite o lancamento).

2. NUMERO DA NF: como impresso no documento (ex: "36534", "NF-2026-045"). Esse numero combinado com o fornecedor sera usado pra detectar duplicidade.

3. CHAVE DE ACESSO: 44 digitos da NF eletronica (codigo de barras). Obrigatoria para NFs eletronicas. O sistema usa esse codigo como chave primaria de duplicidade — se essa NF ja foi lancada antes, o sistema BLOQUEIA o envio.

4. SERIE: como impressa na nota.

5. VALOR: em reais. Use virgula como separador decimal (ex: 1.234,56). O sistema aplica mascara automatica — basta digitar e ele formata.

6. VENCIMENTO: data limite para pagamento. Formato dd/mm/aaaa. REGRA IMPORTANTE: deve ser maior ou igual a hoje + 5 dias uteis. Se for menor (vencimento proximo), o sistema EXIGE uma checkbox de confirmacao "ja alinhei previamente com o gestor financeiro" — sem marcar, NAO eh possivel enviar.

7. UNIDADE: SP, RJ ou ES (radio buttons). Se o fornecedor atende apenas UMA unidade, o sistema bloqueia automaticamente nessa. Se atende mais de uma (fornecedor marcado como "atende as 3 unidades"), voce escolhe qual unidade EMITIU a NF.

8. CATEGORIA E DIRETORIA: preenchidos automaticamente do cadastro do fornecedor. EXCECAO: se o fornecedor for marcado como "atende mais de uma diretoria", aparecem dropdowns pra voce escolher na hora qual categoria/diretoria essa NF especifica pertence.

9. DESCRICAO: breve resumo do servico/produto. Texto livre.

10. SOLICITANTE: quem demandou o servico (texto livre — pode ser o nome de quem pediu o servico, da Pronep).

11. ANEXO PDF: arraste e solte o arquivo da NF (ou clique pra escolher). Maximo 6 MB. So aceita PDF.

Apos preencher tudo, clique em Enviar para Aprovacao. O sistema:
- DETECTA DUPLICIDADE automaticamente (3 mecanismos: chave de acesso 44 digitos, hash do PDF, e combinacao CNPJ + numero). Se ja existir uma NF identica e ela NAO foi rejeitada, o sistema bloqueia. NFs rejeitadas podem ser RELANCADAS com mesmo numero.
- ROTEIA a NF para o aprovador correto baseado em Diretoria + Unidade do fornecedor.
- ENVIA notificacao imediata (e-mail + Teams + push) ao gestor responsavel.

ERROS COMUNS NO LANCAMENTO E COMO RESOLVER:
- "Fornecedor incompleto, cadastro pendente": o fornecedor existe mas falta CNPJ/CPF. Va em Fornecedores, edite e complete o cadastro.
- "Vencimento fora do prazo, marque o checkbox": vencimento ate 5 dias uteis exige confirmacao previa do gestor.
- "NF duplicada (chave de acesso ja lancada)": essa NF ja existe no sistema. Verifique em Minhas NFs.
- "PDF muito grande": comprima o PDF (existem ferramentas online ou no Adobe). Maximo 6 MB.
- "Anexei o PDF errado": antes de enviar, arraste o arquivo correto sobre o campo (sobrescreve). Apos envio, peca ao gestor rejeitar e relance.

ACOMPANHANDO SUAS NFs:
- Minhas NFs: lista todas as NFs que voce lancou (qualquer status).
- Fila de Aprovacao: mostra apenas as suas que estao pendentes — util pra ver se o gestor ja aprovou.

QUANDO UMA NF EH REJEITADA:
Voce recebe e-mail com motivo e notificacao Teams. Acesse Notas Rejeitadas no menu para detalhes. Corrija o que foi apontado e relance — o sistema permite reenvio com mesmo numero desde que a anterior esteja com status Rejeitada.

FAQ SUBMETEDOR:
- Fornecedor nao cadastrado: peca ao Financeiro cadastrar antes.
- Lancar a mesma NF duas vezes: nao eh permitido — sistema bloqueia por chave de acesso, hash do PDF ou CNPJ+numero.
- Posso editar uma NF apos envio: nao — peca ao gestor rejeitar e relance corrigida.
`;

const MANUAL_APROVADOR = `
=== MANUAL DO APROVADOR (CONHECIMENTO COMPLETO) ===

Quem eh o Aprovador: gestor responsavel por uma ou mais combinacoes de Unidade + Diretoria. Quando uma NF eh lancada no seu escopo, voce recebe notificacao imediata e tem 4 caminhos pra aprovar ou rejeitar.

NOTIFICACOES AUTOMATICAS:
- E-mail com dados da NF e link direto para o PDF.
- Notificacao privada no Microsoft Teams (chat 1:1 com o sistema).
- Push notification no celular (se PWA instalado).

4 CAMINHOS PARA APROVAR/REJEITAR:

CAMINHO 1 — Pela tela (Fila de Aprovacao):
Menu lateral -> Fila de Aprovacao. Voce ve apenas NFs do seu escopo. Clique pra abrir detalhe (PDF, valor, vencimento). Acoes:
- Aprovar: registra, aplica watermark no PDF, arquiva em Notas Aprovadas, notifica submetedor.
- Rejeitar: pede motivo (obrigatorio), arquiva em Notas Rejeitadas, notifica submetedor.

CAMINHO 2 — Pelo e-mail:
Cada e-mail tem 2 botoes grandes: Aprovar e Rejeitar. Clique direto sem abrir o sistema — ideal pra aprovacoes rapidas no celular.

CAMINHO 3 — Pelo Teams:
Notificacao Teams chega como cartao interativo com dados + 2 botoes. Aprovar/Rejeitar direto do Teams eh valido.

CAMINHO 4 — Pela SAN (chat e voz):
Abra a Fila de Aprovacao, clique no FAB da SAN. Digite ou fale:
- "SAN, aprove a NF 12345"
- "SAN, rejeite a NF 12345 porque o valor esta errado"
A SAN busca, mostra os dados e abre modal de confirmacao. Voce clica Sim, aprovar. Comando de voz eh especialmente util no celular.

ALERTAS AUTOMATICOS DA SAN (e-mail):
- 9h dias uteis: relatorio matinal das NFs pendentes no seu escopo, ordenadas por vencimento.
- 17h seg-qui: atualizacao da tarde.
- 17h sexta: resumo semanal consolidado.
Inclui paragrafo de insight da IA destacando D+5, valores atipicos, concentracao em fornecedor. Se nao ha pendencias, e-mail nao eh enviado (zero spam).

DASHBOARD:
No menu, indicadores do seu escopo: total pendente, aprovado no mes, NFs em D+5, top fornecedores. Reflete apenas NFs onde voce eh o aprovador responsavel.

D+5 — REGRA DE PRIORIDADE:
NFs com vencimento em ate 5 dias UTEIS sao destacadas (linha amarela na fila). Sao prioridade alta — priorize a aprovacao pra evitar atraso de pagamento ao fornecedor.

APROVACAO MULTI-NIVEL:
NFs acima de valor configurado pelo Admin exigem dois niveis de aprovacao. Voce aprova como Nivel 1 e a NF segue automaticamente pro Nivel 2 (gestor financeiro/diretor). Voce eh notificado quando a NF eh finalmente aprovada.

FAQ APROVADOR:
- Vou estar de ferias: peca ao Admin ajustar temporariamente o mapeamento de Diretoria/Unidade no SharePoint pra outro gestor.
- Cliquei em Aprovar no e-mail e abriu a tela: apenas no primeiro acesso do dia (autenticacao). Apos isso, os proximos cliques aprovam direto.
- O que eh D+5: vencimento em ate 5 dias uteis a partir de hoje. Linha amarela na fila = prioridade.
- Rejeitar exige motivo: sim, sempre. O motivo vai pro e-mail/Teams do submetedor pra ele corrigir.
`;

const MANUAL_FINANCEIRO = `
=== MANUAL DO FINANCEIRO (CONHECIMENTO COMPLETO) ===

Quem eh o Financeiro: responsavel pela conferencia das NFs aprovadas, marcacao de pagamentos e gestao completa da base de fornecedores. Unico perfil (junto com Admin) que pode cadastrar, editar ou inativar fornecedores.

NOTAS APROVADAS — CONFERENCIA E PAGAMENTO:
Menu -> Notas Aprovadas. Lista todas NFs aprovadas (todas unidades, todas diretorias). Para cada NF:
- Ver PDF: documento original com watermark de aprovacao.
- Marcar como Pago: checkbox que registra confirmacao de pagamento (persiste no SharePoint).
- Filtros disponiveis: periodo de aprovacao, periodo de vencimento, unidade, diretoria, status de pagamento (Pago Sim/Nao).
Visual: linha em amarelo ate 5 dias uteis antes do vencimento; vermelha apos vencimento.

NOTAS REJEITADAS:
Acompanhe rejeicoes recentes pra entender padroes (fornecedor problematico, motivos recorrentes). Util pra feedback aos submetedores.

GESTAO DE FORNECEDORES — Menu Fornecedores:

CADASTRAR NOVO FORNECEDOR:
1. Tipo: CNPJ (PJ) ou CPF (PF).
2. Documento: o sistema valida digito verificador e busca dados publicos automaticamente via BrasilAPI (quando CNPJ).
3. Razao Social / Nome Completo: ate 200 caracteres.
4. Categoria: Servicos, Materiais, Locacao, Reembolso, Medicamentos ou Outros. RESTRICAO: categoria "Outros" so permite CPF (PF). Pra CNPJ, use uma das outras.
5. UF de Atendimento e Diretoria Padrao: definem o roteamento.
6. Atende as 3 unidades: marque se atende SP, RJ e ES.
7. Atende mais de uma Diretoria: marque se o fornecedor presta servicos variados. Quando NF for lancada, o submetedor escolhera categoria/diretoria na hora.

EDITAR FORNECEDOR:
Clique em Editar. Pode alterar qualquer campo EXCETO o documento (CNPJ/CPF eh a chave unica).

INATIVAR FORNECEDOR:
Desmarque a checkbox "Fornecedor ativo". Bloqueia novos lancamentos mas preserva historico. Pra reativar, marque novamente. NUNCA exclua — sempre inative.

IMPORTACAO EM MASSA:
Pra cadastrar muitos fornecedores de uma vez, use Importar. Baixe o template CSV/XLSX, preencha, suba. O sistema valida CNPJs, evita duplicatas e mostra preview antes de gravar.

USO DA SAN — comandos uteis pro Financeiro:
- "Resumo do que foi aprovado este mes"
- "Top 10 fornecedores em volume"
- "NFs ainda nao pagas que vencem essa semana"
- "Compare o gasto deste mes com o mes passado"
A SAN gera relatorios formatados — clique em Exportar PDF na resposta dela pra gerar documento com papel timbrado Pronep pronto pra distribuir.

FAQ FINANCEIRO:
- Marquei como Pago por engano: clique novamente na checkbox para desmarcar.
- Posso aprovar NFs: NAO — Financeiro confere e marca pagamento. Aprovacao eh responsabilidade exclusiva dos Aprovadores (Gestores) por diretoria.
- Cadastrar fornecedor multi-diretoria: marque "Atende mais de uma Diretoria/Categoria" no cadastro. Quando NF for lancada, o submetedor escolhe na hora.
- Cadastro de fornecedor novo: solicite Razao Social, CNPJ, Nome Fantasia, UF principal, dados bancarios. Sistema preenche resto via BrasilAPI.
`;

const MANUAL_GERAL = `
=== INFORMACOES GERAIS DO SISTEMA ===

ACESSO:
- URL principal: o usuario abre pelo navegador (Chrome, Edge, Safari, Firefox) ou pelo Teams (aba pessoal) ou pelo PWA instalado (icone na barra do navegador permite instalar como app).
- Autenticacao: automatica via Entra ID (Microsoft 365 da Pronep). Nao precisa criar conta — o sistema reconhece pelo seu e-mail Pronep.

MATRIZ DE PERMISSOES (resumo):
- Submetedor: lanca NFs proprias, ve Minhas NFs, ve sua Fila de Aprovacao (proprias pendentes), recebe rejeicoes.
- Aprovador (Gestor): tudo do Submetedor + ve Fila de Aprovacao do seu escopo + aprova/rejeita + Dashboard + Notas Aprovadas do seu escopo.
- Financeiro: ve Notas Aprovadas (TUDO), marca Pago/Processado, gestao completa de Fornecedores, NUNCA aprova/rejeita.
- Admin: tudo + Configuracoes + Limpar Base.

TELAS DO SISTEMA:
- Dashboard: indicadores e graficos consolidados.
- Fila de Aprovacao: NFs pendentes de aprovacao (escopo varia por perfil).
- Minhas NFs: NFs que o usuario logado lancou.
- Notas Aprovadas: NFs ja aprovadas (visiveis pra Aprovador no seu escopo, Financeiro ve tudo).
- Notas Rejeitadas: NFs que foram rejeitadas.
- Lancamento de NF: form de criacao de nova NF.
- Fornecedores: cadastro e gestao da base (so Financeiro e Admin).
- Mapa de Aprovadores: visualiza quem aprova o que (Diretoria + Unidade -> Gestor).
- Configuracoes: ajustes do sistema (so Admin).

CONCEITOS-CHAVE:
- Escopo: combinacao de Diretoria + Unidade que define o que cada Aprovador ve. Cada Gestor pode ter 1 ou N escopos.
- D+5: vencimento em ate 5 DIAS UTEIS a partir de hoje. Eh prioridade. Sistema destaca em amarelo na fila, depois vermelho apos vencer.
- Duplicidade: 3 mecanismos protegem — chave de acesso (44 digitos da NF-e), hash do PDF, combinacao CNPJ+Numero. NFs rejeitadas podem ser relancadas com mesmo numero.
- Multi-nivel: NFs acima de valor configurado (admin define) exigem 2 aprovacoes em cascata. Nivel 1 (gestor da diretoria) -> Nivel 2 (gestor financeiro/diretor).
- Multi-diretoria: fornecedor pode atender mais de uma diretoria. Quando NF lancada, submetedor escolhe categoria/diretoria na hora.
- Multi-unidade: fornecedor pode atender SP, RJ e ES. Submetedor escolhe qual unidade emitiu a NF.

CONTATO:
- Suporte tecnico: rafael.machado@pronep.com.br
- Fornecedor nao cadastrado: contate o Financeiro.
- Aprovacao travada: contate o Admin (mapeamento de Diretoria/Unidade).
`;

/**
 * Retorna o trecho do manual relevante pra view atual.
 * Pra views que sao do dia-a-dia de um perfil especifico, retorna o conteudo completo
 * daquele perfil. Pra views genericas, retorna apenas info geral.
 */
function getManualForView(viewAtual) {
  switch (viewAtual) {
    case 'lancamento':
      return MANUAL_GERAL + '\n' + MANUAL_SUBMETEDOR;
    case 'fila-aprovacao':
      return MANUAL_GERAL + '\n' + MANUAL_APROVADOR;
    case 'aprovadas':
      return MANUAL_GERAL + '\n' + MANUAL_FINANCEIRO;
    case 'fornecedores':
      return MANUAL_GERAL + '\n' + MANUAL_FINANCEIRO;
    case 'rejeitadas':
    case 'rejeitadas-minhas':
      return MANUAL_GERAL + '\n' + MANUAL_SUBMETEDOR;
    case 'minhas-nfs':
      return MANUAL_GERAL + '\n' + MANUAL_SUBMETEDOR;
    case 'dashboard':
    case 'configuracoes':
    case 'mapa-aprovadores':
    default:
      return MANUAL_GERAL;
  }
}

module.exports = { getManualForView, MANUAL_GERAL, MANUAL_SUBMETEDOR, MANUAL_APROVADOR, MANUAL_FINANCEIRO };
