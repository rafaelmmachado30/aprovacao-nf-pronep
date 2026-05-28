# Setup Push Notifications — passo a passo

**Objetivo:** habilitar push notifications nativas no celular/desktop quando uma NF cai na fila do aprovador. Funciona como **canal extra** (não substitui email nem Teams).

Como funciona: usuário ativa notificacoes em **Configuracoes** -> browser cria uma `PushSubscription` única -> backend salva ela no SharePoint vinculada ao email -> quando dispara `notificar(evento, destinatarios)`, o backend envia push via VAPID pra todos os endpoints registrados de cada destinatario.

---

## Acao 1 — Adicionar VAPID keys no Azure SWA

As VAPID keys autenticam o servidor que envia o push. As suas foram **geradas localmente** (rodaram uma vez no script `python3` do dev). Salve com cuidado — a `VAPID_PRIVATE_KEY` é secreta e nunca deve sair do servidor.

**IMPORTANTE — NUNCA commita as VAPID keys neste arquivo.** Elas devem ficar APENAS no Azure App Settings (e num cofre de senhas pessoal). A `VAPID_PRIVATE_KEY` eh secreta e qualquer um que tenha acesso a ela pode enviar push em nome do servidor.

**Gerar VAPID keys (rodar uma unica vez):**

```bash
python3 -c "
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives import serialization
import base64
priv = ec.generate_private_key(ec.SECP256R1())
pub = priv.public_key()
pb = priv.private_numbers().private_value.to_bytes(32, 'big')
pubb = pub.public_bytes(encoding=serialization.Encoding.X962, format=serialization.PublicFormat.UncompressedPoint)
print('VAPID_PUBLIC_KEY :', base64.urlsafe_b64encode(pubb).rstrip(b'=').decode())
print('VAPID_PRIVATE_KEY:', base64.urlsafe_b64encode(pb).rstrip(b'=').decode())
"
```

3 variaveis vao no Azure (NAO no arquivo):

```
VAPID_PUBLIC_KEY  = <gerada acima>
VAPID_PRIVATE_KEY = <gerada acima — SECRETA>
VAPID_SUBJECT     = mailto:datanalytics@pronep.com.br
```

**Onde colocar:**

1. Azure Portal -> Static Web App `purple-forest-09588fe10` -> **Configuration**
2. Clica em **+ Add** e adiciona as 3 variáveis acima
3. Save
4. Reinicia o SWA (Overview -> Restart) pra Functions pegarem as novas envs

> **Importante:** se você gerar novas VAPID keys depois, todas as subscriptions registradas viram inválidas (browsers vão rejeitar). Manter as mesmas pra sempre.

---

## Acao 2 — Criar lista SharePoint PRONEP-NF-PushSubscriptions

1. Vá em `https://pronepadmin.sharepoint.com/sites/Aprovacao-NotasFiscaisServicos`
2. **+ Novo** -> **Lista** -> **Lista em branco** (ou **+ New -> List -> Blank list** em ingles)
3. Nome: `PRONEP-NF-PushSubscriptions`
4. Crie estas colunas (botão **+ Adicionar coluna** / **+ Add column**):

   | Nome (display) | Tipo (PT-BR) | Tipo (EN) | Notas |
   |---|---|---|---|
   | `Title` (já vem) | Uma linha de texto | Single line of text | Email do usuário — coluna padrão, NÃO precisa criar |
   | `Endpoint` | **Várias linhas de texto** | **Multiple lines of text** | URL do push service (pode >255 chars) |
   | `P256DH` | Uma linha de texto | Single line of text | Chave pública da subscription |
   | `Auth` | Uma linha de texto | Single line of text | Auth secret da subscription |
   | `UserAgent` | Uma linha de texto | Single line of text | User-agent do browser (debug) |

> **Endpoint precisa ser Multi-line / Várias linhas** porque URLs do FCM/Apple/Mozilla podem ter 200-400 caracteres.

5. Salvar

A Function `PushSubscribe` grava nesta lista automaticamente quando o user ativa.

---

## Acao 3 — Garantir permissoes Graph no App Reg

Já existentes do sistema. **Não precisa mexer.** As Functions usam as mesmas permissões `Sites.ReadWrite.All` (Application) que já estavam.

---

## Acao 4 — Verificar package.json e deploy

O `api/package.json` já foi atualizado pra incluir `web-push@^3.6.7`. Quando o GitHub Actions rodar o deploy, ele faz `npm install` automaticamente e instala a lib. **Não precisa fazer nada manual.**

---

## Acao 5 — Habilitar pelo lado do usuário

Cada usuário precisa ativar individualmente em **cada dispositivo** que quiser receber:

1. Acessa o sistema -> **Configurações**
2. Card **"Notificações no celular / desktop"** -> botão **🔔 Ativar notificacoes**
3. Browser pede permissão -> **Permitir**
4. Pronto. Pode testar via botão **▶ Testar notificacao** (notificacao local).

Pra receber em **iPhone**: precisa estar com o app **adicionado a tela de inicio** (PWA instalado) ANTES de ativar. iOS Safari só suporta push em PWAs instalados.

---

## Como funciona depois

1. Você lança uma NF -> backend dispara `notificar('lancada', [emailAprovador], dados)`
2. `notificar()` envia: email + Teams + **push** (best-effort)
3. Push chega no celular/desktop do aprovador como notificacao nativa do SO
4. Aprovador toca/clica na notif -> abre o app direto na Fila de Aprovacao
5. Caso o aprovador tenha múltiplos dispositivos (celular + desktop), recebe em todos

**Eventos que disparam push:**
- `lancada` -> "Nova NF para aprovar"
- `aprovada` -> "NF aprovada"
- `rejeitada` -> "NF rejeitada"

---

## Troubleshooting

**Botão "Ativar notificacoes" não aparece / fica em loading:**
- VAPID keys não configuradas. Confere Acao 1.
- `/api/PushPublicKey` retorna 500 (env var faltando).

**Browser diz "Notificacoes bloqueadas":**
- Usuário negou anteriormente. Tem que abrir as configurações do site no browser e mudar pra "Perguntar/Permitir" manualmente.

**Push não chega no celular:**
- Em iPhone: confere se o app está instalado como PWA (Adicionar à Tela de Início). Push só funciona em PWA instalado em iOS.
- Em Android: confere se o app não está em modo "Battery saver agressivo" (alguns fabricantes bloqueiam push em background).
- Confere no SharePoint se a subscription foi criada na lista `PRONEP-NF-PushSubscriptions`.

**Endpoint expirou (410 Gone):**
- O sistema remove automaticamente subscriptions inválidas. Usuário precisa reativar em Configurações se quiser continuar recebendo.

**Quero ver os logs:**
- Application Insights do SWA -> filtrar por `[pushNotif]` ou por `enviarPushPraEmail`.
