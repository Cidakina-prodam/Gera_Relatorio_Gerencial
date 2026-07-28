# Como publicar o app (sem precisar programar)

Você vai precisar de 2 coisas gratuitas: uma conta no GitHub (você já tem) e uma conta no
Streamlit Community Cloud (criar com o mesmo login do GitHub, 1 clique).

## Passo 1 — Criar o repositório no GitHub

1. Acesse **github.com** → botão verde **"New"** (ou "Novo repositório").
2. Nome sugerido: `conferencia-gds1`
3. Marque **Private** (só você acessa o código) ou Public, como preferir.
4. **Não** marque nenhuma opção de "adicionar README" — deixe vazio.
5. Clique em **Create repository**.

## Passo 2 — Subir os arquivos (sem usar linha de comando)

1. Na página do repositório recém-criado, clique em **"uploading an existing file"**
   (ou "enviar arquivos existentes").
2. Arraste estes 4 arquivos pra dentro da página:
   - `app.py`
   - `requirements.txt`
   - `template_relatorio.docx` (o modelo já formatado — o app preenche em cima dele)
   - `COMO_PUBLICAR.md` (opcional, é só este guia)
3. Clique em **Commit changes** (botão verde, embaixo).

## Passo 3 — Publicar no Streamlit Community Cloud

1. Acesse **share.streamlit.io**
2. Clique em **Sign in with GitHub** e autorize.
3. Clique em **"New app"** (ou "Create app").
4. Escolha:
   - **Repository**: `conferencia-gds1` (o que você criou)
   - **Branch**: `main`
   - **Main file path**: `app.py`
5. Clique em **Deploy**.

Em 1-2 minutos o app fica no ar, com um link fixo tipo:
`https://conferencia-gds1-xxxxx.streamlit.app`

Salva esse link nos favoritos — é só abrir, subir os arquivos do Portal e conferir na hora,
qualquer dia, sem precisar de mim.

## Atualizações futuras

Se eu (Claude) ajustar o `app.py` depois (por exemplo, pra adicionar o cálculo de despesa por
fornecedor), é só repetir o Passo 2 (subir o arquivo novo no mesmo repositório, sobrescrevendo)
— o Streamlit Cloud atualiza sozinho em ~1 minuto.

## Se algo der errado

- **"ModuleNotFoundError"** ao abrir o app → confira se o `requirements.txt` foi mesmo enviado
  junto no repositório.
- **App "dormindo"** (Streamlit Cloud gratuito hiberna apps sem uso por alguns dias) → é só abrir
  o link, ele acorda sozinho em ~30 segundos.
