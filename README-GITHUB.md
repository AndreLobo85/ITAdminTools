# Publicar no GitHub & distribuir

Este doc explica como publicares o `ITAdminToolkit` no teu GitHub e depois como o colega/tu instalares no PC do trabalho com **um único comando**.

## 1. Criar o repo (uma vez)

No GitHub, cria um repo (público ou privado):
- **Nome**: `ITAdminToolkit` (ou outro)
- **Visibilidade**: privado se preferires — o installer funciona na mesma desde que tenhas acesso

Localmente (nesta pasta):

```powershell
cd D:\AI\QWEN\Projetos\ITAdminToolkit
git init
git remote add origin https://github.com/AndreLobo85/ITAdminTools.git
```

Os defaults ja apontam para `AndreLobo85/ITAdminTools`. Se mudares de repo, ajusta `$Repo` em:
- `Install-ITAdminToolkit.ps1`
- `Update-App.ps1`

Primeiro commit:

```powershell
git add -A
git commit -m "initial import"
git push -u origin main
```

O `.gitignore` exclui `lib/`, `webui/assets/vendor/`, `webui/assets/fonts/` e logs — o source no GitHub fica limpo (~5MB).

## 2. Publicar uma release (cada versão)

Cada release é um zip pré-built que inclui **tudo** (DLLs WebView2, React, fonts). O PC do trabalho descarrega este zip, não precisa de correr nenhum Setup-*.ps1.

```powershell
# 1. Certifica-te que a versao em version.json esta correcta
# 2. Build do zip
.\Build-Release.ps1

# 3. Commit & tag
git add -A
git commit -m "release v1.0.0"
git tag v1.0.0
git push --tags

# 4. Criar release no GitHub (usa a CLI gh, ou faz pela UI)
gh release create v1.0.0 .\release\ITAdminToolkit-v1.0.0.zip `
    --title "v1.0.0" `
    --notes "Primeira release publica"
```

Se não tiveres a `gh` CLI, vai ao GitHub web → **Releases** → **Draft a new release** → escolhe a tag `v1.0.0` → upload do zip → publica.

## 3. Instalar no PC do trabalho

### Primeira vez

Abre PowerShell (user normal, não precisa admin) e cola:

```powershell
iex (irm https://raw.githubusercontent.com/AndreLobo85/ITAdminTools/main/Install-ITAdminToolkit.ps1)
```

O installer:
1. Detecta/instala o **WebView2 Runtime** (silencioso)
2. Descarrega a última release do GitHub
3. Extrai para `%LOCALAPPDATA%\ITAdminToolkit\ITAdminToolkit\`
4. Cria um **atalho no ambiente de trabalho** ("IT Admin Toolkit")
5. Lança a app

Para **repositório privado**, tens de autenticar. Opções:
- `gh auth login` (GitHub CLI) antes do `iex`
- Usar Personal Access Token: `$env:GITHUB_TOKEN = 'ghp_...'` antes do `iex` (requer adaptar o installer para usar este token nos headers — diz-me se precisas)

### Actualizações futuras

Mesmo one-liner re-corrido:

```powershell
iex (irm https://raw.githubusercontent.com/AndreLobo85/ITAdminTools/main/Install-ITAdminToolkit.ps1)
```

- Se a versão instalada == latest → mostra "já tens a versão mais recente" e lança
- Se for mais antiga → descarrega + substitui + lança

Ou corre o `Update-App.ps1` que já ficou instalado localmente.

## 4. Flow de desenvolvimento

```
dev (nesta maquina)                   user (PC trabalho)
─────────────────                     ───────────────────
editas codigo                             .
  |                                       .
git commit + push                         .
  |                                       .
bump version em version.json             .
  |                                       .
Build-Release.ps1 → .zip                  .
  |                                       .
git tag vX.Y.Z && push                    .
  |                                       .
gh release create ...  ────────────>  re-corre one-liner iex(irm ...)
                                       → app actualiza
```

## 5. Próximos passos opcionais

- **Botão de update dentro da app**: posso adicionar uma pill "Update disponível" no titlebar que chama `Update-App.ps1`
- **Repo privado**: configurar autenticação com PAT
- **Assinatura digital**: assinar as `.ps1` com um certificado de code-signing (elimina warnings de SmartScreen)
- **SCCM / Intune package**: empacotar como MSIX para distribuição corporativa
