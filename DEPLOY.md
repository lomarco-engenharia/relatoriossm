# Deploy — GitHub Pages

## Pré-requisito único: Git instalado no computador

## Passo 1 — Clonar o repositório (apenas na primeira vez)

```bash
git clone https://github.com/lomarco-engenharia/relatoriossm.git
cd relatoriossm
```

## Passo 2 — Copiar os arquivos da Fase 2

Copie **todo o conteúdo** da pasta `fase2/` para dentro da pasta `relatoriossm/`.
A estrutura deve ficar:

```
relatoriossm/
  index.html
  dashboard.html
  projects.json
  etl.py
  requirements.txt
  .gitignore
  .github/workflows/processar_csv.yml
  projects/eteca/
  projects/etemp/
  data/eteca/
  data/etemp/
  entrada_csv/eteca/.gitkeep
  entrada_csv/etemp/.gitkeep
```

## Passo 3 — Commit e push inicial

```bash
cd relatoriossm
git add .
git commit -m "Fase 2: site único multi-projeto ETECA + ETEMP"
git push origin main
```

> Se o branch principal for `master`, use `git push origin master`.

## Passo 4 — Ativar GitHub Pages

1. Abra: https://github.com/lomarco-engenharia/relatoriossm/settings/pages
2. Em **Source**, selecione o branch `main` (ou `master`) e pasta `/` (root)
3. Clique **Save**
4. Aguarde 1–2 minutos
5. O site estará em: `https://lomarco-engenharia.github.io/relatoriossm/`

## Passo 5 — Ativar permissão de escrita para o Actions

1. Abra: https://github.com/lomarco-engenharia/relatoriossm/settings/actions
2. Em **Workflow permissions**, selecione **"Read and write permissions"**
3. Clique **Save**

---

## Como atualizar com novo CSV

1. Baixe os CSVs da plataforma We Handle
2. No GitHub, navegue até `entrada_csv/<projeto>/`
3. Clique em **Add file → Upload files**
4. Suba os CSVs com os nomes originais (mantendo o timestamp)
5. Commit direto na branch `main`
6. O GitHub Actions roda automaticamente e atualiza o site

---

## Testar localmente antes do deploy

```bash
cd relatoriossm
python -m http.server 8080
# Abrir: http://localhost:8080
```
