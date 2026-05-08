# Conferencia de Caixa

Ferramenta web para importar o relatorio CX do sistema, montar a fila de conferencia dos movimentos que afetam o caixa e apontar divergencias de totais, dinheiro fisico, Pix CNPJ, vales, cartoes, cheques/pix, pre e devolucoes.

## Rodar localmente

```powershell
cd "C:\Users\kauar\Documents\Codex\2026-05-07\preciso-criar-uma-ferramenta-para-facilitar"
C:\Users\kauar\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe server.mjs
```

Depois acesse:

```text
http://127.0.0.1:4173
```

## Deploy na Vercel

1. Crie um repositorio no GitHub.
2. Envie estes arquivos para o repositorio.
3. Na Vercel, clique em `Add New...` e depois `Project`.
4. Importe o repositorio do GitHub.
5. Framework Preset: `Other`.
6. Build Command: deixe vazio ou use `npm run vercel-build`.
7. Output Directory: deixe vazio.
8. Deploy.

## Arquivos importantes

- `public/`: tela da ferramenta.
- `api/parse-cx.mjs`: API serverless usada pela Vercel para ler o XLSX.
- `lib/parse-cx-core.mjs`: regras de conferencia e extracao do CX.
- `server.mjs`: servidor local para uso no computador da empresa.
- `vercel.json`: configuracao da Vercel.

## Seguranca

Ao hospedar na Vercel, o arquivo CX sera enviado para o servidor da Vercel no momento da importacao. A funcao nao salva o arquivo em disco, mas e recomendado adicionar login antes de usar em producao.
