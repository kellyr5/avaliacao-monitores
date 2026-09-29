# Avaliação da plataforma UNIFEI

Dois formulários estáticos publicados no GitHub Pages que gravam as respostas numa planilha do Google.

- `index.html`: formulário dos monitores.
- `aluno.html`: formulário dos alunos (o que a plataforma agrega e quais componentes são mais úteis).
- `apps-script.gs`: script que recebe as respostas, grava em abas separadas e monta a aba `Estatisticas`.

## Configuração

1. Crie uma planilha no Google Planilhas.
2. Vá em Extensões > Apps Script, cole o conteúdo de `apps-script.gs` e salve.
3. Rode a função `configurar` uma vez, para criar as abas e as estatísticas.
4. Implante como "App da Web": executar como "Eu", acesso para "Qualquer pessoa". Copie a URL terminada em `/exec`.
5. Cole a URL em `const ENDPOINT = "";` nos dois arquivos HTML.
6. Publique no GitHub Pages (branch `main`, pasta raiz).

## Análise

A aba `Estatisticas` calcula média, mediana, desvio padrão e concordância (notas 4 e 5) de cada item, além das distribuições de uso, funcionalidade mais útil, canal atual das dúvidas e curso.
