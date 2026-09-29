# Avaliação da plataforma UNIFEI pelos monitores

Formulário estático, publicado no GitHub Pages, que grava as respostas numa planilha do Google.

## 1. Planilha e script

1. Crie uma planilha nova no Google Planilhas.
2. Vá em Extensões > Apps Script e apague o conteúdo do editor.
3. Cole o conteúdo de `apps-script.gs` e salve.
4. Clique em Implantar > Nova implantação > tipo "App da Web".
5. Em "Executar como", escolha "Eu"; em "Quem tem acesso", escolha "Qualquer pessoa".
6. Autorize o acesso pedido e copie a URL terminada em `/exec`.

## 2. Ligar o formulário

Abra `index.html`, procure `const ENDPOINT = "";` e cole a URL entre as aspas.

## 3. Publicar

Envie os arquivos para um repositório no GitHub e ative o Pages em Settings > Pages, branch `main`, pasta raiz.
O endereço será `https://SEU-USUARIO.github.io/NOME-DO-REPOSITORIO/`.

## Análise

As respostas aparecem na planilha, uma linha por monitor. Exporte em CSV para calcular médias, concordância e alfa de Cronbach.
