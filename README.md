# gerador-de-documento
Aplicação WPF para gerar documentos docx

A aplicação foi feita para preencher um template docx com valores passados pelo usuário. 

Para instalar a aplicação para uso pessoal, siga os seugintes passos:
 - Clone o repositório e abra no Visual Studio.
 - Após abrir o projeto no Visual Studio, monte a tela com os campos dinâmicos do relatório (Form1.Designer.cs)
 - Após montar a tela, altere os nomes dos atributos no objeto json e as faça receber o valor do campos que você criou (classe Form.cs e altere somente entre as linhas 26 e 50)
 - Feito isso, crie seu template docx com os parâmetros entre chaves ({parametro_1}). Os parâmetros deverão ter o mesmo nome dos atributos criados no objeto JSON no WPF.
  - Por exemplo, você quer um relatório que nome e sobrenome apareçam de acordo como que o usuário digitar, então no relatório coloque {nome} e {sobrenome} onde você quer que apareça.
  - Lembrando que WPF, os atributos do objeto JSON devem ter o mesmo nome (neste caso, sem as chaves {}).  
 - Salve com o nome template.aspx
 - Baixe o node.js e execute:
  - npm install docxtemplater
  - npm install jszip@2
 - Compile o projeto no Visual Studio, insira os arquivos docx.js, generator.bat e o template.docx dentro da pasta Release.
 - Salve a pasta onde deseja e crie um atalho do arquivo .exe na área de trabalho.
 - Feito isso, é só executar o atalho. Insira seus dados, clique no botão para gerar o relatório e pronto, você tem seu relatório personalizado.
