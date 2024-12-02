Projeto de Data Analysis para ADVBOX.

Esse repo contém a minha entrada para o teste prático requerido pelos recrutadores da ADVBOX.

1) Descrição do Projeto:
Esse projeto consiste em migrar a informação de uma série de tabelas .csv para duas tabelas chamadas "CLIENTS.xlsx" e "PROCESSOS.xlsx", convertendo todas as informações e seus respectivos cabeçalhos em novos formatos e transformando a data de acordo com as regras e padronizações prescritas.

2) Conteúdo do Repo:
"migrador.py" contém todo o código da aplicação.
"requirements.txt" contém as informações dos pacotes da aplicação.
"Orientações para migração.docx" contém as instruções para esse projeto, providas pelos recrutadores da ADVBOX. Eles solicitam, além de outras coisas, que o projeto apresente uma GUI, a qual eu fiz.
"Backup_de_dados_92577.rar" contém todas as tabelas .csv originais.
A pasta "Advbox" contém as tabelas de exemplo nos seus respectivos arquivos  ("CLIENTS.xlsx" e "PROCESSOS.xlsx"). "MIGRAÇÃO PADRÕES NOVO.xlsx" contém uma lista de regras e novos padrões para a data final, o que requere a transformação da data extraída das tabelas .csv originais.
A pasta "templates" possui versões limpas das tabelas de exemplo que eu usei no meu código para construir as tabelas migradas. Eu as nomeei "CLIENTS_template.xlsx" e "PROCESSOS_template.xlsx".
"README-pt.txt" é isso que você está lendo (em PT-BR). "README.txt", que é automaticamente aberta pelo Github, é esse mesmo conteúdo em inglês.

Nesse repo eu também vou disponibilizar um .exe AIO compilado pelo pyinstaller para facilitar o uso e melhorar a deployabilidade. Eu urjo que o use.

3) Documentação:
Simplesmente rode o AIO migrador.exe, ou migrador.py. Uma GUI vai aparecer requerindo que você entre a localização do arquivo "Backup_de_dados_92577.rar". Ao fazer isso e clicando em OK a aplicação começará a processar os arquivos e irá criar as tabelas finais  "CLIENTS.xlsx" e "PROCESSOS.xlsx" no mesmo diretório onde a aplicação está localizada. Assim que a aplicação concluir sua operação ela deletará qualquer arquivo temporário criado por ela e uma janela de notificação surgirá. Clicando em confirmar irá fechar a GUI e a aplicação como um todo.

4) Conclusão
Tendo feito minha parte agora eu humildemente me defiro ao julgamento da equipe de recrutadores da ADVBOX, ou quem mais que esteja vendo isso no futuro. Eu espero que esse projeto possa ser gentilmente visto como uma fotografia do desenvolvedor que eu fui, mas não do desenvolvedor que eu busco ser.

Agradeço a oportunidade e espero ouvir de vocês em breve.