# extrato simplificado do fluxo do treinamento de 05/10 - turma FHEMIG - Progressão de servidores no SISAP

- iniciar excel (arrastar e soltar) doc em branco ou doc a seguir (planilha predeterminada - janela com caminho do documento) instância visível; somente leitura NÂO; armazenada em variável %Progressao

- executar aplicativo: caminho (Desktop ou arquivos de programas86); estilo janela maximizado; aguardar carregamento aplicativo; 5 segundos

- adicionar variável de entrada (menu direito) 
	- nome da variável = nome externo = login (texto) m1565211
	- nome variável, confidencial, senha mg1061

- enviar teclas SISAP
	- inserir teclas especiais - diversos - TAB
	- selecionar variável %Login
	- inserir teclas especiais - diversos - TAB
	- inserir variável %senha
	- inserir teclas especiais - diversos - enter {return}
	- atraso de 10miliseconds
	- SISAP não aceita teclado virtual (ENVIAR TEXTO COMO CHAVES DE HARDWARE)

- aguardar 1 segundo (para poder logar)

- enviar teclas {CTRL}({A}){CTRL}({C})
	- inserir modificador {CTRL}({A}){CTRL}({C}), 80 miliseconds
- obter texto da área de transferência variável produzida %check%

- IF %Check% contém `Logon executado com sucesso`
	- Enviar teclas SIAP{Return}
	- Else
	- Enviar teclas {Return} 
	- End

- rótulo (VOLTAR) entre 4 e 13, antes de END do if

- enviar teclas 
	- X{Return}{F3}

- aguardar 1seg

- enviar teclas
	- evo *{Return}{Tab}X{Return}01{REturn}
- ler da planilha excel (entre 15 e 16) 
	- instância do excel
	- valor de uma única célula
	- coluna A
	- linha 2
	- nova variável MASP
- enviar teclas
	- evo *{Return}{Tab}X{Return}01{REturn}%MASP%{Return}

- texto(célula; "DDMMAAA")

- ler da planilha 
	- progressao
	- unica celula
	- B
	- 2
	- admissão

- ler da planilha 
	- progressao
	- unica celula
	- C
	- 2
	- evolucao

- ler da planilha 
	- progressao
	- unica celula
	- F
	- 2
	- publicacao

- ler da planilha 
	- progressao
	- unica celula
	- G
	- 2
	- inicio

- enviar teclas
	- evo *{Return}{Tab}X{Return}01{REturn}%MASP%%admissao%{Return}%evolucao%{Tab}%Publicacao%%inicio%{Return}{F3}

- obter primeira linha livre da planilha do excel (entre 14 e 15)
	- coluna A
	- variável produzida: tamanhoPlanilha
- obter primeira linha livre da planilha do excel (entre 15 e 16)
	- coluna H
	- variável produzida: linhaAtual
- condição de loop (entre 16 e 17 - após obter e antes de 'aguardar')
	- linha atual diferente de tamanhoPlanilha (enquanto linha A for diferente de H

- retirar referência da linha 2 e atribuir variável linha atual

- gravar na planilha do excel (anres do end)
	- instância excel: %progressao%
	- valor CONCLUÍDO
	- na célula especificada
	- coluna H
	- %linhaatual%A

- aumentar variável
	- nome: %linhaAtual%
	- aumentar em 1  

- enviar teclas (Alt F4 para fechar SISAD)

- fechar excel (Salvar)

- shutdown (só desliga o computador se salvar o fluxo)

**para copiar o código, ctrl A + ctrl C e salvar manualmente em bloco e notas**
