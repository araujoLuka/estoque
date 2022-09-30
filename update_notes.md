# **Anotações de Atualizações Realizadas** *(Atualização 2.10)*

## Modulo *xmlScripts*
- Entrada por XML implementada e testada
- Sub *entradaXML* serve como guia para botões para chamar o procedimento de importação de XML
- sub *movXML* atualizado
	- Carrega o XML em uma matriz
	- Exibe uma tabela na tela com o formulario *xmlForm*
		> Leia sobre o novo formulario [xmlForm](##Formulario-de-Login) para mais informações
	- Se confirmado a movimentação, checa se os produtos estão cadastrados/atualizados
		- Caso não estejam, usa a nova function *geraVetorCadXML*
			> Leia atualizações do módulo [cadScripts](##Modulo-*cadScripts*) para mais informações
		- Para registrar a movimentação usa-se a nova function *geraVetorMovXML*
			> Leia atualizações do modulo [movScripts](##Modulo-*movScripts*) para mais informações
- Nova Function *checkData*
	- Recebe dois parametros:
		- Vetor com informações cadastradas
		- Vetor com informações do XML
	- Verifica se existe alguma diferença entre os dois vetores
	- Em caso de diferença, exibe uma mensagem para perguntar se o usuario deseja aceitar as alterações do XML
		- Se não for aceito as alterações, considera os vetores como iguais
	- Retorno é um booleano: True para vetores iguais, caso contrario retorna False

## Modulo *cadScripts*
- Nova Function *geraVetorCadXML*
	- Recebe um parametro obrigatorio e dois opcionais:
		- Uma linha da matriz de dados importados do XML
		- O limite e o estoque do produto (opcionais)
	- Retorna um vetor com informações do XML para cadastrar/atualizar um produto
- Sub *removeProduto* atualizado
	- Codigo modularizado
	- Aplicado tecnica de desativar atualizacao de tela para melorar eficiencia
	- Adicionado 'msgbox' para verificar exclusao de registros de movimentacao
		- Utiliza o sub "removeMovimMult"
- Corrigido bug na sub *cadastraProduto*, ao inserir primeiro produto do sistema
- Formula para contagem de linhas atualizada para corrigir erro ao excluir produto da primeira linha do cadastro
	- Formula é utilizada nas functions *geraVetorCad*, *geraVetorCadXML* e *geraVetorEstoque*


## Modulo *movScripts*
- Subs *remMov* e *buscaMov* renomeados para *removeMovim* e *buscaMovim*
- Sub *deleteAllMovCod* renomeado para *removeMovimMult*
- Nova Function *geraVetorMovXML*
	- Recebe quatro parametros:
		- Uma linha da matriz de dados importados do XML
		- O numero da nota fiscal extraido do XML
		- A data da movimentacao
		- A hora em que foi iniciado a importacao do XML
	- Retorna um vetor com informações do XML para registrar a movimentacao de estoque

## Modulo *estqScripts*
- Novo Sub *removeEstoque*
	- Recebe o codigo do produto a ser removido
	- Busca e remove o produto da planilha *Estoque*
- Subs *listaCompra* e *listaEstoqueB* atualizadas para casos de não haver produtos na condição especificada por cada Sub

## Formulario de Login
- Novo nome *logginForm*
- Atualizado para chamar procedimento *loggin_A*
	> Leia alterações do módulo [logScripts](##Modulo-*logScripts*) para mais informações
- Corrigido ponteiro dinamico do mouse
- Corrigido bug ao tentar logar com campos em branco

## Modulo *logScripts*
- Sub *loggin* renomeado para *iniciaLoggin*
	- Somente chama o formulario *logginForm*
- Sub *loggin_A* atualizado
	- Agora é o procedimento principal para acessos
	- Recebe dois parametros opcionais para ser chamar pelo formulario
- Function *buscaAcesso* atualizada
	- Retorna um intervalo e não mais uma string
	- Loop atualizado para verificar um vetor e não um intervalo
		- Melhora a eficiencia do Excel
- Sub *planAccess* atualizado
	- Recebe dois parametros: 
		- Vetor com informações de acesso
		- Intervalo que indica o ultimo acesso
	- Apos finalizar acesso verifica se o ultimo acesso é diferente da data atual
		- Se sim, exibe uma mensagem de boas-vindas
		> Leia o topiceo [Outras atualizações](##Outras-atualizações) para mais informações sobre 'ultimo acesso'

## Modulo *tblScripts*
- Sub *sortCad* renomeado para *sortTbl* e atualizado
	- Ajustado para ordenar qualquer tabela que possua ao menos uma coluna de "PRODUTOS" e/ou de "TIPO"
	- Funcionado para a planilha de "Estoque"

## Modulo *functions*
- Corrigido erros na Function *bSearch_c*
- Function *validaForm* atualizado
	- Implementado verificação de formato nos campos digitados (Ex. Codigo deve ser numerico e Produto deve ser texto)
	- Ajustado "set-and-change" no algoritmo
	- Inserido mensagens de aviso para erros (OBS.: Muito necessario criar modulo para mensagens de aviso)

## Outras atualizações
- Planilha *Acesso* atualizada e corrigido bugs
- Planilha *Usuarios* modificada
	- Adicionado duas colunas na tabela de usuarios:
		- **acesso** - para identificar o tipo de acesso do usuario (Para implementações futuras)
		- **ultimo acesso** - para guardar a ultima data de acesso do usuario
- Aumentado fonte e tamnho da linha nas planilhas *Cadastro* e *Estoque*
- Adicionado coluna *TIPO* na planilha *Estoque* para permitir ordenar tabela
- Corrigido bugs no formulario *movForm*
- Corrigido bug que permitia excluir uma quantidade maior que o estoque no formulario *mvmForm*
- Nova constante *PROP_SIZE* no modulo *iconScripts*
	- Define a proporcao de tamanho do icone em relacao à celula que ele se encontra
- Implementado icones "simulados" para caso de erro ao inserir imagens de icones
