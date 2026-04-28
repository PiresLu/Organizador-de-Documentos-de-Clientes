ORGANIZADOR DE DOCUMENTOS DE CLIENTES v4.1
==========================================


CONFIGURACAO INICIAL (apenas uma vez)
--------------------------------------
1. Clique com o botao direito em:
      Criar_Atalho.ps1

2. Selecione "Executar com PowerShell"

3. O atalho "Organizar Docs Clientes" sera criado no Desktop.

   ATENCAO: Nao mova nem renomeie os arquivos desta pasta apos criar o atalho.


USO DIARIO
----------
Abra o atalho no Desktop e escolha uma opcao:

   [1]  Operadora individual  >>  Somente pasta Enviados   <- uso padrao
   [2]  Operadora individual  >>  Varredura Completa
   [3]  Operadora individual  >>  Simulacao (sem alterar nada)
   [4]  Todas as operadoras   >>  Somente pasta Enviados
   [5]  Todas as operadoras   >>  Varredura Completa
   [0]  Sair


REGRA FUNDAMENTAL
-----------------
Para o sistema processar um documento, ele deve estar dentro
da pasta "Enviados" da operadora correspondente.

   CORRETO  >>  ...\0041 - Jau\Enviados\123456 - Nome.pdf
   IGNORADO >>  ...\0041 - Jau\123456 - Nome.pdf

Isso garante que apenas propostas concluidas sejam enviadas
ao R:\, mantendo as pendentes intactas.


SEGURANCA
---------
Nenhum arquivo e excluido permanentemente.

Tudo que e removido vai para:
   R:\_LIXEIRA_SCRIPT\<data>\<operadora>\

Um log completo de cada execucao e salvo em:
   R:\_LIXEIRA_SCRIPT\Logs\

Em caso de duvida, use a opcao [3] Simulacao antes de executar.
Ela mostra exatamente o que sera feito, sem alterar nada.


PROBLEMAS CONHECIDOS
--------------------
Se o atalho nao funcionar:
   - Abra o PowerShell manualmente
   - Navegue ate a pasta dos scripts
   - Execute: .\Menu_Operadoras.ps1

   Ou use o comando abaixo diretamente no PowerShell:
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "T:\Cadastro\Grupo Regional 3\Organizar Docs Clientes\Menu_Operadoras.ps1"


CONTATO TECNICO
---------------
Em caso de problemas, entre em contato com o responsavel
pelo sistema antes de executar qualquer operacao manual.
