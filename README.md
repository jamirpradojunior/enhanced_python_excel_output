Neste código, você irá explorar diversas funções para personalizar seu output Python para Excel.

Definição do problema

Personalizar um arquivo Excel pode demandar bastante tempo, ainda mais quando os dados vêm de fontes externas. Além disso, a padronização desses outputs é fundamental para construção de relatórios de qualidade. Por isso, nada melhor que automatizar essas tarefas!

Para exemplificar uma rotina, vamos construir um arquivo Excel que mostra um relatório de uma carteira teórica de ativos. Os datasets foram extraídos do Yahoo Finance.
* Primeiro dataset é composto pelos dados de cotações (abertura, máxima, mínima e fechamento) desses ativos cobrindo um período de um ano.
* Nosso segundo dataset são os eventos de pagamentos de proventos (ex., dividendos, JCP) dos ativos. 
* O relatório mostrará a participação desses ativos na carteira, valor de mercado inicial e atual e o saldo líquido de cada ativo.
* Para gerar o output Excel personalizado, utilizaremos a biblioteca xlsxwriter. Busquei cobrir grande parte das seções da documentação para explorar as diferentes funções.
