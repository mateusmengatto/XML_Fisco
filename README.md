# XML Fisco
XML Fisco é um programa que analisa uma NFe Brasileira em formato XML e calcula os impostos não destacados na NFe. Fornece uma interface amigável permitindo que os usuários selecionem um arquivo XML, executem a análise, visualizem os resultados e realizem ações adicionais, como editar a tabela NCM e imprimir relatórios.


## Recursos
- Selecionar Arquivo XML: Os usuários podem escolher um arquivo XML contendo informações relacionadas a impostos.
- Executar Análise: Analisa o arquivo XML selecionado e fornece resultados detalhados.
- Editar Tabela NCM: Abre a tabela NCM no Excel para edição.
- Abrir no Excel: Abre os resultados gerados das planilhas no MS Excel ou programa padrão de abertura xlsx
- Imprimir Relatório: Permite a impressão do relatório de análise.


## Requisitos
Python 3.x
PySide6
pandas
numpy
xml.etree.ElementTree
xml.dom.minidom
subprocess

## Como utilizar
1.Executar Aplicativo: Execute o script para iniciar o aplicativo.

2 Selecionar Arquivo XML: Clique no botão "Selecionar Arquivo XML" para escolher um arquivo XML.
![image](https://github.com/mateusmengatto/XML_Fisco/assets/65681163/6d39d756-e43d-44ce-8dcc-48dfacfcace7)

3 Executar Análise: Clique no botão "Executar Análise" para analisar o arquivo XML selecionado.

4. Visualizar Resultados: Os resultados da análise serão exibidos em tabelas dentro do aplicativo.
5. ![image](https://github.com/mateusmengatto/XML_Fisco/assets/65681163/7e119791-3ff6-43ae-a598-d522f5baae95)

## Ações Adicionais:
1 . Editar Tabela NCM: Clique no botão "Editar Tabela NCM" para abrir a tabela NCM no Excel para edição.
    Atenção: Os padrões de string devem ser seguidos de acordo com a tabela modelo.
2.  Imprimir Relatório: Clique no botão "Imprimir relatório" para imprimir o relatório de análise.

## Observações
Certifique-se de que as bibliotecas necessárias estão instaladas (PySide6, pandas, numpy).
O arquivo da tabela NCM (TabelaNCM.xlsx) deve estar no mesmo diretório que o aplicativo.

O XML Fisco esta ajustado para um MVA interno de 19,5% referente ao estado do Paraná e empresa de regime fiscal normal. Se for utilizar para outros estados e ou o simples nacional, verificar o MVA do seu estado e mudar as configurações para exibir o cálculo de DIFAL (que ocorre em alguns casos). Possiveis erros e bugs podem ocorre pelos parâmetros utilizados serem específicos, contudo as próximas versões do XML Fisco terão as opções para mudança dos parâmetros atráves em interface gráfica.

## Licença
Este projeto está licenciado sob a Licença GNU General Public License. Consulte o arquivo LICENSE para obter mais detalhes.

## Contato
Se você tiver dúvidas ou sugestões, sinta-se à vontade para entrar em contato através do email: mateusmengatto@gmail.com





