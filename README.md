# Lista-2-Ex-3
descrição : > -
  Layout do suplemento: a coisa mais simples do mundo, um único botão!
  Comportamento do suplemento: quando o botão é clicado, a planilha ativa no
  Excel deve ser limpa e, então, uma API pública sobre os planetas do universo
  Star Wars deve ser lida de forma assíncrona em formato JSON. Quando o
  resultado para imprimir, uma lista de planetas deste universo deve ser exibida
  linha a linha na planilha ativa, exibindo-se as colunas: Nome, Clima,
  População.
  Tempo estimado para conclusão: 1 hora.
anfitrião : EXCEL
api_set : {}
script :
  conteúdo : |
    run.addEventListener ("click", () => {
      Excel.run (função (contexto) {
        deixe currentWorksheet = context.workbook.worksheets.getActiveWorksheet ();
        currentWorksheet.getRange (). clear ();
        deixe despesasTable = currentWorksheet.tables.add ("A1: C1", true / * hasHeaders * /);
        ablesTable.name = "ExpensesTable";
        despesasTable.getHeaderRowRange (). values ​​= [["Nome", "Clima", "População"]];
        deixe planet_list = [];
        fetch ("https://swapi.dev/api/planets/") .then (função (resposta) {
          response.json (). then (function (data) {
            for (let planet of data.results) {
              planet_list.push (
                [planet.name, planet.climate, planet.population]
              )
            }
            despesasTable.rows.add (null, planet_list);
            despesasTable.getRange (). format.autofitColumns ();
            despesasTable.getRange (). format.autofitRows ();
            return context.sync ();
          });
        }). catch (function (err) {
          console.error ('Falha ao recuperar informações', err);
        });
        
      }). catch (função (erro) {
        console.log ("Erro:" + erro);
        if (instância de erro de OfficeExtension.Error) {
          console.log ("Informações de depuração:" + JSON.stringify (error.debugInfo));
        }
      });
    });
  linguagem : texto datilografado
modelo :
  conteúdo : | -
    <button id = "run" class = "ms-Button">
        <span class = "ms-Button-label"> Executar </span>
    </button>
  linguagem : html
estilo :
  conteúdo : | -
    section.samples {
        margem superior: 20px;
    }
    section.samples .ms-Button, section.setup .ms-Button {
        display: bloco;
        margin-bottom: 5px;
        margem esquerda: 20px;
        largura mínima: 80px;
    }
  idioma : css
bibliotecas : |
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js
  @ types / office-js
  office-ui-fabric-js@1.4.0/dist/css/fabric.min.css
  office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css
  core-js@2.4.1/client/core.min.js
  @ types / core-js
  jquery@3.1.1
  @ types / jquery @ 3.3.1
