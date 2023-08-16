function quantosMinutos(hora){
    var horaUtilizada = "" + hora;
    var [horas, minutos] = horaUtilizada.split('h');
    var ss = SpreadsheetApp.getActive().getSheetByName("Dados");
    let horaAtt = ((parseInt(horas)*60)+(parseInt(minutos)*1))
    return horaAtt;
  }
  
  function buscarFaixa(horaMedia, jogadores, posicao){
    var ss = SpreadsheetApp.getActive().getSheetByName("Base Pontuação");
    var quantidadeFaixas = ss.getLastRow();
    var temposFaixas = ss.getRange(2,1,quantidadeFaixas,1).getValues();
    var numeroJogadores = [2, 2, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5];
    let sinal = 0;
    let posicaoTempo = 0;
    for(let i=0; i<temposFaixas.length; i++){
      if(sinal===0){
        if(parseInt(temposFaixas[i])>=parseInt(horaMedia)){
          sinal = 1;
          posicaoTempo = i+2;
        }
      }
    }
    var arrayPontuacoes = new Array (jogadores);
    let contador = 0;
    if(parseInt(jogadores)<6){
      for(let j=0; j<numeroJogadores.length; j++){
        if(numeroJogadores[j] == parseInt(jogadores)){
          let pontuacao = parseInt(ss.getRange(posicaoTempo, j+2).getValue());
          arrayPontuacoes[contador] = pontuacao;
          contador = contador + 1;
        }
      }
    }else{
      nomeJogadores = 5;
      let sobram = parseInt(jogadores) - 5;
      for(let k=0; k<numeroJogadores.length; k++){
        if(numeroJogadores[k] == parseInt(nomeJogadores)){
          let pontuacao = parseInt(ss.getRange(posicaoTempo, k+2).getValue());
          arrayPontuacoes[contador]=(parseInt(pontuacao))+sobram;
          contador = contador + 1;
        }
      }
      for(let m=5; m<parseInt(jogadores); m++){
        arrayPontuacoes[contador] = parseInt(jogadores) - (m+1);
        contador = contador + 1;
      }
      arrayPontuacoes[parseInt(jogadores) - 1] = 0;
    }
    return arrayPontuacoes[parseInt(posicao)-1];
  }
  
  function quantasHorasTem(nome){
    var ss = SpreadsheetApp.getActive().getSheetByName("Jogadores"); 
    let ultJogador = ss.getLastRow();
    let nomeMaius = nome.toUpperCase();
    var nomesJogadores = [parseInt(ultJogador) - 2];
    var horarios = [parseInt(ultJogador) - 2];
    for(let j=2; j<(parseInt(ultJogador)+1); j++){
      nomesJogadores[j]=ss.getRange(j, 1).getValue();
      horarios[j]=ss.getRange(j, 3).getValue();
    }
    let horas = 0;
    for(let i=0; i<nomesJogadores.length; i++){
      if(nomesJogadores[i]==nomeMaius){
        horas = horas + parseInt(horarios[i]);
      }
    }
    return horas;
  }
  
  function maisHoras(){
    var ss = SpreadsheetApp.getActive().getSheetByName('Jogadores');
    var quantidadeJogadores = ss.getLastRow();
    var temposJogadores = ss.getRange(2,3,parseInt(quantidadeJogadores),3).getValues();
    var nomeJogadores = ss.getRange(2,1,parseInt(quantidadeJogadores),1).getValues();
    let jogadorMaisHoras = 0;
    for (let i=0; i<temposJogadores.length; i++){
      let tempoJogador = parseInt(temposJogadores[i]);
      if(tempoJogador>jogadorMaisHoras){
        jogadorMaisHoras = tempoJogador;
      }
    }
    return jogadorMaisHoras; 
  }
  
  function novoConvidado(nome){
    var ss = SpreadsheetApp.getActive().getSheetByName('Convidados');
    let ultConvidado = ss.getLastColumn();
    ss.getRange(1,ultConvidado+1).setValue(nome.toUpperCase());
  }
  
  function adicionaDados(linha, coluna, pontos){
    var ss = SpreadsheetApp.getActive().getSheetByName("Dados");
    ss.getRange(parseInt(linha), parseInt(coluna)).setValue(pontos);
  }
  
  function adicionaConvidados(linha, coluna, pontos){
    var ss = SpreadsheetApp.getActive().getSheetByName("Convidados");
    ss.getRange(parseInt(linha), parseInt(coluna)).setValue(pontos);
  }
  
  
  function adicionaJogadores(nomeJogador, horas, pontos){
    var ss = SpreadsheetApp.getActive().getSheetByName("Jogadores");
    var ultimaLinha = ss.getLastRow();
    var nomes = ss.getRange(2,1,ultimaLinhas,1).getValues();
    let posicao = 0;
    for(let i=0; i<nomes.length; i++){
      if(nomes[i] == toUpperCase(nomeJogador)){
        posicao = i;
      }
    }
    var pontosAntes = ss.getRange(posicao+1,2).getValue();
    let pontuacaoNova = parseInt(pontosAntes) + pontos;
    ss.getRange(posicao+1,2).setValue(pontuacaoNova);
    var horasAntes = ss.getRange(posicao+1, 3).getValue;
    let horasMinutos = quantosMinutos(horasAntes);
    let novoTempo = horasMinutos + horas;
    let horasInteiras = Math.floor(novoTempo/60);
    let minutosFaltantes = novoTempo % 60;
    var tempoFormatado = horasInteiras + ":" + minutosFaltantes;
    ss.getRange(posicao+1, 3).setValue(tempoFormatado);
  }
  
  function verificaJogador(nomeJogador){
    var ss = SpreadsheetApp.getActive().getSheetByName("Dados");
    let ultJogador = ss.getLastColumn();
    let sinal =0;
    var nomesJogadores = [(parseInt(ultJogador) - 5)];
    for(let j=5; j<ultJogador+1; j++){
      nomesJogadores[j-5] = ss.getRange(1, j).getValue();
    }
    for (let i=0; i<(nomesJogadores.length); i++){
      if(nomesJogadores[i] == (nomeJogador.toUpperCase())){
        sinal = 1;
      }
    }
    if(sinal>0){
      return true;
    } else{
      return false;
    }
  }
  
  function temConvidado(nomeConvidado){
    var ss = SpreadsheetApp.getActive().getSheetByName("Convidados");
    let ultConvidado = ss.getLastColumn();
    var nomeConvidadoMaiusculo = nomeConvidado.toUpperCase();
    var nomesConvidados = [(parseInt(ultConvidado)) - 4];
    for(let j=5; j<ultConvidado+1; j++){
      var nome = ss.getRange(1, j).getValue();
      nomesConvidados[j-5] = nome;
    }
    for(let i=0; i<nomesConvidados.length; i++){
      if(nomesConvidados[i] == nomeConvidadoMaiusculo){
        return true;
      }
    }
    return false;                 
  }
  
  function procuraPosicaoConvidado(nomeConvidado){
    var ss = SpreadsheetApp.getActive().getSheetByName("Convidados");
    let ultConvidado = ss.getLastColumn();
    var nomes=[(parseInt(ultConvidado))-5];
    for(let j=5; j<ultConvidado+1; j++){
      var nome = ss.getRange(1, j).getValue();
      nomes[j-5] = nome;
    }
    let posicao = 0;
    var nomeMaiusculo = nomeConvidado.toUpperCase();
    for(let i=0; i<nomes.length; i++){
      if(nomes[i] === (nomeMaiusculo)){
        posicao = i+5;
      }
    }
    return posicao;                 
  }
  
  function procuraPosicaoDados(nomeJogador){
    var ss = SpreadsheetApp.getActive().getSheetByName("Dados");
    let ultJogador = ss.getLastColumn();
    var nomes = [(parseInt(ultJogador-5))];
    for(let j=5; j<ultJogador+1; j++){
      var nome = ss.getRange(1,j).getValue();
      nomes[j-5] = nome;
    }
    let posicao = 0;
    var nomeJogadorMaiusculo = nomeJogador.toUpperCase();
    for(let i=0; i<nomes.length; i++){
      if(nomes[i] == (nomeJogadorMaiusculo)){
        posicao = i+5;
      }
    }
    return posicao;
  }
  
  
  function existeConvidado(lista, numeroJogadores){
    var listaSeparada = lista.split(',');
    var nomes = [parseInt(numeroJogadores)];
    var contador = 0;
    let jogadores = 0;
    for(let i=0; i<listaSeparada.length; i++){
      var lista2 = listaSeparada[i].split("-");
      for(let j=0; j<lista2.length; j++){
        nomes[j+i] = lista2[j];
      }
    }
    for(let k=0; k<nomes.length; k++){
      if(verificaJogador(nomes[k])){
        jogadores = jogadores + 1;
      }                                                        
    }
    if(jogadores==(nomes.length)){
      return false;
    } else{
      return true;
    }       
  }
  
  function calculaPontos() {
    var ss = SpreadsheetApp.getActive().getSheetByName("Adicionar Jogos");
    var jogadores1 = ss.getRange('A12').getValue();
    var jogadores = "" + jogadores1;
    var nomeJogo = ss.getRange('A4').getValue();
    var tempoCaixa = (ss.getRange('A6').getValue())+":"+(ss.getRange('C6').getValue());
    var tempoJogo = (ss.getRange('A8').getValue())+":"+(ss.getRange('C8').getValue());
    var numeroJogadores = ss.getRange('A10').getValue();
  
    if(jogadores ==''||nomeJogo==''||tempoCaixa==''||tempoJogo==''||numeroJogadores==''){
      SpreadsheetApp.getUi().alert("ERRO: algum dos campos não foi preenchido")
      return;
    }
  
    let tempoCaixaMinutos = 60*(parseInt(ss.getRange('A6').getValue())) + (parseInt(ss.getRange('C6').getValue()));
    let tempoJogoMinutos = 60*(parseInt(ss.getRange('A8').getValue())) + (parseInt(ss.getRange('C8').getValue()));
    let tempoMedia = (parseInt(tempoCaixaMinutos) + parseInt(tempoJogoMinutos))/2;
    
  
    var sslista = SpreadsheetApp.getActive().getSheetByName('Dados');
    var ssJogadores = SpreadsheetApp.getActive().getSheetByName('Jogadores');
    var nomesJogadores = sslista.getRange(1,5,1,18).getValues();
  
    var ssConvidados = SpreadsheetApp.getActive().getSheetByName('Convidados');
    var ssHoras = SpreadsheetApp.getActive().getSheetByName('Horas Adicionadas');
    let ultJogoH = ssHoras.getLastRow();
  
    let maiorHora = maisHoras();
  
    var jogadoresOrdem = jogadores.split(',');
    var ultJogo = sslista.getLastRow(); 
    sslista.getRange(ultJogo+1,1).setValue(nomeJogo);
    sslista.getRange(ultJogo+1,2).setValue(numeroJogadores);
    sslista.getRange(ultJogo+1,3).setValue(tempoCaixa);
    sslista.getRange(ultJogo+1,4).setValue(tempoJogo);
    ssHoras.getRange(ultJogoH+1, 1).setValue(nomeJogo);
  
    if(existeConvidado(jogadores, numeroJogadores)){
      var ultJogoConvidado = ssConvidados.getLastRow();
      ssConvidados.getRange(ultJogoConvidado+1,1).setValue(nomeJogo);
      ssConvidados.getRange(ultJogoConvidado+1,2).setValue(numeroJogadores);
      ssConvidados.getRange(ultJogoConvidado+1,3).setValue(tempoCaixa);
      ssConvidados.getRange(ultJogoConvidado+1,4).setValue(tempoJogo);    
    }
    
    let empatesTotal = 0;
    for(let i=0; i<jogadoresOrdem.length; i++){
      var nomeAtual = jogadoresOrdem[i];
        if(nomeAtual.includes('-')){
          var nomesEmpates = nomeAtual.split('-');
          for(let j=0; j<nomesEmpates.length; j++){
            var nomeE = nomesEmpates[j];
            if(verificaJogador(nomeE)){
              let minutos = parseInt(quantasHorasTem(nomeE));
              let tempoJog = tempoMedia;
              let diferenca = maiorHora - minutos;
              if(diferenca>=600){
                if(diferenca<1200){
                  tempoJog = tempoJog * 2;
                  if(tempoJog>240){
                    tempoJog = 240;
                  }
                }else{
                  tempoJog = Math.ceil((tempoJog*(2.5)));
                  if(tempoJog>240){
                    tempoJog = 240;
                  }
                }
              }else if(diferenca < 300){
                tempoJog = tempoJog;
              }else{
                tempoJog = Math.ceil(tempoJog*(1.5));
                if(tempoJog>240){
                  tempoJog = 240;
                }
              }
              tempoJog = 5*(Math.ceil(tempoJog/5));
              let somaEmp = 0;
              for(let k=0; k<nomesEmpates.length; k++){
                let posicao = k+i+1+empatesTotal;
                let pont = buscarFaixa(tempoJog, numeroJogadores, posicao);
                somaEmp = somaEmp + parseInt(pont);
              }
              let pontuacao = Math.floor(somaEmp/(nomesEmpates.length));
              if(pontuacao<1){
                  pontuacao = 1;
              }
              let coluna = procuraPosicaoDados(nomeE);
              adicionaDados(ultJogo+1, coluna, parseInt(pontuacao));
              ssHoras.getRange(ultJogoH+1, coluna-3).setValue(tempoJog);
              
            }else{
              if(temConvidado(nomeE)){
                let somaEmp = 0;
                for(let l=0; l<nomesEmpates.length; l++){
                  tempoX = 5*(Math.ceil(tempoMedia/5));
                  let posicao = l+i+empatesTotal +1;
                  let pont = buscarFaixa(tempoX, numeroJogadores, posicao);
                  somaEmp = somaEmp + pont;
                }
                let pontuacao = Math.floor(somaEmp/(nomesEmpates.length));
                if(pontuacao<1){
                  pontuacao = 1;
                }
                let coluna = procuraPosicaoConvidado(nomeE);
                adicionaConvidados(ultJogoConvidado+1, coluna, pontuacao);
              }else{
                var ultConvidado = ssConvidados.getLastColumn();
                ssConvidados.insertColumnAfter(ultConvidado);
                var nomeM = nomeE.toUpperCase();
                ssConvidados.getRange(1, ultConvidado+1).setValue(nomeM);
                let somaEmp = 0;
                for(let m=0; m<nomesEmpates.length; m++){
                  let posicao = m+i+empatesTotal +1;
                  tempoX = 5*(Math.ceil(tempoMedia/5));
                  let pont = buscarFaixa(tempoX, numeroJogadores, posicao);
                  somaEmp = somaEmp + pont;
                }
                let pontuacao = Math.floor(somaEmp/(nomesEmpates.length));
                if(pontuacao<1){
                  pontuacao = 1;
                }
                let coluna = procuraPosicaoConvidado(nomeE);
                adicionaConvidados(ultJogoConvidado+1, coluna, pontuacao);              
              }
            }
          } 
          empatesTotal = empatesTotal + nomesEmpates.length - 1;       
        }else{
          if(verificaJogador(nomeAtual)){
            let minutos = parseInt(quantasHorasTem(nomeAtual));
            let tempoJog = tempoMedia;
            let diferenca = maiorHora - minutos;
            if(diferenca >= 600){
              if(diferenca<1200){
                  tempoJog = tempoJog * 2;
                  if(tempoJog>240){
                    tempoJog = 240;
                  }
                }else{
                  tempoJog = Math.ceil((tempoJog*(2.5)));
                  if(tempoJog>240){
                    tempoJog = 240;
                  }
                }
            }else if(diferenca < 360){
              tempoJog = tempoJog;
            }else{
              tempoJog = Math.ceil((tempoJog*(1.5)));
              if(tempoJog>240){
                tempoJog = 240;
              }
            }
            tempoJog = 5*(Math.ceil(tempoJog/5));
            let posicao = i+empatesTotal+1;
            let pontuacao = buscarFaixa(tempoJog, numeroJogadores, posicao);
            let coluna = procuraPosicaoDados(nomeAtual);
            adicionaDados(ultJogo+1, coluna, pontuacao);
            let tempoAnterior = ssJogadores.getRange(coluna-3, 3).getValue();
            ssHoras.getRange(ultJogoH+1, coluna-3).setValue(tempoJog);
          }else{
            if(temConvidado(nomeAtual)){
              let posicao = i+empatesTotal+1;
              tempoX = 5*(Math.ceil(tempoMedia/5));
              let pontuacao = buscarFaixa(tempoX, numeroJogadores,posicao);
              let coluna = procuraPosicaoConvidado(nomeAtual);
              adicionaConvidados(ultJogoConvidado+1, coluna, pontuacao);
            }else{
              var ultConvidado = ssConvidados.getLastColumn();
              ssConvidados.insertColumnAfter(ultConvidado);
              var nomeM = nomeAtual.toUpperCase(); 
              ssConvidados.getRange(1, ultConvidado+1).setValue(nomeM);
              let posicao = i+empatesTotal+1;
              tempoX = 5*(Math.ceil(tempoMedia/5));
              let pontuacao = buscarFaixa(tempoX, numeroJogadores,posicao);
              let coluna = procuraPosicaoConvidado(nomeAtual);
              adicionaConvidados(ultJogoConvidado+1, coluna, pontuacao);           
            }
          }
        }
    }
    var ssTempo = SpreadsheetApp.getActive().getSheetByName("Cronologia");
    let ultJogadorC = ssTempo.getLastRow();
    let ultJogoC = ssTempo.getLastColumn();
    ssTempo.insertColumnAfter(ultJogoC);
    ssTempo.getRange(1, ultJogoC+1).setValue(parseInt(ultJogoC));
    if(ultJogoC == 1){
      for(let b = 2; b<ultJogadorC+1; b++){
        let valorJogo = sslista.getRange(ultJogo+1, b+3).getValue();
        if(valorJogo == 0){
          ssTempo.getRange(b, 2).setValue(0);
        }else{
          ssTempo.getRange(b, 2).setValue(valorJogo);
        }
      }
    }
    for(let a = 2; a<ultJogadorC+1; a++){
      let valorJogo = sslista.getRange(ultJogo+1, a+3).getValue();
      let ultPont = parseInt(ssTempo.getRange(a, ultJogoC).getValue());
      if(valorJogo == null){
        ssTempo.getRange(a, ultJogoC+1).setValue(ultPont);
      }else{
        ssTempo.getRange(a, ultJogoC+1).setValue(ultPont + valorJogo);
      }
    }
    
    var ssMedia = SpreadsheetApp.getActive().getSheetByName("Médias");
    let ultJogoM = ssMedia.getLastColumn();
    let ultJogadorM = ssMedia.getLastRow();
    ssMedia.insertColumnAfter(ultJogoM);
    ssMedia.getRange(1, ultJogoM +1).setValue(parseInt(ultJogoM));
    for(let p = 2; p<ultJogadorM+1; p++){
      let media = ssJogadores.getRange(p,5).getValue();
      ssMedia.getRange(p, ultJogoM+1).setValue(media);
    }
  
    ss.getRange('A4').clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange("A6").clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange('A8').clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange('C8').clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange('C6').clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange('A10').clear({contentsOnly: true, skipFilteredRows: true});
    ss.getRange('A12').clear({contentsOnly: true, skipFilteredRows: true});  
  }
  
  function apagaJogo (){
    var ss = SpreadsheetApp.getActive().getSheetByName("Adicionar Jogos");
    var ssDados = SpreadsheetApp.getActive().getSheetByName("Dados");
    var ssTempo = SpreadsheetApp.getActive().getSheetByName("Horas Adicionadas");
    var ssCrono = SpreadsheetApp.getActive().getSheetByName("Cronologia");
    var ssMedia = SpreadsheetApp.getActive().getSheetByName("Médias");
  
    let num = parseInt(ss.getRange("H5").getValue());
    let ultD = ssDados.getLastRow();
  
    ssDados.deleteRows(num, 1);
    ssDados.insertRowAfter(ultD);
  
    ssTempo.deleteRows(num, 1);
    ssTempo.insertRowAfter(ultD);
  }
  
  function finalDeLiga (){
    var ss = SpreadsheetApp.getActive().getSheetByName("Adicionar Jogos");
    var ssJogadores = SpreadsheetApp.getActive().getSheetByName("Jogadores");
  
    var msgFinal = ss.getRange("H13").getValue().toUpperCase();
    if(msgFinal == "TERMINAR"){
      let ultJogador = ssJogadores.getLastRow();
      for(let i=2; i<ultJogador+1; i++){
        let minutos = ssJogadores.getRange(i, 3).getValue();
        let media = ssJogadores.getRange(i, 5).getValue();
        if(minutos >= 2220){
          let novaPontuacao = Math.floor((2400*media)/60);
          ssJogadores.getRange(i,3).setValue(2520);
          ssJogadores.getRange(i, 2).setValue(novaPontuacao); 
        }else if(minutos < 2220){
          let novoTempo = minutos + 180;
          let novaPontuacao = Math.floor((novoTempo * media)/60);
          ssJogadores.getRange(i, 3).setValue(novoTempo);
          ssJogadores.getRange(i, 2).setValue(novaPontuacao);
        }
      }
    }else{
      SpreadsheetApp.getUi().alert("ERRO: as instruções não foram realizadas para o final da liga!");
    }
  }