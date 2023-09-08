const { google } = require('googleapis');
const puppeteer = require('puppeteer');
const cheerio = require('cheerio');
const path = require('path');

(async () => {
  // Código Puppeteer para extrair os dados
  var valor = 0;
  var aux = [1, 6, 11, 16, 21];
  var mainUrl = ['http://10.20.196.53', 'http://10.20.196.54', 'http://10.20.198.117', 'http://10.20.197.244/web/guest/br/websys/webArch/getStatus.cgi', 'http://192.168.12.22'];
  
  do{
    try {
      // Inicializa o array data zerada cada vez que começa um loop
      var data = [];

      // Testa se é o ip da impressora da EMC que está em outro range de IP, caso seja inicia o browser com proxy para acessar
      if(mainUrl[valor] != 'http://192.168.12.22'){
        var browser = await puppeteer.launch({
          headless: 'new',
        });
      }else{
        var browser = await puppeteer.launch({
          headless: 'new',
          args: ['--proxy-server=10.20.198.198:3128'],
        });
      }
      
      var page = await browser.newPage();
      console.log("rodando pela: "+valor);
      await page.goto(mainUrl[valor]);
      
      // Aguarde o frame carregar
      if(mainUrl[valor] != 'http://10.20.197.244/web/guest/br/websys/webArch/getStatus.cgi'){
        await page.waitForSelector('frame[name="wlmframe"]');
        
        var frameElement = await page.$('frame[name="wlmframe"]');

        // Obtém o conteúdo do frame
        var frame = await frameElement.contentFrame();

        // Aguarda o iframe "toner" carregar dentro do frame "wlmframe"
        await frame.waitForSelector('iframe[name="toner"]');
        var iframeElement = await frame.$('iframe[name="toner"]');

        // Obtém o conteúdo do iframe "toner"
        var iframe = await iframeElement.contentFrame();
        var iframeContent = await iframe.content();

        // Usa o cheerio para carregar o conteúdo HTML do iframe "toner"
        var $ = cheerio.load(iframeContent);

        var tdElements = $('td[class="style363"]').slice(0, 8); 

        // Filtra os elementos para pegar apenas os valores de %, e os insere no array data
        tdElements.each((index, element) => {
          var resultado = index % 2;
          if(resultado == 1){
            var content = $(element).html();
            data.push([content]);
          }
        });

      }else{

        // Aguarda o carregamento da img com a classe bdr-1px-666
        await page.waitForSelector('img.bdr-1px-666');

        // Extrai a largura das imagens com a classe 'bdr-1px-666'
        var widths = await page.$$eval('img.bdr-1px-666', (imgs) => {
          var result = [];
          for (let i = 0; i < 4 && i < imgs.length; i++) {
            result.push(imgs[i].width);
          }
          return result;
        });

        // Realiza um calculo que utiliza valores contidos no HTML para dar o valor em % corretamente
        for(let i = 0; i < widths.length; i++){
          var x = (((widths[i]/160)*100))+"%";
          data.push([x]);
        }
        
      }

      // Finaliza o browser
      await browser.close();

      var keyFilePath = path.resolve('C:/Users/nilson.sena/Desktop/nodeScrap/planilha-impressoras-redeminas-9b57b374879d.json');
      // Configuração da API do Google Sheets
      var auth = new google.auth.GoogleAuth({
        keyFile: keyFilePath, // Arquivo de credenciais baixado
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      
      var sheets = google.sheets({ version: 'v4', auth });

      // ID da planilha do Google Sheets
      var spreadsheetId = '1bhe5pBh8Davt_DIE8gGeHuz9_5bopSlSQhNp086XR50';

      // Escrever os dados na planilha
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: ('Q'+aux[valor]), // Célula de início
        valueInputOption: 'RAW',
        resource: {
          values: data,
        },
      });

      console.log('Dados inseridos na planilha.');
    }catch (error){
      console.log('Erro ao conectar-se a impressora '+ mainUrl[valor]);
    }
    valor++;
    console.log(valor);
    console.log(mainUrl.length);
  }while(valor < mainUrl.length);
  process.exit(0);
})();

