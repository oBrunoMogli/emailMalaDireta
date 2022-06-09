function aoAbrir() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Enviar Email')
      .addItem('Enviar','enviar')
      .addToUi();
}

function enviar()
{
   var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
   var row = cell.getRow();
  
   var candidate = getCandidateFromRow(row); 
  
   var ui = SpreadsheetApp.getUi();
   var response = ui.alert('Enviar Email para '+candidate.name+'?', ui.ButtonSet.YES_NO);

   if(response == ui.Button.YES)
   {
     enviarEmail(row,candidate);
     SpreadsheetApp.getUi().alert("Email Enviado");

   }
}

function getCandidateFromRow(row)
{
  var values = SpreadsheetApp.getActiveSheet().getRange(row, 1,row,10).getValues();
  var rec = values[0];
  
  var candidate = 
      {
        name:         rec[0],
        emailAdress:  rec[1],
        emailAdress2: rec[2],
        newEmail:     rec[3],
        file:         rec[4],
        linkDocs:     rec[5],
        subject:      "TESTE 003",
        reply: "sistemas@educamericana.sp.gov.br"
      };
  
   //candidate.name = candidate.first_name+' '+candidate.last_name;
   
  
   return candidate;
}

function enviarEmail(row, candidate)
{
  
  //criando template html
  var templ = HtmlService
      .createTemplateFromFile('3corpoEmail');

  templ.candidate = candidate;
  
  //criando a variavel de mensagem HTML
  var message = templ.evaluate().getContent();

  // função para enviar os emails
  MailApp.sendEmail({
    to: candidate.emailAdress,
    to: candidate.emailAdress2,
    replyTo: candidate.reply,
    subject: candidate.subject,
    htmlBody: message
  });

  SpreadsheetApp.getActiveSheet().getRange(row, 9).setValue('ENVIADO');

}

/*
array 0 - col A = 1 -  nome_completo
array 1 - col B = 2 -  e-mail1
array 2 - col C = 3 -  e-mail2
array 3 - col D = 4 -  e-mail_novo
array 4 - col E = 5 -  Merged Doc ID - Informações e-mail @educamericana.sp.gov.br
array 5 - col F = 6 -  Merged Doc URL - Informações e-mail @educamericana.sp.gov.br
array 6 - col G = 7 -  Link to merged Doc - Informações e-mail @educamericana.sp.gov.br
array 7 - col H = 8 -  Document Merge Status - Informações e-mail @educamericana.sp.gov.br
array 8 - col I = 9 -  Em Branco/Vazia
array 9 - col J = 10 - Número de linhas a ser processada (inserir na linha 1)
*/