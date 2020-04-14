/**
** Quote Generator v0.4
**
**
*/
var startRow;
var endRow;
var localPrice;

function quoteRequest() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  startRow = showPrompt("Start Row", "Enter the number of the starting row");
  endRow = showPrompt("End Row", "Enter the number of the ending row");
  getInfo(startRow,endRow);
}

function getInfo(sRow, eRow)
{
  var row;
  for(var i = 0; i <= (eRow-sRow); i++)
  {
    row=Number(sRow)+i;
    createQuote(row);
  }
}

function createQuote(row) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Form Responses 1');
  var templateID = '0a5YL2HcFAf91uE7C5dwjLnMqsVTXmM7Fivgo';
  
  var entries = Sheets.Spreadsheets.Values.get('1nFprH9WAlRf0vioPRvg4kPfF1zQKGL4sENX6XF1z8FQ', 'A' + row + ':AK' + row);
  
  for(var i = 0; i < entries.values.length; i++){
    
    //Logger.log(entries);
    var data = entries.values[i];
    
    //Make a copy of the template file
    var documentId = DriveApp.getFileById(templateID).makeCopy().getId();
    
    //Rename the copied file
    DriveApp.getFileById(documentId).setName('Quote #' + data[2] + ' ' + data[3] + ' ' + data[4]);
    
    //Get the document body as a variable
    var body = DocumentApp.openById(documentId).getBody();
    
    localPrice = showAlert("Quote #" + data[2] + " - Local Price", "Use local pricing?"
    + "\r\nCity: " + data[11]
    + "\r\nGuests: " + data[13]);
    
    //Insert the data
    body.replaceText('<<Quote #>>', data[2]);
    var formattedDate = Utilities.formatDate(new Date(data[0]), "EST", "MM/dd/yyyy");
    body.replaceText('<<Timestamp>>', formattedDate);
    body.replaceText('<<Name>>', data[3]);
    body.replaceText('<<Last Name>>', data[4]);
    body.replaceText('<<Telephone>>', data[5]);
    body.replaceText('<<Email>>', data[6]);
    body.replaceText('<<Date>>', data[7]);
    body.replaceText('<<Approximate Time Event Starts>>', data[8]);
    body.replaceText('<<Address>>', data[10]);
    body.replaceText('<<City>>', data[11]);
    body.replaceText('<<Zip>>', data[12]);
    body.replaceText('<<Approximate Number of Guests>>', Number(data[13]));
    body.replaceText('<<Type of Event>>', data[9]);
    
    //prices master list
    var msvLocal=5.00;
    var msvNonLocal=10.00;
    var chefp=5.00;
    var tablewarep=5.00;
    var salestax=5.00;

    //paella
    var mixta=0;
    var subtotal=0;
    var total=0;

    if(data[31] == 'Yes' || data[27] == 'Yes' || data[30] == 'Yes')
    {
    //mixta
    if(localPrice == true && data[31] == 'Yes')
    {
      mixta = Number((showPrompt("Paella Mixta", "Enter custom quantity below. Max = " + Number(data[13]).toFixed(2) 
      + "\r\nMixta=" + data[31]
      + "\r\nSeafood=" + data[27]
      + "\r\nValencian=" + data[30]
      + "\r\nLobster=" + data[28]
      + "\r\nVegan=" + data[29]
      + "\r\nClassic=" + data[32])));
      body.replaceText('#mixtaq#', Number(mixta).toFixed());
      mixta=Number(msvLocal*mixta);
      subtotal+=mixta;
      body.replaceText('<<PAELLA MIXTA: >>', '$'+mixta.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'));
      
    }
    else if(localPrice == false && data[31] == 'Yes')
    {
      mixta=Number((showPrompt("Paella Mixta", "Enter custom quantity below. Max = " + Number(data[13]).toFixed(2) 
      + "\r\nMixta=" + data[31]
      + "\r\nSeafood=" + data[27]
      + "\r\nValencian=" + data[30]
      + "\r\nLobster=" + data[28]
      + "\r\nVegan=" + data[29]
      + "\r\nClassic=" + data[32])));
      body.replaceText('#mixtaq#', Number(mixta).toFixed());
      mixta=Number(msvNonLocal*mixta);
      subtotal+=mixta;
      body.replaceText('<<PAELLA MIXTA: >>', '$'+mixta.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'));
    }
    else{ 
      body.replaceText('<<PAELLA MIXTA: >>', '');
      body.replaceText('Paella Mixta', '');
      body.replaceText('#mixtaq#', '');
      body.replaceText('^Pearl(.+?)paella$', '');
      mixta=0;
    }
    
    body.replaceText('#subtotal#', Number(subtotal).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'));
    
    //cleanup whitespace
    //body.replaceText("\\v+", "")
    
    //chef
    var chef=0;
    var total=subtotal;
    if(localPrice == true && data[15] == 'Cooked on-site')
    {
      chef=Number(chefp)+Number(data[13]);
      body.replaceText('<<Paella cooked fee>>', '$' + Number(chef).toFixed(2));
    }
    else if(localPrice == false && data[15] == 'Cooked on-site')
    {
      chef=Number(chefp) + (2*(Number(data[13])));
      body.replaceText('<<Paella cooked fee>>', '$' + Number(chef).toFixed(2));
    }
    else{ body.replaceText('<<Paella cooked fee>>', ''); chef=0; }
    
    //all-inclusive discount
    var discountamount=0;
    var sub2discount=0;
    if (data[17] != 'No Salad' && data[16] == 'Yes' && data[22] != 'No' && data[18] != 'No Dessert')
    {
      body.replaceText('^Promotional(.+?)[Salad)*]$', 'All-Inclusive package discount - 10%');
        body.replaceText('^#discount#(.+?)[Salad)*]$', '10% Off (Paella + Salad + Traditional Tapas + Sangria + Dessert)');
      discountamount=Number((10/100)*(mixta+seafood+valencian+lobster+vegan+classic+salad+tapas+sangria+dessert));
      body.replaceText('#discountamount#', '- $' + discountamount.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,') + '');
      sub2discount=Number(subtotal-discountamount);
      body.replaceText('#sub2discount#', sub2discount.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,')); 
      subtotal=sub2discount;
      total=sub2discount;
    }
    //promo discount
    else
    {
    var discount=Number((showPrompt("Quote #" + data[2] + " - Discount", "Enter the appropriate discount percentage for " + Number(data[13]) 
      + " people " + data[15] + "\r\nEnter 0 for no discount")));
    if(discount != 0)
    { 
      body.replaceText('#discount#', discount + '%');
      discountamount=Number((discount/100)*(mixta+seafood+valencian+lobster+vegan+classic+salad));
      body.replaceText('#discountamount#', '- $' + discountamount.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,') + ''); 
      sub2discount=Number(subtotal-discountamount);
      body.replaceText('#sub2discount#', sub2discount.toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,')); 
      subtotal=sub2discount;
      total=sub2discount;
    }
    else
    { 
      body.replaceText('^Promotional(?s)(.+?)Salad[)*]', '');
      body.replaceText('^#discount#(?s)(.+?)Salad[)*]', '');
      body.replaceText('#discountamount#', '');
      body.replaceText('#sub2discount#', Number(subtotal).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,')); 
    }
    }
    
    body.replaceText('<<Date_erd>>', data[7]);
    var tip=Number(((19-salestax)/100)*(total+discountamount+chef));
    var tax=Number((salestax/100)*(subtotal+discountamount));
    var servicecharge=tip+tax;
    body.replaceText('#servicech#', '+ ($' + Number(servicecharge).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,') + ')');
    
    if(data[25] == '')
    {
       body.replaceText('<<Additional Comments>>', '');
    }
    else{ body.replaceText('<<Additional Comments>>', data[25]); }
    //tableware
    var tableware=0;
    if(data[35] == 'Compostable Plates, Forks, and Napkins (FREE)')
    {
       body.replaceText('<<Tableware>>', 'Compostable Plates, Forks, and Napkins (eco-friendly)');
       tableware=0;
       body.replaceText('#tablecalc#', '$0.00');
    }
    else if(data[35] == 'Porcelain Plates and Stainless Steel Forks ($5.50/guest)')
    {
       body.replaceText('<<Tableware>>', 'Porcelain Plates and Stainless Steel Forks ($5.50/guest)');
       tableware=Number(Number(tablewarep)*Number(data[13])).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,');
       body.replaceText('#tablecalc#', '$' + tableware);
    }
    else
    {
       body.replaceText('Tablewares', '');
       body.replaceText('<<Tableware>>', '');
       body.replaceText('#tablecalc#', '');
       tableware=0;
    }    
    body.replaceText('#subtotalsh#', Number(subtotal+chef+Number(tableware)).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'));
    body.replaceText('#total#', Number(total+chef+Number(tableware)).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'));
    
    //require a deposit if subtotal >=$800
    if(Number(total) >= 800)
    {
      body.replaceText('^To make(.+?)event.$', '');
      body.replaceText('^Email or(.+?)reservation$', '');
    }
    if(data[15] == 'Cooked on-site')
    {
      body.replaceText('^To make(.+?)event.$', '');
      body.replaceText('^Email or(.+?)reservation$', '');
      body.replaceText('#deposits#', 'Chef Fee + 10% Deposit');
      body.replaceText('#10percent#', Number(chef+(0.10*subtotal)).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'));
    }
    else
    {
      body.replaceText('^Cooked(.+?)Fee$', '');
      body.replaceText('^Chef(.+?)only$', '');
      body.replaceText('<<Paella cooked fee>>', '');
      body.replaceText('#deposits#', '10% Deposit');
      body.replaceText('#10percent#', Number(0.10*(subtotal)).toFixed(2).replace(/(\d)(?=(\d{3})+\.)/g, '$1,'));
    }

    //final cleanup
    body=removeBlankRows(body);
    
    var folderName=showPrompt("Quote #" + data[2] + " - Destination Folder", 'Enter the destination folder name (4300s, 4400s, etc).\r\nNote: This folder must already exist');
    moveFiles(documentId, folderName);
    
  }

}

//https://www.practicalecommerce.com/create-google-docs-google-sheet
//https://developers.google.com/apps-script/guides/sheets

//https://stackoverflow.com/questions/38808875
function moveFiles(sourceFileId, targetFolderName) {
  var file = DriveApp.getFileById(sourceFileId);
  file.getParents().next().removeFile(file);
  DriveApp.getFoldersByName(targetFolderName).next().addFile(file);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Quote Request')
      .addItem('Generate Quote...', 'quoteRequest')
      //.addSeparator()
      //.addSubMenu(ui.createMenu('Sub-menu')
      //    .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function showPrompt(title, info) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      title,
      info,
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  return text;
}

function showAlert(title, info) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     title,
     info,
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    //ui.alert('Confirmation received.');
    return true;
  } else {
    // User clicked "No" or X in the title bar.
    //ui.alert('Permission denied.');
    return false;
  }
}

//https://ctrlq.org/code/20527-delete-blank-rows-google-document
function removeBlankRows(bodyId) {

    //var document = docId ?
        //DocumentApp.openById(docId) :
        //DocumentApp.getActiveDocument();

    var body = bodyId;
    var search = null;
    var tables = [];

    // Extract all the tables inside the Google Document
    while (search = body.findElement(DocumentApp.ElementType.TABLE, search)) {
        tables.push(search.getElement().asTable());
    }

    tables.forEach(function (table) {
        var rows = table.getNumRows();
        // Iterate through each row of the table
        for (var r = rows - 1; r >= 0; r--) {
            // If the table row contains no text, delete it
            if (table.getRow(r).getText().replace(/\s/g, "") === "") {
                table.removeRow(r);
            }
        }
    });

    return body;
}