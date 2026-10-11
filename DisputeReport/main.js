function doGet(e)
{
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dispute Report')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
