function doGet(e)
{
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ToolCS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
