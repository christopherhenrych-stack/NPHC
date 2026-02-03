function doGet() {
    // This tells Google to look for an HTML file specifically named "Index"
  return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('NPHC Status Board')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
