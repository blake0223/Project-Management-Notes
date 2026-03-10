function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Job Review Tools')
    .addItem('Begin Job Review', 'openReviewModal')
    .addItem('Add Post-Project Note to Job','openPostProjectNotesModal')
    .addItem('Refresh Jobs', 'loadJobTrackingIntoData')
    .addToUi();
}
