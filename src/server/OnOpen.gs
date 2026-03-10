function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Job Review Tools')
    .addItem('Begin Job Review', 'openReviewModal')
    .addItem('Add Post-Project Note to Job','openPostProjectNotesModal')
    .addItem('Manual Job Entry', 'openManualJobEntryModal')
    .addSeparator()
    .addItem('Refresh Jobs', 'loadJobTrackingIntoData')
    .addToUi();
}
