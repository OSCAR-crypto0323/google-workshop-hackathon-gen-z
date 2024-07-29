function myFunction() {
  try {
    
    let interview = SpreadsheetApp.openById('1uYuEydXV0I86Vt5M6D1tLVoNhpwNwBQ8MIWmuPmoTjY').getSheetByName('interview');
    let lastrow = interview.getLastRow();

    for (let i = 2; i <= lastrow; i++) { 
      let name = interview.getRange(i, 3).getValue();
      let email = interview.getRange(i, 10).getValue();
      let firstinterview = interview.getRange(i, 23).getValue();
      let lastinterview = interview.getRange(i, 24).getValue();
      let firstEmailSent = interview.getRange(i, 26).getValue(); 
      let lastEmailSent = interview.getRange(i, 27).getValue(); 

      
      if (firstEmailSent !== 'Sent') {
        if (firstinterview === 'Yes') {
          let firstapproveletter = DocumentApp.openById('1jAywF5-IEJViifaoKANJm53SJc_GnLz4sa0gouwFKl0').getBody().getText();
          GmailApp.sendEmail(email, 'JOB APPROVE LETTER', firstapproveletter);
          interview.getRange(i, 26).setValue('Sent');
        } else if (firstinterview === 'No') {
          let rejectletter = DocumentApp.openById('1GjDnNlJnmKPD0wTiIIbhvxE2wBLd8NY_Ha5MYo_6gqY').getBody().getText();
          GmailApp.sendEmail(email, 'JOB REJECTION LETTER', rejectletter);
          interview.getRange(i, 26).setValue('Sent'); 
        }
      }

      
      if (firstEmailSent === 'Sent' && lastinterview && lastEmailSent !== 'Sent') { 
        if (lastinterview === 'Yes') {
          let approveletter = DocumentApp.openById('1fABFUua_c140IutZXaZDchJbUtrnmqOjQqgSwD4ufo0').getBody().getText();
          GmailApp.sendEmail(email, 'FINAL JOB APPROVE LETTER', approveletter);
          interview.getRange(i, 27).setValue('Sent'); 
        } else if (lastinterview === 'No') {
          let rejectletter = DocumentApp.openById('1GjDnNlJnmKPD0wTiIIbhvxE2wBLd8NY_Ha5MYo_6gqY').getBody().getText();
          GmailApp.sendEmail(email, 'FINAL JOB REJECTION LETTER', rejectletter);
          interview.getRange(i, 27).setValue('Sent');
        }
      }

     
      console.log(`Emails sent to: ${email}`);
    }
  } catch (e) {
    
    console.error(`Error: ${e.message}`);
  }
}
