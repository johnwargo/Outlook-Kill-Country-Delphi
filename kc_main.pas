{ ****************************************************************************
  Kill Country

  As I've switched from one outlook version to another and as I've
  synchronized my outlook contact list with multiple work devices
  (smartphones and tablets), the state of the contact information has become
  suspect. One thing I noticed repeatedly is that some contacts, but not all
  had a country value set (United States of America). Not that this is such
  a big deal except that in some cases, that's all that was in the address
  field. Besides, I'm a US resident, why would I care to know what country
  the contact is in? If country is blank, I'll assume it's the US.

  Anyway, with all of my work lately on integrating Delphi applications with
  Outlook, I decided to solve this particular problem this morning. I made
  this app which essentially whacks any country field if it contains United
  States of America.

  Is it the most efficient code? No, but it does work. I'd preferred to be
  able to pass the country field (there are 4 of them in a contact record) to
  a function, but there's no way that I could find to get an outlook contact
  item value by name.

  Be sure to backup your outlook.pst file before you start this process as
  it is not reversable.

  I only tested this on my local, stand-alone outlook instance. This process
  should work when Outlook is connected to a Microsoft Exchange server, but
  that's not something I tested.

  John M. Wargo
  October 6, 2015
  www.johnwargo.com
  **************************************************************************** }
unit kc_main;

interface

uses
  ComObj, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,

  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls,
  Vcl.StdCtrls;

type
  TfrmMain = class(TForm)
    StatusBar1: TStatusBar;
    output: TMemo;
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

procedure TfrmMain.FormActivate(Sender: TObject);
Const
  // Outlook contact item
  olContactItem = $00000002;
  blankStr = '';
  countryStr = 'United States of America';

var
  outlook, oiItem, ns, folder: OLEVariant;
  i, intFolderType, numItems: Integer;
  msgText, fullName, otherCountry, mailingCountry, workCountry,
    homeCountry: String;

  { to find a default Contact folder }
  function GetOutlookFolder(folder: OLEVariant): OLEVariant;
  var
    i: Integer;
  begin
    for i := 1 to folder.Count do
    begin
      if (folder.Item[i].DefaultItemType = olContactItem) then
        result := folder.Item[i]
      else
        result := GetOutlookFolder(folder.Item[i].Folders);
      if not VarIsNull(result) and not VarIsEmpty(result) then
        break;
    end;
  end;

begin
  output.Lines.add('Creating Outlook Object');
  // initialize a connection to Outlook
  outlook := CreateOLEObject('Outlook.Application');
  // get the MAPI namespace
  ns := outlook.GetNamespace('MAPI');
  // get a default Contacts folder
  folder := GetOutlookFolder(ns.Folders);
  output.Lines.add('Default contact folder: ' + string(folder));
  // if  Calendar folder is found
  if VarIsNull(folder) and not VarIsEmpty(folder) then
  begin
    // Then tell the user
    msgText := 'Unable to determine the default contact folder';
    ShowMessage(msgText);
    output.Lines.add(msgText);
  end else begin
    // Process entries in the folder
    output.Lines.add(Format('Searching "%s" folder', [folder]));
    intFolderType := folder.DefaultItemType;
    numItems := folder.Items.Count;
    if (numItems > 0) then
    begin
      output.Lines.add(Format('Found %d items', [numItems]));
      // Process the list of contacts
      for i := 1 to numItems do
      begin
        // Make sure Windows gets a chance to do its stuff
        Application.ProcessMessages;
        // Get the nth outlook contact item
        oiItem := folder.Items[i];
        // get the contact's full name
        fullName := oiItem.fullName;
        // Check to see if we have a full name
        if (Length(fullName) < 1) then
          // If not, use the company name
          fullName := oiItem.companyName;
        // =========================================
        // Process the home address country field
        // =========================================
        homeCountry := oiItem.HomeAddressCountry;
        if (homeCountry = countryStr) then
        begin
          output.Lines.add(Format('%s: Deleting home country (%d)',
            [fullName, i]));
          oiItem.HomeAddressCountry := blankStr;
          // Could I have checked to see if an item needed saving and only
          // do it once? Yes.
          oiItem.Save;
        end;
        // =========================================
        // Process the work address country field
        // =========================================
        workCountry := oiItem.BusinessAddressCountry;
        if (workCountry = countryStr) then
        begin
          output.Lines.add(Format('%s: Deleting work country (%d)',
            [fullName, i]));
          oiItem.BusinessAddressCountry := blankStr;
          oiItem.Save;
        end;
        // =========================================
        // Process the other address country field
        // =========================================
        otherCountry := oiItem.OtherAddressCountry;
        if (otherCountry = countryStr) then
        begin
          output.Lines.add(Format('%s: Deleting other country (%d)',
            [fullName, i]));
          oiItem.OtherAddressCountry := blankStr;
          oiItem.Save;
        end;
        // =========================================
        // Process the mailing address country field
        // =========================================
        mailingCountry := oiItem.MailingAddressCountry;
        if (mailingCountry = countryStr) then
        begin
          output.Lines.add(Format('%s: Deleting mailing country (%d)',
            [fullName, i]));
          oiItem.MailingAddressCountry := blankStr;
          oiItem.Save;
        end;
        // =========================================
      end;
      output.Lines.add('All done!');
    end else begin
      output.Lines.add('No contact entries found');
    end;
  end;
end;

end.
