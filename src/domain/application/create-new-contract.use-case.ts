import { getLastRowSpecial, getNamedOptions, pad } from "../../utils"

export class CreateNewContractUseCase {
  readonly contractTemplate: GoogleAppsScript.Drive.File
  readonly newContract: GoogleAppsScript.Drive.File
  readonly paymentControlSS: GoogleAppsScript.Spreadsheet.Spreadsheet
  readonly paymentControlSheet: GoogleAppsScript.Spreadsheet.Sheet
  customerFolder: GoogleAppsScript.Drive.Folder | undefined
  constructor (
    private readonly spreadsheet: GoogleAppsScript.Spreadsheet.Sheet,
    private readonly contractTemplateID: string,
    private readonly paymentControlID: string,
  ) {
    this.contractTemplate = DriveApp.getFileById(this.contractTemplateID)
    this.newContract = this.contractTemplate.makeCopy()
    this.paymentControlSS = SpreadsheetApp.openById(this.paymentControlID)
    const proposalControlSheet = this.paymentControlSS.getSheetByName('Propostas')
    if (!proposalControlSheet || proposalControlSheet == null){
      throw new Error('Proposal control sheet not found.')
    }
    this.paymentControlSheet = proposalControlSheet
  }

  execute (): void {
    this.createNewContract()
    this.updateStatus()
  }

  private createNewContract (): void {
    const activeCell = this.spreadsheet.getActiveRange()!
    const range = this.spreadsheet.getRange(activeCell.getRowIndex(), 1, 1, 9)
    const folderCell = range.getCell(1,6)
    const dimensioningCell = range.getCell(1,7)
    const proposalCell = range.getCell(1,8)
    const folderCellUrl = folderCell.getRichTextValue()!.getLinkUrl()!
    const proposalUrl = proposalCell.getRichTextValue()!.getLinkUrl()!
    
    const folders = DriveApp.getFolders()

    const values = this.spreadsheet.getRange(activeCell.getRowIndex(), 1, 1, 9).getValues()![0]
    const customerName = values[1]
    const dimensioningName = values[6]
    const proposalName = values[7]

    const date = new Date(Date.now())

    const month = date.getMonth()
    const year = date.getFullYear()

    const folderDate = `${month}/${year}`
    
    this.customerFolder = this.getCustomerFolder(folders, folderCellUrl)
    if (!this.customerFolder || this.customerFolder == undefined) {
      throw new Error('Customer folder not found.')
    }
    const filesInDestFld = this.customerFolder.getFiles()

    let proposalNumber = 1
    
    for ( ; filesInDestFld.hasNext(); ) {
      const file = filesInDestFld.next()
      if (file) {
        const fileName = file.getName().match('Contrato')
        if (fileName != null) {
          proposalNumber += 1
        }
      }
    }
    const templateName = this.contractTemplate.getName().match('Template')

    if (templateName == null ) {
      proposalNumber -=1
    }

    const newContractName = `${pad(proposalNumber, 2)} - ${folderDate} - ${customerName} - Contrato`
    this.newContract.setName(newContractName)
    this.newContract.moveTo(this.customerFolder)
    const newContractID = this.newContract.getId()
    const newContractUrl = this.newContract.getUrl()

    const columnToCheck = this.paymentControlSheet.getRange("A:A").getValues()

    const lastRow = getLastRowSpecial(columnToCheck)
    const lastColumn = this.paymentControlSheet.getLastColumn()

    Logger.log(`Last Row: ${lastRow}`)

    //EXAMPLE: Get the data range based on our selected columns range.
    const rangeToSetData = this.paymentControlSheet.getRange(lastRow + 1, 1, 1, lastColumn)
    const paymentDimensioningCell = rangeToSetData.getCell(1,2)
    const paymentFolderCell = rangeToSetData.getCell(1,8)
    const paymentProposalCell = rangeToSetData.getCell(1,9)
    const paymentContractCell = rangeToSetData.getCell(1,10)

    const dimensioningUrl = dimensioningCell.getRichTextValue()!.getLinkUrl()!

    //data To fetch in table
    const valuesTosSet = [
      [
        newContractID,
        dimensioningName, // Proposal name should add link
        customerName,
        '',
        '',
        '',
        '',
        'Pasta', 
        proposalName, // PDf Data, could be get by 
        this.newContract.getName(),
        date.toLocaleDateString('pt-BR')
      ]
    ]
    rangeToSetData.setValues(valuesTosSet)

    const dimensioningPathRichValue = SpreadsheetApp.newRichTextValue()
      .setText(dimensioningName)
      .setLinkUrl(dimensioningUrl.toString())
      .build();
    paymentDimensioningCell.setRichTextValue(dimensioningPathRichValue)

    const folderPathRichValue = SpreadsheetApp.newRichTextValue()
      .setText('Pasta')
      .setLinkUrl(folderCellUrl.toString())
      .build();
    paymentFolderCell.setRichTextValue(folderPathRichValue)
    
    const proposalPathRichValue = SpreadsheetApp.newRichTextValue()
      .setText(proposalName)
      .setLinkUrl(proposalUrl.toString())
      .build();
    paymentProposalCell.setRichTextValue(proposalPathRichValue)

    const contractPathRichValue = SpreadsheetApp.newRichTextValue()
      .setText(newContractName)
      .setLinkUrl(newContractUrl.toString())
      .build();
    paymentContractCell.setRichTextValue(contractPathRichValue)
  }

  private updateStatus (): void {
    const SHEET_NAME = 'Propostas'
    const ESTATE_NAME = 'Estatdos/Status'
    const PROPSAL_WS = this.paymentControlSS.getSheetByName(SHEET_NAME)!
    const ESTATE_WS = this.paymentControlSS.getSheetByName(ESTATE_NAME)!

    //OPTIONS
    const OPTIONS = ESTATE_WS.getNamedRanges()
    const stateOptions = getNamedOptions(OPTIONS, 'States').filter((type: string) => type != '')
    const paymentOptions = getNamedOptions(OPTIONS, 'Payments_Options').filter((type: string) => type != '')
    const pixStatusOptions = getNamedOptions(OPTIONS, 'Pix_Status').filter((type: string) => type != '')

    //Select the column we will check for the first blank cell
    const columnToCheck = PROPSAL_WS.getRange("A:A").getValues()
    
    // Get the last row based on the data range of a single column.
    const lastRow = getLastRowSpecial(columnToCheck)
    //Set State
    const stateField = PROPSAL_WS.getRange(lastRow, 4)
    this.applyValidationToCell(stateOptions, stateField)
    //Set Payments
    const paymentField = PROPSAL_WS.getRange(lastRow, 5)
    this.applyValidationToCell(paymentOptions, paymentField)
    //Set Status
    const statusField = PROPSAL_WS.getRange(lastRow, 7)
    this.applyValidationToCell(pixStatusOptions, statusField)

    this.paymentControlSheet.getFilter()!.sort(11, false)
  }

  private getCustomerFolder (folders: GoogleAppsScript.Drive.FolderIterator, folderCellUrl: string): GoogleAppsScript.Drive.Folder | undefined {

    for (; folders.hasNext();) {
      const folder = folders.next()
      const folderURL = folder.getUrl()
      if (folderURL == folderCellUrl) {
        return folder
      }
    }
  }

  private getPDFFile (customerFolder: GoogleAppsScript.Drive.Folder, pdfUrl: string): GoogleAppsScript.Drive.File | undefined  {
    const files = customerFolder.getFiles()

    for (; files.hasNext();) {
      const file = files.next()
      if(file.getUrl() == pdfUrl) {
        return file
      }
    }
    return 
  }

  private applyValidationToCell (list: string[], cell: GoogleAppsScript.Spreadsheet.Range) {
    const rule = SpreadsheetApp
      .newDataValidation()
      .requireValueInList(list)
      .setAllowInvalid(false)
      .build()
  
      cell.setDataValidation(rule).setValue(list[0])
  }
}