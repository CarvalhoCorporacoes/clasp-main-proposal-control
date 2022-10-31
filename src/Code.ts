import { CreateNewContractUseCase } from "./domain/application/create-new-contract.use-case"

const CONTRACT_TEMPLATE_ID = '1YgVlTSshqM9esd3SHQpfrYP9Wqy40q_ypwdOw8RhS40'
const PAYMENT_CONTROL_ID = '1iPVnFImzzSADEcAWWWyboauqbgxzjLwOWpRBtk_BaDk'

function onOpen () {

  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('Contratos')
  menu.addItem('Gerar Novo Contrato', 'createNewContract')
  menu.addToUi()

}

function createNewContract () {
  const spreadsheet = SpreadsheetApp.getActiveSheet()

  const createNewContractUSeCase = new CreateNewContractUseCase(spreadsheet, CONTRACT_TEMPLATE_ID, PAYMENT_CONTROL_ID)
  try {
    createNewContractUSeCase.execute()
  } catch (error) {
    Logger.log(`Deu Ruim, Chame o Suport: ${error}`)
  }
}