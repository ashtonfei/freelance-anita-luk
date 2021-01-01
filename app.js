const CASE_PRINT_FOLDER_ID = "1jTefPVZDz7yIaRUNzN03DRMvwKx6lgv1"
const SN_MASTER = "master"
const SN_LIST = "list"
const APP_NAME = "MCO App"
const IMAGE_NOT_FOUND = "Image Not Found"


function getColumnIndex(name) {
    return SpreadsheetApp.getActive().getActiveSheet().getRange(`${name.trim()}1`).getColumn() - 1
}

function downloadImages() {
    try {
        const app = new OrderApp()
        return app.downloadImages()
    } catch (e) {
        SpreadsheetApp.getActive().toast(e.message, APP_NAME, 30)
    }
}

function addOrders({ id, sheetName }) {
    try {
        const app = new OrderApp()
        app.addOrders({ id, sheetName })
    } catch (e) {
        SpreadsheetApp.getActive().toast(e.message, APP_NAME, 30)
    }
}

function getReports() {
    const app = new OrderApp()
    const sheets = app.getSpreadsheets()
    return JSON.stringify(sheets)
}

function openApp() {
    const app = new OrderApp()
    app.openDialog()
}

function onEdit_(e) {
    const app = new ImageApp()
    app.onEdit(e)
}

function updateImages() {
    try {
        const app = new ImageApp()
        app.updateImages()
    } catch (e) {
        SpreadsheetApp.getActive().toast(e.message, APP_NAME, 30)
    }
}

function createTrigger() {
    const functionName = "onEdit_"
    const sheet = SpreadsheetApp.getActive()
    ScriptApp.newTrigger(functionName).forSpreadsheet(sheet).onEdit().create()
    onOpen()
}

function onOpen() {
    const ui = SpreadsheetApp.getUi()
    const menu = ui.createMenu(APP_NAME)
    menu.addItem("Open", "openApp")

    const functionName = "onEdit_"

    const triggers = ScriptApp.getProjectTriggers()
    const trigger = triggers.find(trigger => trigger.getHandlerFunction() === functionName)
    if (!trigger) {
        menu.addItem("Create trigger", "createTrigger")
    }

    menu.addItem("Update images", "updateImages")
    menu.addToUi()
}
