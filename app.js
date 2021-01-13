const APP_NAME = "Label Process"
const SN_LABEL_PASTE_HERE = "LABEL PASTE HERE"
// C > I, D > J, E > K, F > L, G > M, H > O, J > N
const SOURCE_FOLDER_ID = "1uhi5F3FZECo7chWpdfsvjUyMagPDPgX7"
const SN_ORDER_IMPORT = "OrderImport"

const START_COLUMN_INDEX = 8 // column I

function onOpen() {
    const ui = SpreadsheetApp.getUi()
    const menu = ui.createMenu(APP_NAME)
    menu.addItem("Get label data", "getLabelData")
    menu.addToUi()
}

const getLabelData = () => {
    const startTime = new Date().getTime()
    try {
        new App().getLabelData()
        const endTime = new Date().getTime()
        const usedTime = Math.floor((endTime - startTime) / 1000)
        SpreadsheetApp.getActive().toast(`Label data have been refreshed. Used time ${usedTime}s.`, APP_NAME)
    } catch (e) {
        console.log(e)
        SpreadsheetApp.getUi().alert(APP_NAME, e.message, SpreadsheetApp.getUi().ButtonSet.OK)
    }
}

class App {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.ui = SpreadsheetApp.getUi()
        this.sourceFolder = this.getFolderById(SOURCE_FOLDER_ID)
        this.wsLabelPasteHere = this.ss.getSheetByName(SN_LABEL_PASTE_HERE)
    }

    getFolderById(id) {
        try {
            return DriveApp.getFolderById(id)
        } catch (e) {
            return null
        }
    }

    isAppValid() {
        if (!this.sourceFolder) return { valid: false, message: `Invalid source folder id: "${SOURCE_FOLDER_ID}".` }
        if (!this.wsLabelPasteHere) return { valid: false, message: `We can't find the sheet "${SN_LABEL_PASTE_HERE}".` }
        const message = `"${APP_NAME}" APP is ready to go.`
        this.ss.toast(message, APP_NAME, 2)
        return { valid: true, message }
    }

    getFilesByType(mimeType) {
        this.ss.toast(`Searching for ${mimeType} files from source folder ...`, APP_NAME)
        const allFiles = this.sourceFolder.getFilesByType(mimeType)
        const files = []
        while (allFiles.hasNext()) {
            files.push(allFiles.next())
        }

        return files
    }

    getSourceDataFromSheet(sheet, ssName) {
        this.ss.toast(`Retriving data from ${ssName} ...`, APP_NAME)
        const data = {}
        const shop = ssName.trim().toUpperCase()
        const values = sheet.getDataRange().getDisplayValues()
        values.forEach(v => {
            const [order, orderTime, buyer, addressFirstLine, addressSecondLine, city, state, zip, fullAddress, country] = v
            const key = order.trim().toUpperCase() + shop
            data[key] = {
                shop,
                order,
                orderTime,
                buyer,
                addressFirstLine,
                addressSecondLine,
                city,
                state,
                zip,
                fullAddress,
                country,
            }
        })
        return data
    }

    getSourceData() {
        const files = this.getFilesByType(MimeType.GOOGLE_SHEETS)
        let sourceData = {}
        this.ss.toast(`Collecting label data from source data ...`, APP_NAME)
        files.forEach(file => {
            const id = file.getId()
            const ss = SpreadsheetApp.openById(id)
            const ssName = ss.getName()
            const sheet = ss.getSheetByName(SN_ORDER_IMPORT)
            if (sheet) {
                sourceData = { ...sourceData, ...this.getSourceDataFromSheet(sheet, ssName) }
            }
        })
        return sourceData
    }

    getLabelData() {
        const { valid, message } = this.isAppValid()
        if (!valid) {
            this.ui.alert(APP_NAME, message, this.ui.ButtonSet.OK)
            return
        }
        // App is ready to go
        const sourceData = this.getSourceData()
        const dataRange = this.wsLabelPasteHere.getDataRange()
        const values = dataRange.getDisplayValues()
        const newValues = []
        values.forEach((v, i) => {
            const [order, shop] = v
            newValues.push(v.slice(START_COLUMN_INDEX))
            if (i > 0) {
                newValues[i].fill(null)
                const key = order.trim().toUpperCase() + shop.trim().toUpperCase()
                const dataset = sourceData[key]
                if (dataset) {
                    newValues[i][0] = dataset.buyer
                    newValues[i][1] = dataset.addressFirstLine
                    newValues[i][2] = dataset.addressSecondLine
                    newValues[i][3] = dataset.city
                    newValues[i][4] = dataset.state
                    newValues[i][5] = dataset.country
                    newValues[i][6] = dataset.zip
                }
            }
        })
        this.wsLabelPasteHere.getRange(1, START_COLUMN_INDEX + 1, newValues.length, newValues[0].length).setValues(newValues)
    }
}