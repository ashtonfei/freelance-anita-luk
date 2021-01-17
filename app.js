const APP_NAME = "Label Process"
const SN_LABEL_PASTE_HERE = "LABEL PASTE HERE"
const SN_USPS = "USPS"
const SN_SHIPMENTS = "SHIPMENTS(TEST)"

// C > I, D > J, E > K, F > L, G > M, H > O, J > N
const SOURCE_FOLDER_ID = "1uhi5F3FZECo7chWpdfsvjUyMagPDPgX7"
const SN_ORDER_IMPORT = "OrderImport"

const START_COLUMN_INDEX = 8 // column I

function onOpen() {
    const ui = SpreadsheetApp.getUi()
    const menu = ui.createMenu(APP_NAME)
    menu.addItem("Get label data", "getLabelData")
    menu.addItem(`Copy to ${SN_USPS}`, 'copyToUsps')
    menu.addSeparator()
    menu.addItem("List Shipments(test)", "listShipments")
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
        SpreadsheetApp.getUi().alert(APP_NAME, e.message, SpreadsheetApp.getUi().ButtonSet.OK)
    }
}

const copyToUsps = () => {
    const startTime = new Date().getTime()
    try {
        new App().copyToUsps()
        const endTime = new Date().getTime()
        const usedTime = Math.floor((endTime - startTime) / 1000)
        SpreadsheetApp.getActive().toast(`Copied to ${SN_USPS}. Used time ${usedTime}s.`, APP_NAME)
    } catch (e) {
        SpreadsheetApp.getUi().alert(APP_NAME, e.message, SpreadsheetApp.getUi().ButtonSet.OK)
    }
}

const listShipments = () => {
    new ShipStation().listShipments()
}

class App {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.ui = SpreadsheetApp.getUi()
        this.sourceFolder = this.getFolderById(SOURCE_FOLDER_ID)
        this.wsLabelPasteHere = this.ss.getSheetByName(SN_LABEL_PASTE_HERE)
        this.wsUsps = this.ss.getSheetByName(SN_USPS)
    }

    getFolderById(id) {
        try {
            return DriveApp.getFolderById(id)
        } catch (e) {
            return null
        }
    }

    isAppValidForGetLabelData() {
        if (!this.sourceFolder) return { valid: false, message: `Invalid source folder id: "${SOURCE_FOLDER_ID}".` }
        if (!this.wsLabelPasteHere) return { valid: false, message: `We can't find the sheet "${SN_LABEL_PASTE_HERE}".` }
        const message = `"${APP_NAME}" APP is ready to go.`
        this.ss.toast(message, APP_NAME, 2)
        return { valid: true, message }
    }

    isAppValidForCopyToUsps() {
        if (this.ss.getActiveSheet().getName() !== SN_LABEL_PASTE_HERE) return { valid: false, message: `Active sheet is not "${SN_LABEL_PASTE_HERE}."` }
        if (!this.wsUsps) return { valid: false, message: `We can't find the sheet "${SN_USPS}".` }
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
        const { valid, message } = this.isAppValidForGetLabelData()
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

    copyToUsps() {
        const { valid, message } = this.isAppValidForCopyToUsps()
        if (!valid) {
            this.ui.alert(APP_NAME, message, this.ui.ButtonSet.OK)
            return
        }
        // copy selected rows to USPS sheet
        const ws = this.ss.getActiveSheet()
        const selection = ws.getSelection()
        const ranges = selection.getActiveRangeList().getRanges()
        const items = []
        ranges.forEach(range => {
            const startRow = range.getRow()
            const values = range.getDisplayValues()
            values.forEach((v, i) => {
                const isRowHidden = ws.isRowHiddenByUser(startRow + i) || ws.isRowHiddenByFilter(startRow + i)
                if (!isRowHidden) {
                    const [order, shop, , code, , , , , buyer, addressFirstLine, addressSecondLine, city, state, country, zip] = v
                    items.push([order, shop, code, buyer, `${addressFirstLine} ${addressSecondLine}`, city, state, country, zip])
                }
            })
        })
        const lastRow = this.wsUsps.getDataRange().getLastRow()
        this.wsUsps.getRange(lastRow + 1, 1, items.length, items[0].length).setValues(items).activate()
    }
}

class ShipStation {
    constructor() {
        const key = "API_KEY"
        const secret = "API_SECRET"
        this.token = Utilities.base64Encode(`${key}:${secret}`)
        this.endpoint = "https://ssapi.shipstation.com/"
        this.ss = SpreadsheetApp.getActive()
        this.ui = SpreadsheetApp.getUi()
        this.wsShipments = this.ss.getSheetByName(SN_SHIPMENTS) || this.ss.insertSheet(SN_SHIPMENTS)
    }
    listShipments() {
        const url = `${this.endpoint}shipments?page=1&pageSize=50`
        const params = {
            method: "get",
            headers: {
                Authorization: `Basic ${this.token}`
            },
        }
        const response = UrlFetchApp.fetch(url, params)
        const code = response.getResponseCode()
        if (code !== 200) return this.ui.alert(APP_NAME, `Error Code ${code}`, this.ui.ButtonSet.OK)
        const content = JSON.parse(response.getContentText())
        const shipments = content.shipments
        if (!shipments) return this.ui.alert(APP_NAME, `No shipments found`, this.ui.ButtonSet.OK)
        const values = shipments.map(shipment => {
            return [
                shipment.shipmentId,
                shipment.shipmentCost,
                shipment.shipDate,
                shipment.trackingNumber,
                shipment.shipTo.name,
                shipment.shipTo.street1,
                shipment.shipTo.street2,
                shipment.shipTo.city,
                shipment.shipTo.state,
                shipment.shipTo.country,
                shipment.shipTo.postalCode,
            ]
        })
        values.unshift(["ID", "Cost", "Date", "Tracking #", "Name", "Street1", "Street 2", "City", "State", "Country", "Postal Code"])
        this.wsShipments.clear()
        this.wsShipments.getRange(1, 1, values.length, values[0].length).setValues(values).activate()
    }
}