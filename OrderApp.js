const CN_STATUS = "A"
const CN_ORDER = "B"
const CN_SHOP = "D"
const CN_ADDRESS = "H"

const CN_MATERIAL_1 = "J"
const CN_MATERIAL_2 = "Q"
const CN_MATERIAL_3 = "X"
const CN_MATERIAL_4 = "AE"
const CN_MATERIAL_5 = "AL"
const COLOR_PINK = "#f4cccc"

class OrderApp {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.ui = SpreadsheetApp.getUi()
        this.folder = DriveApp.getFolderById(CASE_PRINT_FOLDER_ID)

        this.downloadsFolder = this.getDownloadsFolder()

        this.indexStatus = getColumnIndex(CN_STATUS)
        this.indexOrder = getColumnIndex(CN_ORDER)
        this.indexShop = getColumnIndex(CN_SHOP)
        this.indexAddress = getColumnIndex(CN_ADDRESS)

        this.indexMaterial1 = getColumnIndex(CN_MATERIAL_1)
        this.indexMaterial2 = getColumnIndex(CN_MATERIAL_2)
        this.indexMaterial3 = getColumnIndex(CN_MATERIAL_3)
        this.indexMaterial4 = getColumnIndex(CN_MATERIAL_4)
        this.indexMaterial5 = getColumnIndex(CN_MATERIAL_5)
        this.indexMaterials = [this.indexMaterial1, this.indexMaterial2, this.indexMaterial3, this.indexMaterial4, this.indexMaterial5]
    }

    getDownloadsFolder() {
        const rootFolder = DriveApp.getRootFolder()
        const folderName = `${APP_NAME} Downloads`
        const folders = rootFolder.getFoldersByName(folderName)
        if (folders.hasNext()) return folders.next()
        return rootFolder.createFolder(folderName)
    }

    getValidMaterials() {
        return this.ss.getSheetByName(SN_LIST).getDataRange().getValues().map(([item]) => item.toString().trim()).filter(v => v !== "")
    }

    getInvliadMaterials(materials, validMaterials) {
        return materials.filter(v => validMaterials.indexOf(v) === -1)
    }

    getSpreadsheets() {
        const files = this.folder.getFilesByType(MimeType.GOOGLE_SHEETS)
        const spreadsheets = []
        while (files.hasNext()) {
            const file = files.next()
            const ss = SpreadsheetApp.open(file)
            const tabs = ss.getSheets().map(sheet => sheet.getName())
            spreadsheets.push({
                name: ss.getName(),
                id: ss.getId(),
                tabs: tabs,
            })
        }
        return spreadsheets
    }

    openDialog() {
        const activeSheet = this.ss.getActiveSheet()
        if (activeSheet.getName() !== SN_MASTER) {
            this.ui.alert(APP_NAME, `You are not in the sheet "${SN_MASTER}".`, this.ui.ButtonSet.OK)
            const ws = this.ss.getSheetByName(SN_MASTER)
            if (ws) ws.activate()
        } else {
            const userInterface = HtmlService.createTemplateFromFile("sidebar").evaluate().setTitle(APP_NAME)
            this.ui.showSidebar(userInterface)
        }
    }

    getLastRow(ws) {
        const values = ws.getDataRange().getValues()
        const lastRow = values.findIndex(([a, b]) => a == "" && b == "") + 1
        if (lastRow) return lastRow
        return values.length + 1
    }

    getFileIdByUrl(url) {
        let id = url.split("id=")[1]
        if (!id) return null
        return id
    }

    getFileName(names, phone, material) {
        phone = phone.toString().trim()
        material = material.toString().trim().split("/")[1] || ""
        const filename = `${phone}${material}`.replace(/\s/g, "").replace(/\//g, "&")
        const matchs = names.filter(name => name.indexOf(filename) === 0)
        return `${filename}${matchs.length + 1}`
    }

    createZip(blobs) {
        const zip = Utilities.zip(blobs)
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMyyhhmmss")
        const name = `Downloaded Images ${timestamp}`
        const zipFile = this.downloadsFolder.createFile(zip).setName(name)
        return zipFile
    }

    createZips({ ids, filenames }) {
        const zips = []
        let blobs = []
        let totalSize = 0
        const maxSize = 50000000
        for (let i = 0; i < ids.length; i++) {
            try {
                const file = DriveApp.getFileById(ids[i])
                const blob = file.getBlob()
                const fileType = file.getMimeType()
                let filename = filenames[i]
                if (fileType === MimeType.PNG) {
                    filename += ".png"
                } else {
                    filename += ".jpg"
                }
                totalSize += file.getSize()
                if (totalSize > maxSize) {
                    const zipFile = this.createZip(blobs)
                    zips.push(zipFile)
                    blobs = []
                    totalSize = 0
                }
                blob.setName(filename)
                blobs.push(blob)
            } catch (e) {
                //pass
            }
        }

        if (blobs.length) {
            const zipFile = this.createZip(blobs)
            zips.push(zipFile)
        }

        return zips
    }

    findDuplicateRecipents() {
        const activeSheet = this.ss.getActiveSheet()
        if (activeSheet.getName() !== SN_MASTER) {
            this.ui.alert(APP_NAME, `You are not in sheet "${SN_MASTER}".`, this.ui.ButtonSet.OK)
            return
        }

        const ranges = this.ss.getSelection().getActiveRangeList().getRanges()
        const recipients = {}
        ranges.forEach(range => {
            const rangeStartRow = range.getRow()
            const rangeValues = range.getDisplayValues()
            rangeValues.forEach((v, i) => {
                const row = rangeStartRow + i
                const isRowHidden = activeSheet.isRowHiddenByUser(row) || activeSheet.isRowHiddenByFilter(row)
                //                const isRowHidden = false
                if (!isRowHidden) {
                    if (v[this.indexShop] === undefined || v[this.indexAddress] == undefined) {
                        this.ui.alert(APP_NAME, `You need to select rows!`, this.ui.ButtonSet.OK)
                        return
                    }
                    const shop = v[this.indexShop].trim()
                    const address = v[this.indexAddress].trim()
                    const name = address.split("\n")[0].trim()
                    const order = v[this.indexOrder].trim()
                    const key = `${shop}${name}`.toUpperCase()
                    const recipent = recipients[key]
                    if (recipent) {
                        recipent.push({ shop, address, row, order, name })
                    } else {
                        recipients[key] = [{ shop, address, row, order, name }]
                    }
                }
            })
        })
        const duplicates = Object.keys(recipients).filter(key => recipients[key].length > 1).map(key => recipients[key])
        if (duplicates.length === 0) {
            this.ui.alert(APP_NAME, "No duplicated recipents found in selected rows.", this.ui.ButtonSet.OK)
        } else {
            let html = `<div style="font-family: sans-serif;">
                <h3>${duplicates.length} Duplicate Recipients Found:</h3>
                <table style="border-collapse: collapse; width: 100%;">
                <tr><th style="padding: 3px 6px; border: 1px solid black;">Recipent Name</th>
                <th style="padding: 3px 6px; border: 1px solid black;">Shop</th>
                <th style="padding: 3px 6px; border: 1px solid black;">Count</th>
                <th style="padding: 3px 6px; border: 1px solid black;">Address</th></tr>`
            duplicates.forEach(v => html += `<tr>
                <td style="padding: 3px 6px; border: 1px solid black;">${v[0].name}</td>
                <td style="padding: 3px 6px; border: 1px solid black;">${v[0].shop}</td>
                <td style="padding: 3px 6px; border: 1px solid black;">${v.length}</td>
                <td style="padding: 3px 6px; border: 1px solid black;">${v[0].address}</td></tr>`)
            html += "</table></div>"
            const userInterface = HtmlService.createHtmlOutput(html).setTitle(APP_NAME)
            this.ss.show(userInterface)
        }
    }

    /**
    * update address, phone, material, link, and preview in destination sheet
    */
    updateOrder() {
        const activeSheet = this.ss.getActiveSheet()
        const validColumnIndexes = [this.indexAddress]
        this.indexMaterials.forEach(index => {
            validColumnIndexes.push(index)
            validColumnIndexes.push(index - 1)
            validColumnIndexes.push(index + 2)
        })
        if (activeSheet.getName() !== SN_MASTER) {
            this.ui.alert(APP_NAME, `You are not in the sheet "${SN_MASTER}".`, this.ui.ButtonSet.OK)
            return
        }
        const activeCell = activeSheet.getActiveCell()
        const columnIndex = activeCell.getColumn() - 1
        const row = activeCell.getRow()
        const rowValues = activeSheet.getRange(`${row}:${row}`).getDisplayValues()[0]
        const value = activeCell.getDisplayValue().trim()

        const status = rowValues[this.indexStatus].trim()
        const [spreadsheetName, sheetName] = status.split("\n").map(v => v.trim())
        const order = rowValues[this.indexOrder].trim().toUpperCase()
        const shop = rowValues[this.indexShop].trim().toUpperCase()

        if (validColumnIndexes.indexOf(columnIndex) === -1) {
            this.ui.alert(APP_NAME, `Selected cell is invalid, you can choose "address", "phone", "material", "design".`, this.ui.ButtonSet.OK)
            return
        }
        if (!(status && order && shop)) {
            this.ui.alert(APP_NAME, `"Status", "Order", "Shop" can't be empty.`, this.ui.ButtonSet.OK)
            return
        }
        const spreadsheet = this.getSpreadsheets().find(v => v.name === spreadsheetName)
        if (!spreadsheet) {
            this.ui.alert(APP_NAME, `Can't find spreadsheet "${spreadsheetName}".`, this.ui.ButtonSet.OK)
            return
        }
        const sheet = SpreadsheetApp.openById(spreadsheet.id).getSheetByName(sheetName)
        if (!sheet) {
            this.ui.alert(APP_NAME, `Can't find sheet "${sheetName}" in spreadsheet "${spreadsheetName}".`, this.ui.ButtonSet.OK)
            return
        }
        const findRowIndex = sheet.getDataRange().getDisplayValues().findIndex(v => v[0].trim().toUpperCase() === order && v[1].trim().toUpperCase() === shop)
        if (findRowIndex === -1) {
            this.ui.alert(APP_NAME, `Can't find order "${order}" & shop "shop" in spreadsheet "${spreadsheetName}".`, this.ui.ButtonSet.OK)
            return
        }

        if (columnIndex === this.indexAddress) {
            const cell = sheet.getRange(`F${findRowIndex + 1}`)
            const text = cell.getDisplayValue().trim()
            const note = [cell.getNote().trim(), `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MMM/yy hh:mm:ss")}\n${text}`].filter(v => v !== "").join(`\n\n`)
            if (text) cell.setNote(note)
            cell.setValue(value).setBackground(COLOR_PINK)
            this.ss.toast(`Address infomation has been updated in "${spreadsheetName}".`, APP_NAME)
        } else if ([
            this.indexMaterial1 - 1,
            this.indexMaterial2 - 1,
            this.indexMaterial3 - 1,
            this.indexMaterial4 - 1,
            this.indexMaterial5 - 1].indexOf(columnIndex) !== -1) {
            const i = [
                this.indexMaterial1 - 1,
                this.indexMaterial2 - 1,
                this.indexMaterial3 - 1,
                this.indexMaterial4 - 1,
                this.indexMaterial5 - 1].indexOf(columnIndex)
            const cell = sheet.getRange(findRowIndex + 1, columnIndex - 1 + i - i * 4)
            const text = cell.getDisplayValue().trim()
            const note = [cell.getNote().trim(), `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MMM/yy hh:mm:ss")}\n${text}`].filter(v => v !== "").join(`\n\n`)
            if (text) cell.setNote(note)
            cell.setValue(value).setBackground(COLOR_PINK)
            this.ss.toast(`Phone infomation has been updated in "${spreadsheetName}".`, APP_NAME)
        } else if ([
            this.indexMaterial1,
            this.indexMaterial2,
            this.indexMaterial3,
            this.indexMaterial4,
            this.indexMaterial5].indexOf(columnIndex) !== -1) {
            const i = [
                this.indexMaterial1,
                this.indexMaterial2,
                this.indexMaterial3,
                this.indexMaterial4,
                this.indexMaterial5].indexOf(columnIndex)
            const cell = sheet.getRange(findRowIndex + 1, columnIndex - 1 + i - i * 4)
            const text = cell.getDisplayValue().trim()
            const note = [cell.getNote().trim(), `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MMM/yy hh:mm:ss")}\n${text}`].filter(v => v !== "").join(`\n\n`)
            if (text) cell.setNote(note)
            cell.setValue(value).setBackground(COLOR_PINK)
            this.ss.toast(`Material infomation has been updated in "${spreadsheetName}".`, APP_NAME)
        } else if ([
            this.indexMaterial1 + 2,
            this.indexMaterial2 + 2,
            this.indexMaterial3 + 2,
            this.indexMaterial4 + 2,
            this.indexMaterial5 + 2].indexOf(columnIndex) !== -1) {
            const i = [
                this.indexMaterial1 + 2,
                this.indexMaterial2 + 2,
                this.indexMaterial3 + 2,
                this.indexMaterial4 + 2,
                this.indexMaterial5 + 2].indexOf(columnIndex)
            const cell = sheet.getRange(findRowIndex + 1, columnIndex - 2 + i - i * 4)
            const text = cell.getDisplayValue().trim()
            const note = [cell.getNote().trim(), `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MMM/yy hh:mm:ss")}\n${text}`].filter(v => v !== "").join(`\n\n`)
            const link = rowValues[columnIndex + 2]
            if (text) cell.setNote(note)
            cell.setValue(link).setBackground(COLOR_PINK)
            sheet.getRange(findRowIndex + 1, columnIndex - 1 + i - i * 4).setValue(`=IMAGE("${link}")`)
            this.ss.toast(`Image infomation has been updated in "${spreadsheetName}".`, APP_NAME)
        } else {
            this.ui.alert(APP_NAME, `You need select a cell in the column "${CN_ADDRESS}".`, this.ui.ButtonSet.OK)
            return
        }
    }

    downloadImages() {
        this.ss.toast("Creating images bundle...", APP_NAME)
        const masterSheet = this.ss.getSheetByName(SN_MASTER)
        const ranges = this.ss.getSelection().getActiveRangeList().getRanges()
        const ids = []
        const filenames = []
        ranges.forEach((range, i) => {
            const rangeValues = range.getValues()
            const row = range.getRow()
            rangeValues.forEach((v, i) => {
                const isRowHidden = masterSheet.isRowHiddenByFilter(row + i) || masterSheet.isRowHiddenByUser(row + i)
                if (!isRowHidden) {
                    this.indexMaterials.forEach(index => {
                        const material = v[index]
                        const phone = v[index - 1]
                        const filename = this.getFileName(filenames, phone, material)
                        const url = v[index + 4]
                        const id = this.getFileIdByUrl(url)
                        if (id) {
                            ids.push(id)
                            filenames.push(filename)
                        }
                    })
                }
            })
        })
        if (ids.length) {
            const zips = this.createZips({ ids, filenames })
            const downloadUrls = zips.map(zip => zip.getDownloadUrl())
            this.ui.alert(APP_NAME, `Images ZIP has been created on your Google Dirve, you can download it here.\n${downloadUrls.join("\n")}`, this.ui.ButtonSet.OK)
            return downloadUrls
        } else {
            this.ui.alert(APP_NAME, "There is no valid images in your selection.", this.ui.ButtonSet.OK)
            return null
        }
    }

    addOrders({ id, sheetName }) {
        this.ss.toast("Copy & Paste....", APP_NAME)
        const validMaterials = this.getValidMaterials()
        const ss = SpreadsheetApp.openById(id)
        const targetSheet = ss.getSheetByName(sheetName)

        const masterSheet = this.ss.getSheetByName(SN_MASTER)
        const ranges = this.ss.getSelection().getActiveRangeList().getRanges()
        const values = []
        const errors = []
        ranges.forEach(range => {
            const rangeValues = range.getValues()
            const row = range.getRow()
            rangeValues.forEach((v, i) => {
                const isRowHidden = masterSheet.isRowHiddenByFilter(row + i) || masterSheet.isRowHiddenByUser(row + i)
                if (!isRowHidden) {
                    const status = v[this.indexStatus]
                    const order = v[this.indexOrder]
                    const shop = v[this.indexShop]
                    const address = v[this.indexAddress]
                    const isImageNotFound = v.indexOf(IMAGE_NOT_FOUND) !== -1
                    if (status === "" && !isImageNotFound) {
                        const materials = [
                            v[this.indexMaterial1].toString().trim(),
                            v[this.indexMaterial2].toString().trim(),
                            v[this.indexMaterial3].toString().trim(),
                            v[this.indexMaterial4].toString().trim(),
                            v[this.indexMaterial5].toString().trim()
                        ].filter(v => v !== "")
                        if (materials.length === 0) {
                            const error = "No material type"
                            errors.push(`${error} for ${order} in ${shop}.`)
                            range.getCell(i + 1, 1).setValue(error)
                        } else {
                            const invalidMaterials = this.getInvliadMaterials(materials, validMaterials)
                            if (invalidMaterials.length) {
                                errors.push(`Invalid materials: ${invalidMaterials.join(",")}`)
                                range.getCell(i + 1, 1).setValue(`Invalid material types:\n${invalidMaterials.join("\n")}`)
                            } else {
                                const value = [order, shop, false, null, null, address]
                                this.indexMaterials.forEach(index => {
                                    value.push(v[index - 1]) // phone
                                    value.push(v[index]) // material
                                    value.push(v[index + 4]) // image url
                                    value.push(`=IMAGE("${v[index + 4]}")`) // url
                                })
                                values.push(value)
                                range.getCell(i + 1, 1).setValue(`${ss.getName()}\n${sheetName}`)
                            }

                        }
                    } else {
                        if (status !== "") {
                            const error = `The "${order}" in the shop "${shop}" was added to "${status}"`
                            errors.push(error)
                        } else {
                            const error = `"${IMAGE_NOT_FOUND}" for the order "${order}" in the shop "${shop}".`
                            errors.push(error)
                            range.getCell(i + 1, 1).setValue(error)
                        }
                    }
                }
            })
        })
        if (values.length) {
            const row = this.getLastRow(targetSheet)
            targetSheet.getRange(row, 1, values.length, values[0].length).setValues(values)
        }
        SpreadsheetApp.flush()
        this.ui.alert(APP_NAME, `Copied orders: ${values.length}\nUncopied orders: ${errors.length}`, this.ui.ButtonSet.OK)
    }
}
