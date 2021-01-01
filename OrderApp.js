const CN_STATUS = "A"
const CN_ORDER = "B"
const CN_SHOP = "D"
const CN_ADDRESS = "H"

const CN_MATERIAL_1 = "J"
const CN_MATERIAL_2 = "Q"
const CN_MATERIAL_3 = "X"
const CN_MATERIAL_4 = "AE"
const CN_MATERIAL_5 = "AL"

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
        if(lastRow) return lastRow
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
            }catch(e){
               //pass
            }
        }

        if (blobs.length) {
            const zipFile = this.createZip(blobs)
            zips.push(zipFile)
        }

        return zips
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
                                    value.push(`=IMAGE("${v[index + 4]}")`) // url
                                })
                                values.push(value)
                                range.getCell(i + 1, 1).setValue(`${ss.getName()}\n${sheetName}`)
                            }

                        }
                    } else {
                      if (status !== ""){
                        errors.push(`The "${order}" in the shop "${shop}" was added to "${status}"`)
                      } else {
                        errors.push(`"${IMAGE_NOT_FOUND}" for the order "${order}" in the shop "${shop}".` )
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
