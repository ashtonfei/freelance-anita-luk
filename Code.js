const APP_NAME = "MCO KV1"

const ID_MASTER_CASE_ORDER = "1HJJyixsdhdEa2swtZKEO9Mz1vqY4ZHmsWnc_PaSdYj4"
const SN_MASTER = "master"

const HEADER_TRACKING = "Tracking"
const RN_TRACKING = "E1"

const INDEX_ORDER = 3 // column D
const INDEX_PHONE = 6 // column G

const INDEX_IMAGE = 3 // column D
const INDEX_FILENAME = 5 // column F


const CN_MATERIAL_1 = "H"
const CN_MATERIAL_2 = "L"
const CN_MATERIAL_3 = "P"
const CN_MATERIAL_4 = "T"
const CN_MATERIAL_5 = "X"

const NA = "#N/A"
const KEY_ID = "id="

function onOpen() {
    const menu = SpreadsheetApp.getUi().createMenu(APP_NAME)
    menu.addItem("Send tracking to master", "sendTrackingNumberToMaster")
    menu.addItem("Download selected images", "downloadImages")
    menu.addToUi()
}

function sendTrackingNumberToMaster() {
    const startTime = new Date().getTime()
    const ss = SpreadsheetApp.getActive()
    const ui = SpreadsheetApp.getUi()
    if (ss.getActiveSheet().getRange(RN_TRACKING).getValue() !== HEADER_TRACKING) {
        ui.alert(APP_NAME, `Active sheet is invalid, it must have "${HEADER_TRACKING}" in "${RN_TRACKING}".\nSelect the correct sheet and run it again.`, ui.ButtonSet.OK)
        return
    }
    try {
        ss.toast("Sending ...", APP_NAME, 10)
        const app = new App()
        const { success, fail } = app.sendTrackingNumberToMaster()
        const endTime = new Date().getTime()
        const usedTime = Math.floor((endTime - startTime) / 1000)
        ui.alert(APP_NAME, `Success: ${success}\nOrder&Shop not found in master: ${fail}\nUsed time ${usedTime}s.`, ui.ButtonSet.OK)
    } catch (e) {
        ui.alert(APP_NAME, e.message, ui.ButtonSet.OK)
    }
}

function downloadImages() {
    const startTime = new Date().getTime()
    const ss = SpreadsheetApp.getActive()
    const ui = SpreadsheetApp.getUi()
    if (ss.getActiveSheet().getRange(RN_TRACKING).getValue() !== HEADER_TRACKING) {
        ui.alert(APP_NAME, `Active sheet is invalid, it must have "${HEADER_TRACKING}" in "${RN_TRACKING}".\nSelect the correct sheet and run it again.`, ui.ButtonSet.OK)
        return
    }
    try {
        ss.toast("Preparing downloads ...", APP_NAME, 30)
        const app = new App()
        app.downloadImages()
        const endTime = new Date().getTime()
        const usedTime = Math.floor((endTime - startTime) / 1000)
        ss.toast(`Done. Used time ${usedTime}s.`, APP_NAME, 30)
    } catch (e) {
        ui.alert(APP_NAME, e.message, ui.ButtonSet.OK)
    }

}

class App {
    constructor() {
        this.ui = SpreadsheetApp.getUi()
        this.ss = SpreadsheetApp.getActive()
        this.wsActive = this.ss.getActiveSheet()

        this.downloadsFolder = this.getDownloadsFolder()

        this.indexMaterial1 = this.getColumnIndex(CN_MATERIAL_1)
        this.indexMaterial2 = this.getColumnIndex(CN_MATERIAL_2)
        this.indexMaterial3 = this.getColumnIndex(CN_MATERIAL_3)
        this.indexMaterial4 = this.getColumnIndex(CN_MATERIAL_4)
        this.indexMaterial5 = this.getColumnIndex(CN_MATERIAL_5)
        this.indexMaterials = [this.indexMaterial1, this.indexMaterial2, this.indexMaterial3, this.indexMaterial4, this.indexMaterial5]
    }

    getColumnIndex(name) {
        return this.wsActive.getRange(`${name.trim()}1`).getColumn() - 1
    }

    getTrackingNumbers() {
        const trackingNumbers = {}
        this.wsActive.getDataRange().getDisplayValues().slice(1)
            .forEach(([order, shop, , , trackingNumber]) => {
                order = order.trim().toUpperCase()
                shop = shop.trim().toUpperCase()

                if (order && shop) {
                    trackingNumber = trackingNumber.trim()
                    trackingNumbers[`${order}${shop}`] = trackingNumber
                }
            })
        return trackingNumbers
    }



    sendTrackingNumberToMaster() {
        const trackingNumbers = this.getTrackingNumbers()
        //        const successNumbers = []
        let success = 0

        const wsMaster = SpreadsheetApp.openById(ID_MASTER_CASE_ORDER).getSheetByName(SN_MASTER)
        const rangeOrder = wsMaster.getRange("B:B")
        const orderValues = rangeOrder.getDisplayValues()

        const rangeShop = wsMaster.getRange("D:D")
        const shopValues = rangeShop.getDisplayValues()

        const rangeTrackingNumber = wsMaster.getRange("F:F")
        const trackingNumberValues = rangeTrackingNumber.getDisplayValues()

        //        console.log({trackingNumbers, orders: orderValues.slice(0,10), shops: shopValues.slice(0,10), trackings: trackingNumberValues.slice(0,10)})
        trackingNumberValues.forEach(([v], index) => {
            const order = orderValues[index][0].trim().toUpperCase()
            const shop = shopValues[index][0].trim().toUpperCase()
            const key = `${order}${shop}`
            const trackingNumber = trackingNumbers[key]
            if (trackingNumber !== undefined) {
                trackingNumberValues[index][0] = trackingNumber
                //            successNumbers.push({order, shop, trackingNumber})
                success++
            }
        })
        //        console.log(trackingNumbers)
        //        console.log(successNumbers)
        rangeTrackingNumber.setValues(trackingNumberValues)
        return { success, fail: Object.keys(trackingNumbers).length - success }
    }

    getDownloadsFolder() {
        const rootFolder = DriveApp.getRootFolder()
        const folderName = `${APP_NAME} Downloads`
        const folders = rootFolder.getFoldersByName(folderName)
        if (folders.hasNext()) return folders.next()
        return rootFolder.createFolder(folderName)
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
                const type = file.getMimeType()
                let filename = filenames[i] + ".jpg"
                if (type === MimeType.PNG) {
                    filename = filenames[i] + ".png"
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

    getFileIdByUrl(url) {
        let id = url.split("id=")[1]
        if (!id) return null
        return id
    }

    getFileName(names, phone, material) {
        phone = phone.toString().trim()
        material = material.toString().trim().split("/")[0] || ""
        material = material.trim()
        const filename = `${phone}_${material}`.replace(/\s+/g, " ").replace(/\//g, "&")
        const matchs = names.filter(name => name.indexOf(filename) === 0)
        return `${filename}_${matchs.length + 1}`
    }

    getImages() {
        const ranges = this.ss.getSelection().getActiveRangeList().getRanges()
        const ids = []
        const filenames = []
        ranges.forEach((range, i) => {
            const rangeValues = range.getValues()
            const row = range.getRow()
            rangeValues.forEach((v, i) => {
                const isRowHidden = this.wsActive.isRowHiddenByFilter(row + i) || this.wsActive.isRowHiddenByUser(row + i)
                if (!isRowHidden) {
                    this.indexMaterials.forEach(index => {
                        const material = v[index]
                        const phone = v[index - 1]
                        const filename = this.getFileName(filenames, phone, material)
                        const url = v[index + 1]
                        const id = this.getFileIdByUrl(url)
                        if (id) {
                            ids.push(id)
                            filenames.push(filename)
                        }
                    })
                }
            })
        })
        return { ids, filenames }
    }

    showDownloads(zips) {
        let html = `<div style="font-family: sans-serif;"><h2>Downloads</h2><p>Click the below links to download them.</p><div><ul style="padding: 0px;">`
        zips.forEach(zip => {
            const name = zip.getName()
            const url = zip.getDownloadUrl()
            const li = `<li style="display: block;"><a style="font-family: sans-serif;" href="${url}" target="_balnk">${name}</a></li>`
            html += li
        })
        html += `</ul><p><small>All ZIPs can be found here <a style="font-family: sans-serif;" href="${this.downloadsFolder.getUrl()}" target="_balnk">${this.downloadsFolder.getName()}</a></small></p>`
        html += "</div>"
        const userInterface = HtmlService.createHtmlOutput(html).setTitle(APP_NAME)
        this.ss.show(userInterface)
        this.ss.getSelection().getActiveRangeList().getRanges().forEach(range => range.setBackground("#ffa500"))
    }

    downloadImages() {
        let { ids, filenames } = this.getImages()
        if (ids.indexOf(NA) !== -1) {
            const dialog = this.ui.alert(APP_NAME + " (Warning)", `"${NA}" found in ARRAYFORMULA of sheet "${this.wsPaste.getName()}", download them anyway with "${NA}" orders ignored?.`, this.ui.ButtonSet.YES_NO)
            if (dialog === this.ui.Button.YES) {
                ids = ids.filter(v => v !== NA)
                filenames = filenames.filter(v => v !== NA)
                const zips = this.createZips({ ids, filenames })
                this.showDownloads(zips)
            }
            this.wsPaste.getRange("E1").activate()
        } else {
            const zips = this.createZips({ ids, filenames })
            this.showDownloads(zips)
            this.ss.toast("Done!", APP_NAME)
        }
    }
}