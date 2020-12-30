const APP_NAME = "MCO EMB"
const SN_PASTE_HERE = "PASTE HERE"
const SN_PASTE = "1PASTE"
const INDEX_ORDER = 3 // column D
const INDEX_PHONE = 6 // column G

const INDEX_IMAGE = 3 // column D
const INDEX_FILENAME = 5 // column F

const NA = "#N/A"
const KEY_ID = "id="

function onOpen() {
    const menu = SpreadsheetApp.getUi().createMenu(APP_NAME)
    menu.addItem("Copy & Paste", "copyAndPaste")
    menu.addItem("Download Images", "downloadImages")
    menu.addToUi()
}

function copyAndPaste() {
    const startTime = new Date().getTime()
    const ss = SpreadsheetApp.getActive()
    const ui = SpreadsheetApp.getUi()
    try {
        ss.toast("Copy && Paste ...", APP_NAME)
        const app = new App()
        app.copyAndPaste()
        const endTime = new Date().getTime()
        const usedTime = Math.floor((endTime - startTime) / 1000)
        ss.toast(`Done. Used time ${usedTime}s.`, APP_NAME)
    } catch (e) {

    }
}

function downloadImages() {
    const startTime = new Date().getTime()
    const ss = SpreadsheetApp.getActive()
    const ui = SpreadsheetApp.getUi()
    try {
        const app = new App()
        app.downloadImages()
        const endTime = new Date().getTime()
        const usedTime = Math.floor((endTime - startTime) / 1000)
        ss.toast(`Done. Used time ${usedTime}s.`, APP_NAME)
    } catch (e) {
        ui.alert(APP_NAME, e.message, ui.ButtonSet.OK)
    }

}

class App {
    constructor() {
        this.ui = SpreadsheetApp.getUi()
        this.ss = SpreadsheetApp.getActive()
        this.wsPasteHere = this.ss.getSheetByName(SN_PASTE_HERE)
        this.wsPaste = this.ss.getSheetByName(SN_PASTE)

        this.downloadsFolder = this.getDownloadsFolder()
    }

    getPhones(v, formulaValues) {
        const phones = []
        const order = v[INDEX_ORDER]
        for (let i = 0; i < 5; i++) {
            const index = INDEX_PHONE + i * 3
            const phone = v[index]
            const image = formulaValues[index + 2]
            if (order && phone && image) phones.push([order, phone, image])
        }
        return phones
    }

    getOrders() {
        const dataRange = this.wsPasteHere.getDataRange()
        const values = dataRange.getValues()
        const formulas = dataRange.getFormulas()
        let orders = []
        values.forEach((v, i) => {
            const formulaValues = formulas[i]
            const phones = this.getPhones(v, formulaValues)
            orders = [...orders, ...phones]
        })
        return orders
    }

    copyAndPaste() {
        const orders = this.getOrders()
        this.wsPaste.getRange("B:D").clearContent()
        this.wsPaste.getRange(1, 2, orders.length, orders[0].length).setValues(orders)
        this.wsPaste.activate()
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
        this.ss.toast("Creating ZIPs ...", APP_NAME, 30)
        const zips = []
        let blobs = []
        let totalSize = 0
        const maxSize = 50000000
        for (let i = 0; i < ids.length; i++) {
            try {
                const file = DriveApp.getFileById(ids[i])
                const blob = file.getAs(MimeType.JPEG)
                let filename = filenames[i] + ".jpg"
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
        this.ss.toast("ZIPs created!", APP_NAME, 30)
        return zips
    }

    getImages() {
        const dataRange = this.wsPaste.getDataRange()
        const values = dataRange.getValues()
        const formulas = dataRange.getFormulas()
        const ids = []
        const filenames = []
        values.forEach((v, i) => {
            const formula = formulas[i][INDEX_IMAGE]
            if (formula.indexOf(KEY_ID) !== -1) {
                const id = formula.split(KEY_ID)[1].replace('")', "")
                const filename = v[INDEX_FILENAME]
                if (filename !== NA) {
                    ids.push(id)
                    if (filenames.indexOf(filename) === -1) {
                        filenames.push(filename)
                    } else {
                        const count = filenames.filter(name => name.indexOf(filename) !== -1).length
                        filenames.push(`${filenames}(${count})`)
                    }
                } else {
                    ids.push(NA)
                    filenames.push(NA)
                }
            }
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
