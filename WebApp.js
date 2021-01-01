function doGet() {
    return HtmlService.createTemplateFromFile("index").evaluate().setTitle(APP_NAME).addMetaTag("viewport", "width=device-width, initial-scale=1.0").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function getAppData() {
    const app = new WebApp()
    const data = app.getAppData()
    return JSON.stringify(data)
}

class WebApp {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
    }

    downloadFiles(ids) {
        const blobs = ids.map(id => DriveApp.getFileById(id).getBlob())
        const name = "Download Images " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMyyhhmmss")
        const zip = Utilities.zip(blobs, name)
        const content = zip.getDataAsString()
        const output = ContentService.createTextOutput(content)
        output.downloadAsFile(name)
        return output
    }

    getAppData() {
        const ws = this.ss.getSheetByName(SN_MASTER)
        const values = ws.getDataRange().getValues()
        const headers = ["ID", "Report", "Sheet", ...values.shift()]
        const orders = []
        values.forEach(v => {
            const status = v[0].toString().trim()
            const order = v[1].toString().trim().toUpperCase()
            const shop = v[3].toString().trim().toUpperCase()
            const id = `${order}${shop}`
            const [reportName, sheetName] = status.split("\n")
            if (reportName && sheetName) orders.push([id, reportName, sheetName, ...v])
        })
        return { headers, orders }
    }
}
