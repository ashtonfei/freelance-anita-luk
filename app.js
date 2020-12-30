const APP_NAME = "MCO EMB"
const SN_PASTE_HERE = "PASTE HERE"
const SN_PASTE = "1PASTE"
const INDEX_ORDER = 3 // column D
const INDEX_PHONE = 6 // column G

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
        ui.alert(APP_NAME, e.message, ui.ButtonSet.OK)
    }
}

class App {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.wsPasteHere = this.ss.getSheetByName(SN_PASTE_HERE)
        this.wsPaste = this.ss.getSheetByName(SN_PASTE)
    }

    getPhones(v, formulaValues) {
        const phones = []
        const order = v[INDEX_ORDER]
        for (let i = 0; i < 5; i++) {
            const index = INDEX_PHONE + i * 3
            const phone = v[index]
            const model = v[index + 1]
            const image = formulaValues[index + 2]
            if (phone && model && image) phones.push([phone, model, image])
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
    }
}
