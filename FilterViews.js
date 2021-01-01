// still in developing

function filterTest(){
  const app = new FilterViews()
  console.log(app.getCurrentFilter())
}
class FilterViews {
  constructor(){
    this.id = SpreadsheetApp.getActive().getId()
    this.ss = Sheets.Spreadsheets.get(this.id)
    this.sheets = this.ss.sheets
    this.wsMaster = this.sheets.find(sheet => sheet.properties.title === SN_MASTER)
  }
  getCurrentFilter(){
//    const filter = this.ss.sheets[].filterViews[]
    return this.wsMaster.filterViews.map(filter => {
      return {
                                     id: filter.filterViewId,
                                     title: filter.title,
                                     range: filter.range,
                                     rangeId: filter.namedRangeId,
                                     criteria: filter.criteria,
                                         zero: filter.criteria ? (filter.criteria["0"] ? filter.criteria["0"].hiddenValues : undefined) : undefined,
      }                                 
    })
  }
}
