const IMAGE_FOLDER_ID = "1JcLdeY3fSUhTqmzFAwzUfycj5J-V67rA" // the root folder id of all images on your google drive

const CN_FOLDER_NAME = "D" // column name where folder name located

const CN_IMAGE_NAME_1 = "L" // column name where image name located
const CN_IMAGE_URL_1 = "N" // column name where image url located
const CN_IMAGE_1 = "O" // column name where image located

const CN_IMAGE_NAME_2 = "S" // column name where image name located
const CN_IMAGE_URL_2 = "U" // column name where image url located
const CN_IMAGE_2 = "V" // column name where image located

const CN_IMAGE_NAME_3 = "Z" // column name where image name located
const CN_IMAGE_URL_3 = "AB" // column name where image url located
const CN_IMAGE_3 = "AC" // column name where image located

const CN_IMAGE_NAME_4 = "AG" // column name where image name located
const CN_IMAGE_URL_4 = "AI" // column name where image url located
const CN_IMAGE_4 = "AJ" // column name where image located

const CN_IMAGE_NAME_5 = "AN" // column name where image name located
const CN_IMAGE_URL_5 = "AP" // column name where image url located
const CN_IMAGE_5 = "AQ" // column name where image located


class ImageApp {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.rootFolder = DriveApp.getFolderById(IMAGE_FOLDER_ID)
        this.imageNotFound = IMAGE_NOT_FOUND

        this.indexFolderName = getColumnIndex(CN_FOLDER_NAME)

        this.indexImageName1 = getColumnIndex(CN_IMAGE_NAME_1)
        this.indexImageUrl1 = getColumnIndex(CN_IMAGE_URL_1)
        this.indexImage1 = getColumnIndex(CN_IMAGE_1)

        this.indexImageName2 = getColumnIndex(CN_IMAGE_NAME_2)
        this.indexImageUrl2 = getColumnIndex(CN_IMAGE_URL_2)
        this.indexImage2 = getColumnIndex(CN_IMAGE_2)

        this.indexImageName3 = getColumnIndex(CN_IMAGE_NAME_3)
        this.indexImageUrl3 = getColumnIndex(CN_IMAGE_URL_3)
        this.indexImage3 = getColumnIndex(CN_IMAGE_3)

        this.indexImageName4 = getColumnIndex(CN_IMAGE_NAME_4)
        this.indexImageUrl4 = getColumnIndex(CN_IMAGE_URL_4)
        this.indexImage4 = getColumnIndex(CN_IMAGE_4)

        this.indexImageName5 = getColumnIndex(CN_IMAGE_NAME_5)
        this.indexImageUrl5 = getColumnIndex(CN_IMAGE_URL_5)
        this.indexImage5 = getColumnIndex(CN_IMAGE_5)
    }

    getImages() {
        const folders = this.rootFolder.getFolders()
        const images = {}
        while (folders.hasNext()) {
            const folder = folders.next()
            const folderName = folder.getName().trim().toUpperCase()
            const jpegs = folder.getFilesByType(MimeType.JPEG)
            const pngs = folder.getFilesByType(MimeType.PNG)

            while (jpegs.hasNext()) {
                const image = jpegs.next()
                const imageName = image.getName().trim().toUpperCase()
                const key = `${folderName}${imageName}`.replace(".JPEG", "").replace(".JPG", "")
                const id = image.getId()
                const url = image.getUrl()
                const downloadUrl = `https://drive.google.com/uc?export=view&id=${id}`
                const formula = `=IMAGE("${downloadUrl}")`
                images[key] = { folderName, imageName, id, url, downloadUrl, formula }
            }

            while (pngs.hasNext()) {
                const image = pngs.next()
                const imageName = image.getName().trim().toUpperCase()
                const key = `${folderName}${imageName}`.replace(".PNG", "")
                const id = image.getId()
                const url = image.getUrl()
                const downloadUrl = `https://drive.google.com/uc?export=view&id=${id}`
                const formula = `=IMAGE("${downloadUrl}")`
                images[key] = { folderName, imageName, id, url, downloadUrl, formula }
            }
        }
        return images
    }

    getImageInfo(folderName, imageName) {
        folderName = folderName.trim()
        imageName = imageName.trim().toLowerCase()

        const rootImageFolder = DriveApp.getFolderById(IMAGE_FOLDER_ID)
        if (!rootImageFolder) return { error: "Root image folder not found" }
        const folders = rootImageFolder.getFoldersByName(folderName)
        if (!folders.hasNext()) return { error: "Image folder not found" }
        const folder = folders.next()
        const files = folder.getFiles()
        let image
        while (files.hasNext()) {
            const file = files.next()
            const fileName = file.getName().trim().toLowerCase()
            if (fileName === imageName + ".jpg" || fileName === imageName + ".png") {
                image = file
                break
            }
        }
        if (!image) return { error: "Image not found" }
        return { url: image.getUrl(), image: `=IMAGE("https://drive.google.com/uc?export=view&id=${image.getId()}")` }
    }

    updateImages() {
        const startTime = new Date().getTime()
        this.ss.toast("Running...", APP_NAME)

        const images = this.getImages()

        const ws = this.ss.getSheetByName(SN_MASTER)
        const dataRange = ws.getDataRange()
        const values = dataRange.getValues()
        values.shift()

        const urls1 = []
        const formulas1 = []
        const urls2 = []
        const formulas2 = []
        const urls3 = []
        const formulas3 = []
        const urls4 = []
        const formulas4 = []
        const urls5 = []
        const formulas5 = []

        const urls = [urls1, urls2, urls3, urls4, urls5]
        const formulas = [formulas1, formulas2, formulas3, formulas4, formulas5]

        values.forEach(v => {
            const folderName = v[this.indexFolderName].toString().trim().toUpperCase()

            const imageName1 = v[this.indexImageName1].toString().trim().toUpperCase()
            const imageName2 = v[this.indexImageName2].toString().trim().toUpperCase()
            const imageName3 = v[this.indexImageName3].toString().trim().toUpperCase()
            const imageName4 = v[this.indexImageName4].toString().trim().toUpperCase()
            const imageName5 = v[this.indexImageName5].toString().trim().toUpperCase()

            const key1 = `${folderName}${imageName1}`
            const key2 = `${folderName}${imageName2}`
            const key3 = `${folderName}${imageName3}`
            const key4 = `${folderName}${imageName4}`
            const key5 = `${folderName}${imageName5}`

            const imageNames = [imageName1, imageName2, imageName3, imageName4, imageName5]
            const keys = [key1, key2, key3, key4, key5]

            keys.forEach((key, i) => {
                const image = images[key]
                if (folderName && imageNames[i]) {
                    if (image) {
                        urls[i].push([image.downloadUrl])
                        formulas[i].push([image.formula])
                    } else {
                        urls[i].push([this.imageNotFound])
                        formulas[i].push([this.imageNotFound])
                    }
                } else {
                    urls[i].push([null])
                    formulas[i].push([null])
                }
            })
        })

        const indexUrls = [this.indexImageUrl1, this.indexImageUrl2, this.indexImageUrl3, this.indexImageUrl4, this.indexImageUrl5]
        const indexFormulas = [this.indexImage1, this.indexImage2, this.indexImage3, this.indexImage4, this.indexImage5]
        indexUrls.forEach((_, i) => {
            ws.getRange(2, indexUrls[i] + 1, urls[i].length, 1).setValues(urls[i])
            ws.getRange(2, indexFormulas[i] + 1, formulas[i].length, 1).setValues(formulas[i])
        })


        const endTime = new Date().getTime()
        const usedTime = Math.floor((endTime - startTime) / 1000)
        this.ss.toast(`Done. Used time in seconds ${usedTime}.`, APP_NAME)
    }

    onEdit(e) {
        const { rowStart, rowEnd, columnStart, columnEnd } = e.range
        const images = this.getImages()
        const ws = this.ss.getActiveSheet()
        const indexImageNames = [this.indexImageName1, this.indexImageName2, this.indexImageName3, this.indexImageName4, this.indexImageName5]
        const indexUrls = [this.indexImageUrl1, this.indexImageUrl2, this.indexImageUrl3, this.indexImageUrl4, this.indexImageUrl5]
        const indexFormulas = [this.indexImage1, this.indexImage2, this.indexImage3, this.indexImage4, this.indexImage5]
        const isValidRange = ws.getName() === SN_MASTER && rowStart > 1 && (
            (columnStart <= this.indexFolderName + 1 && columnEnd >= this.indexFolderName + 1) ||
            (columnStart <= this.indexImageName1 + 1 && columnEnd >= this.indexImageName1 + 1) ||
            (columnStart <= this.indexImageName2 + 1 && columnEnd >= this.indexImageName2 + 1) ||
            (columnStart <= this.indexImageName3 + 1 && columnEnd >= this.indexImageName3 + 1) ||
            (columnStart <= this.indexImageName4 + 1 && columnEnd >= this.indexImageName4 + 1) ||
            (columnStart <= this.indexImageName5 + 1 && columnEnd >= this.indexImageName5 + 1)
        )
        if (isValidRange) {
            this.ss.toast("Updating...", APP_NAME)
            for (let row = rowStart; row <= rowEnd; row++) {
                const folderName = ws.getRange(row, this.indexFolderName + 1).getValue().toString().toUpperCase().trim()
                const imageNames = indexImageNames.map(index => ws.getRange(row, index + 1).getValue().toString().toUpperCase().trim())
                imageNames.forEach((imageName, i) => {
                    const indexUrl = indexUrls[i]
                    const indexFormula = indexFormulas[i]
                    if (folderName && imageName) {
                        const key = `${folderName}${imageName}`
                        const image = images[key]
                        if (image) {
                            ws.getRange(row, indexUrl + 1).setValue(image.downloadUrl)
                            ws.getRange(row, indexFormula + 1).setValue(image.formula)
                        } else {
                            ws.getRange(row, indexUrl + 1).setValue(this.imageNotFound)
                            ws.getRange(row, indexFormula + 1).setValue(this.imageNotFound)
                        }
                    } else {
                        ws.getRange(row, indexUrl + 1).setValue(null)
                        ws.getRange(row, indexFormula + 1).setValue(null)
                    }
                })
            }
            this.ss.toast("Done", APP_NAME)
        }
    }
}


