const fs = require('fs');
const excel = require('excel4node');
const crypto = require('crypto');
const nodePath = require('path');
const homedir = require('os').homedir();

const workBook = new excel.Workbook();
const workSheet = workBook.addWorksheet("Inventorization", {});

const pathSelect = document.getElementById("choose-path");
const pathTextInput = document.getElementById("path-text");
const fileNameTextInput = document.getElementById("exel-name");
const generateButton = document.getElementById("button");
const pathErrorText = document.querySelector('.path-error-text');
const fileNameErrorText = document.querySelector('.file-name-error-text');
const checkboxes = document.querySelector('.checkboxes');
const progressBar = document.querySelector('.progress-wrapper');
const progressTextPercent = document.querySelector('.progress-percents');
const progressFill = document.querySelector('.animated-progress span');
// checkboxes values
const sizeCheckbox = document.getElementById("size-checkbox");
const md5Checkbox = document.getElementById("MD5-checkbox");
const shaCheckbox = document.getElementById("SHA-checkbox");
const createDateCheckbox = document.getElementById("create-date-checkbox");
const modifyDateCheckbox = document.getElementById("modify-date-checkbox");
const accessDateCheckbox = document.getElementById("access-date-checkbox");

let rowNumber;
let selectedDirectory;
let numberOfFiles;
let percentOfProgress = 0;

pathSelect.addEventListener("click", () => {
    window.postMessage({
        type: 'select-dirs',
    })
})

pathTextInput.addEventListener("input", (event) => {
    selectedDirectory = event.target.value;
    pathErrorText.innerText = "";
});

const startGenerate = () => {
    checkboxes.style.display = "none";
    progressBar.style.display = "block";
    generateButton.disabled = true;
    fileNameTextInput.disabled = true;
    pathSelect.disabled = true;
}

const endGenerate = () => {
    checkboxes.style.display = "block";
    progressBar.style.display = "none";
    generateButton.disabled = false;
    fileNameTextInput.disabled = false;
    pathSelect.disabled = false;
    fileNameTextInput.value = '';
    pathTextInput.value = '';
}

const countFilesInDir = (directory) => {
    const filesInDirectory = fs.readdirSync(directory);
    for (const file of filesInDirectory) {
        const absolute = nodePath.join(directory, file);
        try {
            const isDir = fs.statSync(absolute).isDirectory()
            numberOfFiles++;
            isDir && countFilesInDir(absolute);
        } catch (err) {}
    }
};

const generateExel = () => {
    if (!pathTextInput.value || !fileNameTextInput.value) {
        if (!pathTextInput.value) pathErrorText.innerText = "Select directory";
        if (!fileNameTextInput.value) fileNameErrorText.innerText = "Enter a name";
        return;
    }
    numberOfFiles = 0;
    rowNumber = 1;
    startGenerate();
    setTimeout(() => {
        countFilesInDir(selectedDirectory);
        writeColumnTitles();
        writeCell(selectedDirectory);
    });
}

generateButton.addEventListener("click", generateExel);

// Exel
const initialColumnsTitles = ["№", "type", "dir or filename", "path", "extensions"]
let columnsTitles = [...initialColumnsTitles];
let columnTitlesMap = {};

const titlesStyle = workBook.createStyle({
    fill: {
        type: 'pattern',
        patternType: 'solid',
        fgColor: "#FFD200",
    },
    font: {
        bold: true,
        size: 12,
    },
    alignment: {
        horizontal: 'center',
    }
})
const dirStyles = workBook.createStyle({
    fill: {
        type: 'pattern',
        patternType: 'solid',
        fgColor: "#FDE9D9",
    },
});
const emptyStyles = workBook.createStyle({});

const addZero = (num) => {
    const str = num.toString();
    if (str.length === 1) {
        return `0${num}`;
    }
    return num;
}

const dateToFormat = (date) => {
    const seconds = addZero(date.getSeconds());
    const minutes = addZero(date.getMinutes());
    const hours = addZero(date.getHours());
    const day = addZero(date.getDate());
    const month = addZero(date.getMonth() + 1);
    const year = addZero(date.getFullYear());
    return `${hours}:${minutes}:${seconds}  ${day}/${month}/${year}`;

}

function writeColumnTitles () {
    sizeCheckbox.checked && columnsTitles.push("size");
    md5Checkbox.checked && columnsTitles.push("MD5");
    shaCheckbox.checked && columnsTitles.push("SHA1");
    createDateCheckbox.checked && columnsTitles.push("CreateDate");
    modifyDateCheckbox.checked && columnsTitles.push("ModifyDate");
    accessDateCheckbox.checked && columnsTitles.push("AccessDate");

    for (let i = 0; i <= columnsTitles.length; i++) {
        workSheet.cell(1, i + 1).string(columnsTitles[i]).style(titlesStyle);
        columnTitlesMap[columnsTitles[i]] = i + 1;
    }
    workSheet.row(1).filter();
}

function writeCell (path) {
    const currentRowNumber = rowNumber++;
    const currentPercentOfProgress = Math.floor((currentRowNumber - 1) / numberOfFiles * 100);
    if (currentPercentOfProgress > percentOfProgress) {
        percentOfProgress = currentPercentOfProgress;
        progressTextPercent.innerText = `progress: ${currentPercentOfProgress} %, number of files: ${numberOfFiles}`;
        progressFill.style.width = `${currentPercentOfProgress}%`;
    }
    fs.stat(path, (err, stats) => {
        if (err) {
            workSheet.cell(currentRowNumber + 1, columnTitlesMap["№"]).number(currentRowNumber);
            workSheet.cell(currentRowNumber + 1, columnTitlesMap["dir or filename"]).string(nodePath.basename(path));
            workSheet.cell(currentRowNumber + 1, columnTitlesMap["path"]).string(nodePath.dirname(path));
            workSheet.cell(currentRowNumber + 1, columnTitlesMap["size"]).number(err.message);
            return;
        }
        const isDir = stats.isDirectory();

        workSheet.cell(currentRowNumber + 1, columnTitlesMap["№"]).number(currentRowNumber);
        workSheet.cell(currentRowNumber + 1, columnTitlesMap["dir or filename"]).string(nodePath.basename(path)).style(isDir ? dirStyles : emptyStyles);
        workSheet.cell(currentRowNumber + 1, columnTitlesMap["type"]).string(isDir ? "directory" : "file").style(isDir ? dirStyles : emptyStyles);
        workSheet.cell(currentRowNumber + 1, columnTitlesMap["path"]).string(nodePath.dirname(path)).style(isDir ? dirStyles : emptyStyles);

        columnsTitles.includes("size") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["size"]).number(stats.size).style(isDir ? dirStyles : emptyStyles);
        columnsTitles.includes("CreateDate") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["CreateDate"]).string(dateToFormat(stats.ctime)).style(isDir ? dirStyles : emptyStyles);

        if (isDir) {
            columnsTitles.includes("ModifyDate") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["ModifyDate"]).string(``).style(dirStyles);
            columnsTitles.includes("AccessDate") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["AccessDate"]).string(``).style(dirStyles);
            columnsTitles.includes("extensions") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["extensions"]).string('').style(dirStyles);
            columnsTitles.includes("SHA1") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["SHA1"]).string('').style(dirStyles);
            columnsTitles.includes("MD5") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["MD5"]).string('').style(dirStyles);
            fs.readdir(path, (err, files) => {
                files.forEach((file) => {
                    writeCell(nodePath.join(path, file));
                })
            })
        } else {
            columnsTitles.includes("ModifyDate") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["ModifyDate"]).string(dateToFormat(stats.mtime));
            columnsTitles.includes("AccessDate") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["AccessDate"]).string(dateToFormat(stats.atime));
            columnsTitles.includes("extensions") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["extensions"]).string(nodePath.extname(path));

            if (stats.size > 2147483646) {
                columnsTitles.includes("SHA1") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["SHA1"]).string("File is too big");
            } else {
                try {
                    const fileData = fs.readFileSync(path);
                    columnsTitles.includes("SHA1") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["SHA1"]).string(crypto.createHash('sha1').update(fileData).digest('base64'));
                    columnsTitles.includes("MD5") && workSheet.cell(currentRowNumber + 1, columnTitlesMap["MD5"]).string(crypto.createHash('md5').update(fileData).digest('base64'));
                } catch (err) {
                    console.log(err);
                }
            }
        }
        if (currentRowNumber - 1 === numberOfFiles) {
            writeWorkBook();
        }
    });
}

function writeWorkBook () {
    workSheet.column(columnTitlesMap["dir or filename"]).setWidth(50);
    workSheet.column(columnTitlesMap["path"]).setWidth(50);
    workSheet.column(columnTitlesMap["extensions"]).setWidth(12);
    columnsTitles.includes("size") && workSheet.column(columnTitlesMap["size"]).setWidth(12);
    columnsTitles.includes("MD5") && workSheet.column(columnTitlesMap["MD5"]).setWidth(30);
    columnsTitles.includes("SHA1") && workSheet.column(columnTitlesMap["SHA1"]).setWidth(35);
    columnsTitles.includes("ModifyDate") && workSheet.column(columnTitlesMap["ModifyDate"]).setWidth(25);
    columnsTitles.includes("AccessDate") && workSheet.column(columnTitlesMap["AccessDate"]).setWidth(25);
    columnsTitles.includes("CreateDate") && workSheet.column(columnTitlesMap["CreateDate"]).setWidth(25);

    workBook.write(nodePath.join(homedir, "Desktop", `${fileNameTextInput.value.trim()}.xlsx`), (err, stats) => {
        endGenerate();
        if (err) {
            console.log(err);
        } else {
            columnTitlesMap = {};
            columnsTitles = [...initialColumnsTitles];
        }
    });
}
