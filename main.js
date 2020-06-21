'use strict';

const xlsx = require('XLSX');

const path = require('path');
const fs = require('fs');
var excel = require('excel4node');

const toCharCode = (a) => a.charCodeAt(0);
const toChar = (a) => String.fromCharCode(toCharCode('A') + a - 1);

const toCell = (row, col) => `${toChar(col)}${row}`;
const isNumeric = (str) => /^\d+$/.test(str);

const convertNameForOutput = (name) => {
    let nameCamelCase = name.split(' ').map((x) => `${x.charAt(0).toUpperCase()}${x.substr(1)}`).join(' ');
    if (!isNumeric(nameCamelCase.charAt(0))) {
        return nameCamelCase;
    }

    const nameFiltered = nameCamelCase
        // split into array
        .split('/')
        // if length less than 3, prepend with zero until length is 3
        .map(prependZero)

        // join back to a string with forward slash
        .join('/');

    return nameFiltered;
};

const aggregateInfo = (offeringInfos) => {
    const headerNameMap = {};
    offeringInfos.forEach((offeringInfo) => {
        const headers = offeringInfo.getHeaderNames();
        headers.forEach(header => {
            headerNameMap[header] = true;
        });
    });

    const totalInfo = {};
    const personInfo = {};
    const headers = Object.keys(headerNameMap);
    let grandTotal = 0;
    headers.forEach(header => {
        totalInfo[header] = 0;
    });
    offeringInfos.forEach((offeringInfo) => {
        offeringInfo.forEachPerson((key, value) => {
            let info = personInfo[key];
            if (!info) {
                info = {
                    total: 0,
                    breakdown: {},
                };
                headers.forEach((header)=> {
                    info.breakdown[header] = 0;
                });
            }

            // if (offeringInfo.getFileName() === '06 Jan 2019.xls') {
            //     console.log(key, offeringInfo.getFileName(), typeof key);
            // }
            // if (key === '3') {
            //     console.log('before', info.total, info.breakdown);
            // }
            headers.forEach((header) => {
                const offeredVal = offeringInfo.getValue(header, key);
                info.breakdown[header] += offeredVal;
                info.total += offeredVal;
                totalInfo[header] += offeredVal;
                grandTotal += offeredVal;
            });

            // if (key === '3') {
            //     console.log('after', info.total, info.breakdown);
            // }
            personInfo[key] = info;
        });
    });

    const monthInfo = {
        headers,
        personInfo,
        totalInfo,
        grandTotal,
    }

    // console.log(monthInfo);
    return monthInfo;
};

const outputExcelWorkSheet = (workbook, month, monthInfo) => {
        const worksheet = workbook.addWorksheet(month);
        const style = workbook.createStyle({
            font: {
                color: '#000000',
                size: 10,
            }
        });
        const {
            headers,
            personInfo,
            totalInfo,
            grandTotal,
        } = monthInfo;

        let col = 1;
        let row = 1;
        worksheet.cell(row, col).string('编号 No.');
        headers.forEach( (header) => {
            col += 1;
            worksheet.cell(row, col).string(header);
        });
        col += 1;
        worksheet.cell(row, col).string('Total');

        const personInfoOutput = {};
        const personNames = Object.keys(personInfo);
        personNames.forEach((name) => {
            const outputName = convertNameForOutput(name);
            personInfoOutput[outputName] = personInfo[name];
        });
        Object.keys(personInfoOutput).sort().forEach((personName) => {
            row += 1;
            col = 1;
            worksheet.cell(row, col).string(personName);

            const { breakdown, total } = personInfoOutput[personName];
            headers.forEach((header) => {
                const val = breakdown[header];
                col += 1;

                if (typeof(val) === 'undefined') {
                    worksheet.cell(row, col).string('undefined');
                } else {
                    worksheet.cell(row, col).number(val);
                }
            });

            // total for the row
            col+= 1;
            worksheet.cell(row, col).number(total);
        });

        // total for each column
        row += 1;
        col = 1;
        worksheet.cell(row, col).string('Total');

        headers.forEach((header) => {
            col += 1;
            worksheet.cell(row, col).number(totalInfo[header]);
        });

        // grand total
        col += 1;
        worksheet.cell(row, col).number(grandTotal);
};

/**
 * @param {String} x
 */
const prependZero = (x) => {
    if (isNumeric(x.charAt(x.length - 1))) {
        return `${x}`.padStart(3, '0');
    }
    return `${x.substring(0, x.length - 1).padStart(3, '0')}${x.substring(x.length - 1)}`;
}

class ExcelReader {
    #file = null
    constructor(filename) {
        this.#file = xlsx.readFile(filename);
    }

    getCellVal(row, col) {
        const cell = toCell(row, col);
        const cellInfo = this.#file.Sheets.wformula[cell];
        if (!cellInfo || typeof cellInfo === 'undefined') {
            return null;
        }
        return cellInfo.v;
    }
}

class OfferingInfo {
    #filename
    #values = {}
    #headers = {}
    #headerNameToCol = {};

    constructor(filepath, filename) {
        this.#filename = filename;
        const file = new ExcelReader(path.join(filepath, filename));

        let col = 2;
        while (true) {
            const val = file.getCellVal(2, col);
            if (val === "总数 Total") {
                break;
            }
            this.#headers[col] = val;
            this.#headerNameToCol[val] = col;
            col += 1;
        }

        const values = {};
        let row = 3;
        while (true) {
            const name = file.getCellVal(row, 1);
            if (!name) {
                break;
            }

            const getNameInfo = (name) => {
                if (isNumeric(name)) {
                    if (name === '003') {
                        console.log('numeric', typeof name);
                    }
                    return {
                        type: 'int',
                        key: `${parseInt(name, 10)}`,
                    };
                }
                if (!isNumeric(name.charAt(0))){
                    return {
                        type: 'name',
                        key: name.toLowerCase(),
                    };
                }
                return {
                    type: 'strnum',
                    key: name.toUpperCase().replace(/,/g, '/').split('/').map(x => x.replace(/^0+/, '')).join('/'),
                };
            };
            const { key } = getNameInfo(name);
            if (name === '003') {
                console.log( 'name', name, 'key', key);
            }
            if (!this.#values[key] || typeof this.#values[key] === 'undefined') {
                this.#values[key] = {};
                Object.keys(this.#headers).forEach((col) => {
                    this.#values[key][col] = 0;
                });
            } else {
                console.log('duplicate detected:', filename, ' -- name:', name)
            }

            Object.keys(this.#headers).forEach((col) => {
                const cellVal = file.getCellVal(row, parseInt(col, 10)) || 0;
                this.#values[key][col] += cellVal;
            });

            row += 1;
        }
    }

    getHeaderNames () {
        return Object.values(this.#headers);
    }


    forEachPerson (func) {
        Object.keys(this.#values).forEach((key) => {
            const values = this.#values[key];
            func(key, values);
        });
    }

    getValue(headerName, key) {
        const idx = this.#headerNameToCol[headerName];
        if (!idx) {
            return 0;
        }
        return this.#values[key][idx];
    }

    getFileName () {
        return this.#filename;
    }

}

const main = async () => {

    try {
        const basePath = path.join(__dirname, 'input');
        const getDirectories = source =>
                        fs.readdirSync(source, { withFileTypes: true })
                        .filter(dirent => dirent.isDirectory())
                        .map(dirent => dirent.name)

        const subDirectories = getDirectories(basePath).map((folder) => ({
            path: path.join(basePath, folder),
            name: folder
        }));

        const outputInfo = {
        };

        const allOfferingInfos = [];
        subDirectories.forEach((dirInfo) => {
            const offeringInfos = [];
            const files = fs.readdirSync(dirInfo.path)
                // exclude files that starts with '~'. (i.e opened files)
                .filter((x) => x.charAt(0) !== '~' );

            files.forEach((file) => {
                const offeringInfo = new OfferingInfo(dirInfo.path, file);
                offeringInfos.push(offeringInfo);
                allOfferingInfos.push(offeringInfo);
            });
            outputInfo[dirInfo.name] = aggregateInfo(offeringInfos);
        });

        const totalInfo = aggregateInfo(allOfferingInfos);
        const workbook = new excel.Workbook();

        const months = Object.keys(outputInfo).sort((a, b) => {
            const priority = (x) => {
                switch (x.toLowerCase()) {
                    case 'jan':
                    case 'January':
                        return 1;
                    case 'feb':
                    case 'february':
                        return 2;
                    case 'mar':
                    case 'march':
                        return 3;
                    case 'apr':
                    case 'april':
                        return 4;
                    case 'may':
                        return 5;
                    case 'jun':
                    case 'june':
                        return 6;
                    case 'jul':
                    case 'july':
                        return 7;
                    case 'aug':
                    case 'august':
                        return 8;
                    case 'sep':
                    case 'september':
                        return 9;
                    case 'oct':
                    case 'october':
                        return 10;
                    case 'nov':
                    case 'november':
                        return 11;
                    case 'dec':
                    case 'december':
                        return 12;
                    default:
                    return x.toLowerCase();;
                }
            };
            return priority(a) - priority(b);
        });
        months.forEach((month) => {
            const monthInfo = outputInfo[month];
            outputExcelWorkSheet(workbook, month, monthInfo);
        });
        outputExcelWorkSheet(workbook, 'Total', totalInfo);

        workbook.write(path.join(__dirname, 'out.xlsx'));

    } catch (err) {
        console.error(err);
    }
};

main();
