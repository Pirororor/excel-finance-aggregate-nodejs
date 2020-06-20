'use strict';

const xlsx = require('XLSX');

const path = require('path');
const fs = require('fs');
var excel = require('excel4node');

const toCharCode = (a) => a.charCodeAt(0);
const toChar = (a) => String.fromCharCode(toCharCode('A') + a - 1);

const toCell = (row, col) => `${toChar(col)}${row}`;
const isNumeric = (str) => /^\d+$/.test(str);

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
                    return {
                        type: 'int',
                        key: name,
                    };
                }
                if (!isNumeric(name.charAt(0))){
                    return {
                        type: 'name',
                        key: name.toLowerCase(),
                    };
                }
                /** @type {String} */
                // const nameLc = name.toLowerCase();
                // let key = nameLc;
                // if(nameLc.includes('/')) {
                //     key = nameLc.substr(0, nameLc.indexOf('/'));
                // } else if(nameLc.includes(',')) {
                //     key =  nameLc.substr(0, nameLc.indexOf(','));
                // }
                return {
                    type: 'strnum',
                    key: name.toLowerCase().replace(/^0+/, ''),
                };
            };
            const { key } = getNameInfo(name);
            if (!this.#values[key] || typeof this.#values[key] === 'undefined') {
                this.#values[key] = {};
                Object.keys(this.#headers).forEach((col) => {
                    this.#values[key][col] = 0;
                });
            } else {
                console.log('duplicate detected:', filename, ' -- name:', name)
            }

            if (name === 'aab') {
                console.log(name, this.#filename);
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

        subDirectories.forEach((dirInfo) => {
            const offeringInfos = [];
            const files = fs.readdirSync(dirInfo.path)
                // exclude files that starts with '~'. (i.e opened files)
                .filter((x) => x.charAt(0) !== '~' );

            files.forEach((file) => {
                const offeringInfo = new OfferingInfo(dirInfo.path, file);
                offeringInfos.push(offeringInfo);
            });

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
                    headers.forEach((header) => {
                        const offeredVal = offeringInfo.getValue(header, key);
                        info.breakdown[header] += offeredVal;
                        info.total += offeredVal;
                        totalInfo[header] += offeredVal;
                    });
                    personInfo[key] = info;
                });
            });

            const monthInfo = {
                headers,
                personInfo,
                totalInfo,
            }
            outputInfo[dirInfo.name] = monthInfo;
        });

        const workbook = new excel.Workbook();
        Object.keys(outputInfo).forEach((month) => {
            const monthInfo = outputInfo[month];
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

            Object.keys(personInfo).forEach((personName) => {
                row += 1;
                col = 1;
                const nameCamelCase = personName.split(' ').map((x) => `${x.charAt(0).toUpperCase()}${x.substr(1)}`).join(' ');
                const prependZero = (name) => {
                    if (!isNumeric(name.charAt(0))) {
                        return name;
                    }

                    const nameFiltered = name
                        // replace all comma with front slash
                        .replace(/[,]/g, '/')
                        // split into array
                        .split('/')
                        // if length less than 3, prepend with zero until length is 3
                        .map((x) => {
                            const padded = `${x}`.padStart(3, '0');
                            return padded;
                        })
                        // join back to a string with forward slash
                        .join('/');

                    return nameFiltered;
                }
                worksheet.cell(row, col).string(prependZero(nameCamelCase));

                const { breakdown, total } = personInfo[personName];
                headers.forEach((header) => {
                    const val = breakdown[header];
                    col += 1;

                    if (typeof(val) === 'undefined') {
                        worksheet.cell(row, col).string('undefined');
                    } else {
                        worksheet.cell(row, col).number(val);
                    }
                });

                col+= 1;
                worksheet.cell(row, col).number(total);
            });

            row += 1;
            col = 1;
            worksheet.cell(row, col).string('Total');

            headers.forEach((header) => {
                col += 1;
                worksheet.cell(row, col).number(totalInfo[header]);
            });
        });

        workbook.write(path.join(__dirname, 'out.xlsx'));

    } catch (err) {
        console.error(err);
    }
};

main();
