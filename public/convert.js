const paint2TopRows = (sheet, colorFirstRow, colorSecondRow) => {
    const firstRow = sheet.findRow(1)
    firstRow.eachCell((cell, colNumber) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: colorFirstRow },
            bgColor: { argb: 'FF0000FF' }
        }
        cell.font = { bold: true }
    })

    const secondRow = sheet.findRow(2)
    secondRow.eachCell((cell, colNumber) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: colorSecondRow },
            bgColor: { argb: 'FF0000FF' }
        }
        cell.font = { bold: true }
    })
}
const paintTypeColumn = (sheet, typeColumnNumber) => {
    const typeColumn = sheet.getColumn(typeColumnNumber)
    typeColumn.eachCell((cell, rowNumber) => {
        if (rowNumber !== 1 && rowNumber !== 2) {
            let colorToFill
            switch (cell._value.model.value) {
                case 'Convertor':
                    colorToFill = 'ff00ff00'
                    break
                case 'Spender':
                    colorToFill = 'ff808080'
                    break
                case 'Bleeding in process':
                    colorToFill = 'ffffff00'
                    break
                case 'Bleeder':
                    colorToFill = 'ffff0000'
                    break
            }
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: colorToFill },
                bgColor: { argb: 'FF0000FF' }
            }
        }
    })
}

const convertSponsoredProductsCampaigns = (desiredWorkSheet) => {
    const dataRows = []
    desiredWorkSheet.eachRow((row, rowNumber) => {

        if (rowNumber !== 1) {
            const newRow = []
            row.eachCell((cell, colNumber) => {
                newRow[4] = ''
                newRow[5] = ''
                newRow[6] = ''
                newRow[7] = ''
                newRow[8] = ''
                newRow[9] = ''
                newRow[13] = ''

                switch (colNumber) {
                    case 4: // Campaign
                        newRow[0] = cell._value.model.value
                        break
                    case 10: // Ad Group
                        newRow[1] = cell._value.model.value
                        break
                    case 12: // Keyword or Product Targeting
                        newRow[2] = cell._value.model.value
                        break;
                    case 14: // Match Type
                        newRow[3] = cell._value.model.value
                        break;
                    case 11: // Max bid
                        newRow[6] = Number(cell._value.model.value)
                        break;
                    case 19: // Impressions
                        newRow[10] = Number(cell._value.model.value)
                        break;
                    case 20: // Clicks
                        newRow[11] = Number(cell._value.model.value)
                        break;
                    case 21: // Spend
                        newRow[14] = Number(cell._value.model.value)
                        break;
                    case 22: // Order
                        newRow[15] = Number(cell._value.model.value)
                        break;
                    case 23: // Total Units
                        newRow[16] = Number(cell._value.model.value)
                        break;
                    case 24: // Sales
                        newRow[17] = Number(cell._value.model.value)
                        break;
                }
            })

            if (newRow[10] === 0 && newRow[11] === 0) {
                newRow[12] = 0
            } else {
                newRow[12] = newRow[11] / newRow[10]
            }

            if (newRow[14] === 0 && newRow[11] === 0) {
                newRow[13] = 0
            } else {
                newRow[13] =  newRow[14] / newRow[11]
            }

            if ((newRow[15] === 0 && newRow[11] === 0 ) || newRow[11] === 0) {
                newRow[18] = 0
            } else {
                newRow[18] = newRow[15] / newRow[11]
            }


            if ((newRow[14] === 0 && newRow[17] === 0) || newRow[17] === 0) {
                newRow[19] = 0
            } else {
                newRow[19] = newRow[14] / newRow[17]
            }

            /* Type */
            if (newRow[15] > 0 && newRow[19] < 0.3) {
                newRow[4] = 'Convertor'
            }
            else if (newRow[15] > 0 && newRow[19] > 0.3) {
                newRow[4] = 'Spender'
            }
            else if (newRow[15] === 0 && newRow[14] < 10) {
                newRow[4] = 'Bleeding in process'
            }
            else if (newRow[15] === 0 && newRow[14] > 10) {
                newRow[4] = 'Bleeder'
            }


            dataRows.push(newRow)
        }

    })

    /* End */
    const workbook2 = new ExcelJS.Workbook()
    const sheet = workbook2.addWorksheet('Sponsored Products Campaigns', {});

    /* Summary */
    const summaryRow = [];
    summaryRow[0] = ''
    summaryRow[1] = ''
    summaryRow[2] = ''
    summaryRow[3] = ''
    summaryRow[4] = ''
    summaryRow[5] = ''
    summaryRow[6] = ''
    summaryRow[7] = ''
    summaryRow[8] = ''
    summaryRow[9] = ''
    summaryRow[10] = ''
    summaryRow[11] = ''
    summaryRow[12] = ''
    summaryRow[13] = ''
    summaryRow[14] = ''
    summaryRow[15] = ''
    summaryRow[16] = ''
    summaryRow[17] = ''
    summaryRow[18] = ''
    summaryRow[19] = ''

    /* Header */
    const headerRow = [];
    headerRow[0] = 'Campaign';
    headerRow[1] = 'Ad Group';
    headerRow[2] = 'Keyword or Product Targeting';
    headerRow[3] = 'Match Type';
    headerRow[4] = 'Type';
    headerRow[5] = 'Why';
    headerRow[6] = 'Max bid';
    headerRow[7] = 'TOS';
    headerRow[8] = 'New Bid';
    headerRow[9] = 'TOS New';
    headerRow[10] = 'Impressions';
    headerRow[11] = 'Clicks';
    headerRow[12] = 'CTR';
    headerRow[13] = 'CPC';
    headerRow[14] = 'Spend';
    headerRow[15] = 'Order';
    headerRow[16] = 'Total Units';
    headerRow[17] = 'Sales';
    headerRow[18] = 'CVR';
    headerRow[19] = 'ACoS';

    /* insert new row and return as row object */
    sheet.insertRow(1, summaryRow);
    sheet.insertRow(2, headerRow);
    dataRows.forEach((rowToAdd) => {
        sheet.addRow(rowToAdd)
    })

    /* Paint */
    paint2TopRows(sheet, 'fff7caac', 'ffb4c6e7')
    paintTypeColumn(sheet, 5)

    /* Formulas */
    sheet.getCell('K1').value = { formula: '=SUBTOTAL(9, K3:K9999)', date1904: false } // Imp
    sheet.getCell('L1').value = { formula: '=SUBTOTAL(9, L3:L9999)', date1904: false } // Clicks
    sheet.getCell('M1').value = { formula: '=L1/K1', date1904: false } // CTR
    sheet.getCell('N1').value = { formula: '=O1/L1', date1904: false } // CPC
    sheet.getCell('O1').value = { formula: '=SUBTOTAL(9, O3:O9999)', date1904: false } // Spend
    sheet.getCell('P1').value = { formula: '=SUBTOTAL(9, P3:P9999)', date1904: false } // Orders
    sheet.getCell('Q1').value = { formula: '=SUBTOTAL(9, Q3:Q9999)', date1904: false } // Total Units
    sheet.getCell('R1').value = { formula: '=SUBTOTAL(9, R3:R9999)', date1904: false } // Sales
    sheet.getCell('S1').value = { formula: '=P1/L1', date1904: false } // CVR
    sheet.getCell('T1').value = { formula: '=O1/R1', date1904: false } // ACoS
    /* Cell Types */
    sheet.getColumn(13).numFmt = '0.00%' // CTR
    sheet.getColumn(14).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // CPC
    sheet.getColumn(15).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // Spend
    sheet.getColumn(18).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // Sales
    sheet.getColumn(19).numFmt = '0.00%' // CVR
    sheet.getColumn(20).numFmt = '0.00%' // ACoS


    sheet.views = [
        { state: 'frozen', xSplit: 0, ySplit: 2 }
    ];

    return workbook2
}

const convertSponsoredProductsSearch = (desiredWorkSheet) => {

    const dataRows = []
    desiredWorkSheet.eachRow((row, rowNumber) => {

        if (rowNumber !== 1) {
            const newRow = []
            row.eachCell((cell, colNumber) => {

                newRow[7] = '' // Type
                newRow[10] = '' // CTR
                newRow[11] = '' // CPC
                newRow[16] = '' // CVR
                newRow[17] = '' // ACoS

                switch (colNumber) {
                    case 1: // Start Date
                        newRow[0] = cell._value.model.value
                        break
                    case 2: // End Date
                        newRow[1] = cell._value.model.value
                        break
                    case 5: // Campaign Name
                        newRow[2] = cell._value.model.value
                        break
                    case 6: // Ad Group Name
                        newRow[3] = cell._value.model.value
                        break
                    case 7: // Targeting
                        newRow[4] = cell._value.model.value
                        break
                    case 8: // Match Type
                        newRow[5] = cell._value.model.value
                        break
                    case 9: // Customer Search Term
                        newRow[6] = cell._value.model.value
                        break
                    case 10: // Impressions
                        newRow[8] = Number(cell._value.model.value)
                        break
                    case 11: // Clicks
                        newRow[9] = Number(cell._value.model.value)
                        break
                    case 14: // Spend
                        newRow[12] = Number(cell._value.model.value)
                        break
                    case 15: // Sales
                        newRow[13] = Number(cell._value.model.value)
                        break
                    case 18: // Orders
                        newRow[14] = Number(cell._value.model.value)
                        break
                    case 19: // Units
                        newRow[15] = Number(cell._value.model.value)
                        break

                }
            })

            /* CTR */
            if ((newRow[9] === 0 && newRow[8] === 0) || newRow[8] === 0) {
                newRow[10] = 0
            } else {
                newRow[10] = newRow[9] / newRow[8]
            }

            /* CPC */
            if ((newRow[12] === 0 && newRow[9] === 0) || newRow[9] === 0) {
                newRow[11] = 0
            } else {
                newRow[11] = newRow[12] / newRow[9]
            }

            /* CVR */
            if ((newRow[14] === 0 && newRow[9] === 0) || newRow[9] === 0) {
                newRow[16] = 0
            } else {
                newRow[16] = newRow[14] / newRow[9]
            }

            /* ACoS */
            if ((newRow[12] === 0 && newRow[13] === 0) || newRow[13] === 0) {
                newRow[17] = 0
            } else {
                newRow[17] = newRow[12] / newRow[13]
            }

            /* Type */
            if (newRow[14] > 0 && newRow[17] < 0.3) {
                newRow[7] = 'Convertor'
            }
            else if (newRow[14] > 0 && newRow[17] > 0.3) {
                newRow[7] = 'Spender'
            }
            else if (newRow[14] === 0 && newRow[12] < 10) {
                newRow[7] = 'Bleeding in process'
            }
            else if (newRow[14] === 0 && newRow[12] > 10) {
                newRow[7] = 'Bleeder'
            }
            dataRows.push(newRow)
        }

    })

    /* End */
    const workbook2 = new ExcelJS.Workbook()
    const sheet = workbook2.addWorksheet('Sponsored Product Search Term R', {});

    /* Summary */
    const summaryRow = [];
    summaryRow[0] = ''
    summaryRow[1] = ''
    summaryRow[2] = ''
    summaryRow[3] = ''
    summaryRow[4] = ''
    summaryRow[5] = ''
    summaryRow[6] = ''
    summaryRow[7] = ''
    summaryRow[8] = ''
    summaryRow[9] = ''
    summaryRow[10] = ''
    summaryRow[11] = ''
    summaryRow[12] = ''
    summaryRow[13] = ''
    summaryRow[14] = ''
    summaryRow[15] = ''
    summaryRow[16] = ''
    summaryRow[17] = ''

    /* Header */
    const headerRow = [];
    headerRow[0] = 'Start Date';
    headerRow[1] = 'End Date';
    headerRow[2] = 'Campaign Name';
    headerRow[3] = 'Ad Group Name';
    headerRow[4] = 'Targeting';
    headerRow[5] = 'Match Type';
    headerRow[6] = 'Customer Search Term';
    headerRow[7] = 'Type';
    headerRow[8] = 'Impressions';
    headerRow[9] = 'Clicks';
    headerRow[10] = 'CTR';
    headerRow[11] = 'CPC';
    headerRow[12] = 'Spend';
    headerRow[13] = 'Sales';
    headerRow[14] = 'Orders';
    headerRow[15] = 'Units';
    headerRow[16] = 'CVR';
    headerRow[17] = 'ACoS';

    sheet.insertRow(1, summaryRow);
    sheet.insertRow(2, headerRow);
    dataRows.forEach((rowToAdd) => {
        sheet.addRow(rowToAdd)
    })

    /* Paint */
    paint2TopRows(sheet, 'ffc5e0b3', 'ffbdd7ee')
    paintTypeColumn(sheet, 8)

    /* Formulas */
    sheet.getCell('I1').value = { formula: '=SUBTOTAL(9, I3:I99999)', date1904: false } // Imp
    sheet.getCell('J1').value = { formula: '=SUBTOTAL(9, J3:J99999)', date1904: false } // Clicks
    sheet.getCell('K1').value = { formula: '=J1/I1', date1904: false } // CTR
    sheet.getCell('L1').value = { formula: '=M1/J1', date1904: false } // CPC
    sheet.getCell('M1').value = { formula: '=SUBTOTAL(9, M3:M99999)', date1904: false } // Spend
    sheet.getCell('N1').value = { formula: '=SUBTOTAL(9, N3:N99999)', date1904: false } // Sales
    sheet.getCell('O1').value = { formula: '=SUBTOTAL(9, O3:O99999)', date1904: false } // Orders
    sheet.getCell('P1').value = { formula: '=SUBTOTAL(9, P3:P99999)', date1904: false } // Units
    sheet.getCell('Q1').value = { formula: '=J1/O1', date1904: false } // CVR
    sheet.getCell('R1').value = { formula: '=M1/N1', date1904: false } // ACoS
    /* Cell Types */
    sheet.getColumn(11).numFmt = '0.00%' // CTR
    sheet.getColumn(12).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // CPC
    sheet.getColumn(13).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // Spend
    sheet.getColumn(14).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // Sales
    sheet.getColumn(17).numFmt = '0.00%' // CVR
    sheet.getColumn(18).numFmt = '0.00%' // ACoS


    sheet.views = [
        { state: 'frozen', xSplit: 0, ySplit: 2 }
    ];

    return workbook2
}


const convert = async (fileToRead) => {

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(fileToRead);
    let reportType
    let desiredWorkSheet

    workbook._worksheets.forEach((worksheet) => {
        switch (worksheet.name) {
            case 'Sponsored Products Campaigns':
                reportType = 1
                desiredWorkSheet = worksheet
                break
            case 'Sponsored Product Search Term R':
                reportType = 2
                desiredWorkSheet = worksheet
                break
        }
    })
    const d = new Date().toLocaleDateString()
    switch (reportType) {
        case 1:
            return { workbook: convertSponsoredProductsCampaigns(desiredWorkSheet), name: `Sponsored Products Campaigns - analyzed ${d}.xlsx` }
        case 2:
            return { workbook: convertSponsoredProductsSearch(desiredWorkSheet), name: `Sponsored Product Search Term R - analyzed ${d}.xlsx` }
    }
}