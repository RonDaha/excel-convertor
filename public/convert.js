const convert = async (fileToRead) => {

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(fileToRead);
    const desiredWorkSheet = workbook._worksheets.filter((worksheet) => worksheet.name === 'Sponsored Products Campaigns')[0]

    const dataRows = []
    desiredWorkSheet.eachRow((row, rowNumber) => {

        if (rowNumber !== 1) {
            const newRow = []
            row.eachCell((cell, colNumber) => {
                newRow[5] = ''
                newRow[6] = ''
                newRow[7] = ''
                newRow[8] = ''
                newRow[9] = ''
                newRow[10] = ''
                newRow[14] = ''

                switch (colNumber) {
                    case 4: // Campaign
                        newRow[0] = cell._value.model.value
                        break
                    case 10: // Ad Group
                        newRow[1] = cell._value.model.value
                        break
                    case 11: // Max bid
                        newRow[2] = Number(cell._value.model.value)
                        break;
                    case 12: // Keyword or Product Targeting
                        newRow[3] = cell._value.model.value
                        break;
                    case 14: // Match Type
                        newRow[4] = cell._value.model.value
                        break;
                    case 19: // Impressions
                        newRow[11] = Number(cell._value.model.value)
                        break;
                    case 20: // Clicks
                        newRow[12] = Number(cell._value.model.value)
                        break;
                    case 21: // Spend
                        newRow[15] = Number(cell._value.model.value)
                        break;
                    case 22: // Order
                        newRow[16] = Number(cell._value.model.value)
                        break;
                    case 23: // Total Units
                        newRow[17] = Number(cell._value.model.value)
                        break;
                    case 24: // Sales
                        newRow[18] = Number(cell._value.model.value)
                        break;
                }
            })

            if (newRow[11] === 0 && newRow[12] === 0) {
                newRow[13] = 0
            } else {
                newRow[13] = newRow[12] / newRow[11]
            }

            if (newRow[15] === 0 && newRow[12] === 0) {
                newRow[14] = 0
            } else {
                newRow[14] =  newRow[15] / newRow[12]
            }

            if (newRow[16] === 0 && newRow[12] === 0) {
                newRow[19] = 0
            } else {
                newRow[19] = newRow[16] / newRow[12]
            }


            if ((newRow[15] === 0 && newRow[18] === 0) || newRow[18] === 0) {
                newRow[20] = 0
            } else {
                newRow[20] = newRow[15] / newRow[18]
            }

            // Type
            if (newRow[16] > 0 && newRow[20] < 0.3) {
                newRow[5] = 'Convertor'
            }
            else if (newRow[16] > 0 && newRow[20] > 0.3) {
                newRow[5] = 'Spender'
            }
            else if (newRow[16] === 0 && newRow[15] < 10) {
                newRow[5] = 'Bleeding in process'
            }
            else if (newRow[16] === 0 && newRow[15] > 10) {
                newRow[5] = 'Bleeder'
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
    summaryRow[20] = ''

    /* Header */
    const headerRow = [];
    headerRow[0] = 'Campaign';
    headerRow[1] = 'Ad Group';
    headerRow[2] = 'Max bid';
    headerRow[3] = 'Keyword or Product Targeting';
    headerRow[4] = 'Match Type';
    headerRow[5] = 'Type';
    headerRow[6] = 'Why';
    headerRow[7] = 'Bid';
    headerRow[8] = 'TOS';
    headerRow[9] = 'New Bid';
    headerRow[10] = 'TOS New';
    headerRow[11] = 'Impressions';
    headerRow[12] = 'Clicks';
    headerRow[13] = 'CTR';
    headerRow[14] = 'CPC';
    headerRow[15] = 'Spend';
    headerRow[16] = 'Order';
    headerRow[17] = 'Total Units';
    headerRow[18] = 'Sales';
    headerRow[19] = 'CVR';
    headerRow[20] = 'ACoS';

    /* insert new row and return as row object */
    sheet.insertRow(1, summaryRow);
    sheet.insertRow(2, headerRow);
    dataRows.forEach((rowToAdd) => {
        sheet.addRow(rowToAdd)
    })

    /* Paint */
    const firstRow = sheet.findRow(1)
    firstRow.eachCell((cell, colNumber) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'fff7caac' },
            bgColor: { argb: 'FF0000FF' }
        }
        cell.font = { bold: true }
    })

    const secondRow = sheet.findRow(2)
    secondRow.eachCell((cell, colNumber) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'ffb4c6e7' },
            bgColor: { argb: 'FF0000FF' }
        }
        cell.font = { bold: true }
    })

    const typeColumn = sheet.getColumn(6)
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

    sheet.getCell('L1').value = { formula: '=SUBTOTAL(9, L3:L9999)', date1904: false } // Imp
    sheet.getCell('M1').value = { formula: '=SUBTOTAL(9, M3:M9999)', date1904: false } // Clicks
    sheet.getCell('N1').value = { formula: '=L1/M1', date1904: false } // CTR
    sheet.getCell('O1').value = { formula: '=P1/M1', date1904: false } // CPC
    sheet.getCell('P1').value = { formula: '=SUBTOTAL(9, P3:P9999)', date1904: false } // Spend
    sheet.getCell('Q1').value = { formula: '=SUBTOTAL(9, Q3:Q9999)', date1904: false } // Orders
    sheet.getCell('R1').value = { formula: '=SUBTOTAL(9, R3:R9999)', date1904: false } // Total Units
    sheet.getCell('S1').value = { formula: '=SUBTOTAL(9, S3:S9999)', date1904: false } // Sales
    sheet.getCell('T1').value = { formula: '=Q1/M1', date1904: false } // CVR
    sheet.getCell('U1').value = { formula: '=P1/S1', date1904: false } // ACoS

    sheet.getColumn(14).numFmt = '0.00%' // CTR
    sheet.getColumn(15).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // CPC
    sheet.getColumn(16).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // Spend
    sheet.getColumn(19).numFmt = '"$"#,##0.00;[Red]\-"$"#,##0.00' // Sales
    sheet.getColumn(20).numFmt = '0.00%' // CVR
    sheet.getColumn(21).numFmt = '0.00%' // ACoS


    sheet.views = [
        { state: 'frozen', xSplit: 0, ySplit: 2 }
    ];

    return workbook2

}