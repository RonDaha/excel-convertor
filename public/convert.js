const convert = async (fileToRead) => {

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(fileToRead);
    const desiredWorkSheet = workbook._worksheets.filter((worksheet) => worksheet.name === 'Sponsored Products Campaigns')[0]

    const dataRows = []
    const impSubTotal = []
    const clicksSubTotal = []
    const spendSubTotal = []
    const ordersSubTotal = []
    const totalUnitsSubTotal = []
    const salesSubTotal = []
    const acosSubTotal = []
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
                newRow[14] = '' // CPC TODO

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
                        impSubTotal.push(newRow[11])
                        break;
                    case 20: // Clicks
                        newRow[12] = Number(cell._value.model.value)
                        clicksSubTotal.push(newRow[12])
                        break;
                    case 21: // Spend
                        newRow[15] = Number(cell._value.model.value)
                        spendSubTotal.push(newRow[15])
                        break;
                    case 22: // Order
                        newRow[16] = Number(cell._value.model.value)
                        ordersSubTotal.push(newRow[16])
                        break;
                    case 23: // Total Units
                        newRow[17] = Number(cell._value.model.value)
                        totalUnitsSubTotal.push(newRow[17])
                        break;
                    case 24: // Sales
                        newRow[18] = Number(cell._value.model.value)
                        salesSubTotal.push(newRow[18])
                        break;
                }
            })

            // CTR TODO - add '%'
            if (newRow[11] === 0 && newRow[12] === 0) {
                newRow[13] = 0
            } else {
                newRow[13] = newRow[12] / newRow[11]
            }

            // CPC
            if (newRow[15] === 0 && newRow[12] === 0) {
                newRow[14] = '$ 0'
            } else {
                newRow[14] =  '$ ' + String(newRow[15] / newRow[12])
            }

            // CVR TODO - add '%' (Same as CPC)
            // if (newRow[15] === 0 && newRow[12] === 0) {
            //     newRow[14] = '$ 0'
            // } else {
            //     newRow[14] =  '$ ' + String(newRow[15] / newRow[12])
            // }
            // headerRow[19]

            // ACos TODO - add '%'
            if ((newRow[15] === 0 && newRow[18] === 0) || newRow[18] === 0) {
                newRow[20] = 0
            } else {
                newRow[20] = newRow[15] / newRow[18]
            }
            acosSubTotal.push(newRow[20])

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

            // Convert 'Spend', 'Sales' to string
            newRow[15] = '$ ' + String(newRow[15])
            newRow[18] = '$ ' + String(newRow[18])
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
    summaryRow[11] = Number(impSubTotal.reduce((a, b) => a + b, 0) / impSubTotal.length)
    summaryRow[12] = Number(clicksSubTotal.reduce((a, b) => a + b, 0) / clicksSubTotal.length)
    summaryRow[13] = ''
    summaryRow[14] = ''
    summaryRow[15] = Number(spendSubTotal.reduce((a, b) => a + b, 0) / spendSubTotal.length)
    summaryRow[16] = Number(ordersSubTotal.reduce((a, b) => a + b, 0) / ordersSubTotal.length)
    summaryRow[17] = Number(totalUnitsSubTotal.reduce((a, b) => a + b, 0) / totalUnitsSubTotal.length)
    summaryRow[18] = Number(salesSubTotal.reduce((a, b) => a + b, 0) / salesSubTotal.length)
    summaryRow[19] = ''
    summaryRow[20] = Number(acosSubTotal.reduce((a, b) => a + b, 0) / acosSubTotal.length)

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
    headerRow[13] = 'CTR'; // TODO top row
    headerRow[14] = 'CPC'; // TODO top row
    headerRow[15] = 'Spend';
    headerRow[16] = 'Order';
    headerRow[17] = 'Total Units';
    headerRow[18] = 'Sales';
    headerRow[19] = 'CVR'; // TODO top row
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
    })

    const secondRow = sheet.findRow(2)
    secondRow.eachCell((cell, colNumber) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'ffb4c6e7' },
            bgColor: { argb: 'FF0000FF' }
        }
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

    return workbook2

}