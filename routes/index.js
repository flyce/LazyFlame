var express = require('express');
var router = express.Router();

if (typeof require !== 'undefined') XLSX = require('xlsx');

router.get('/', function (req, res, next) {
    readXlsx().then((xlsxData) => {
        const sheetData = xlsxData.Sheets.sheet1;
        let tableLength = sheetData['!ref'].match(/\d{2,}/)[0];
        let currentCenter = null, currentCenterStartNumber = null;
        let outputData = new Array();
        for (let index = 1; index < tableLength; index++) {
            const element = sheetData["A" + index];
            if (element.v.match(/[\u4e2d][\u5fc3]/) != null) { // 匹配到 "中心" 才进行处理，否则不是我需要的单元格
                if (currentCenter == null) {
                    currentCenter = element.v;
                }

                if (currentCenterStartNumber == null) {
                    currentCenterStartNumber = sheetData["B" + index].v;
                }

                // 中心发生变化时 改变当前中心和当前中心初识序号
                if (currentCenter != element.v) {
                    currentCenter = element.v;
                    currentCenterStartNumber = sheetData["B" + index].v;
                }

                let currentCenterItemsNumber = sheetData["B" + index].v - currentCenterStartNumber + 1;
                let cellThree = sheetData["C" + index].v;
                let cellFour = sheetData["D" + index] == null ? null : sheetData["D" + index].v;

                let str = currentCenterItemsNumber + "." +
                    combineCellThreeAndCellFour(cellThree, cellFour,
                        // 判断当前行和下一行的中心是否一直，从而决定结束的符号
                        element.v == sheetData["A" + ((index + 1) < tableLength ? index + 1 : tableLength)].v
                    );
                outputData[currentCenter] = outputData[currentCenter] == null ? str + "\n" : outputData[currentCenter] +  str + "\n";


            }
        }
        res.json({
            "工程技术中心": outputData["工程技术中心"],
            "航材保障中心": outputData["航材保障中心"],
            "计划与控制中心": outputData["计划与控制中心"],
            "安全质量中心": outputData["安全质量中心"],
            "培训管理中心": outputData["培训管理中心"],
            "飞机维修中心": outputData['飞机维修中心']
        });
    })
});

/***
 * 合并第三列和第四列的值
 * @param cellThree
 * @param cellFour
 * @param endFlag
 * @returns {string}
 */
function combineCellThreeAndCellFour(cellThree, cellFour, endFlag) {
    cellThree = cellThree.replace(/[\u3002]/, '');
    if (cellFour != null) {
        cellFour = cellFour.toString().replace(/[\u3002]$/, '');
    }
    let endMark = endFlag ? ';' : '。';
    if (handleSemanteme(cellFour)) {
        return cellThree + endMark;
    }

    return cellThree + "，" + cellFour + endMark;
}


/***
 * 处理第三列和第四列的内容，让语意更自然，如果第四列中有下面三种情况，则需要删除第四列的内容
 * @param cellFour
 * @returns {boolean}
 */
function handleSemanteme(cellFour) {
    const meaning = ["跟进执行", "持续进行", "已完成", "跟进", "完成", "准备中", "进行中"];
    // 判断第四列的内容是否和meaning数组中某一个值一致
    for (let index = 0; index < meaning.length; index ++) {
        if (cellFour == meaning[index]) {
            return true;
        }
    }

    if (cellFour == '' || cellFour == null) {
        return true;
    }

    // 判断第四列的内容四否为 201*年**月**日
    if (cellFour.match(/[\d]{2,}[\u5e74][\d]{1,}[\u6708][\d]{1,}[\u65e5]/) && cellFour.length <= 11) {
        return true;
    }

    // 判断第四列的值是否为Excel的时间戳
    if (cellFour.match(/[\d]{5}/)) {
        return true;
    }

    return false;
}

async function readXlsx() {
    return await XLSX.readFile('./public/doc.xlsx')
}

module.exports = router;
