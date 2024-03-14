import Excel from 'exceljs';
import FileSaver from 'file-saver'
import { setConditions } from './export/setConditions'
import { setDataValidation } from './export/setDataValidation'
import { createCellRange } from './export/utils'
 
const exportExcel = (luckysheet, value) => {
  console.log(luckysheet,'sheet数组')
  console.log(luckysheet[0].dataVerification,'sheet1中具体操作')
  // 参数为luckysheet.getluckysheetfile()获取的对象
  // 1.创建工作簿，可以为工作簿添加属性
  const workbook = new Excel.Workbook()
  // 2.创建表格，第二个参数可以配置创建什么样的工作表
  if (Object.prototype.toString.call(luckysheet) === '[object Object]') {
    luckysheet = [luckysheet]
  }
  //遍历sheet，将luckysheet的sheet转换成excel的sheet
  luckysheet.forEach(async function(thesheet) {
    // thesheet为sheet对象数据
    if (thesheet.data.length === 0) return  true
    const worksheet = workbook.addWorksheet(thesheet.name)
    const merge = (thesheet.config && thesheet.config.merge) || {}
    const borderInfo = (thesheet.config && thesheet.config.borderInfo) || {}
    // 3.设置导出操作
    setStyleAndValue(thesheet.data, worksheet)
    setMerge(merge, worksheet)
    setBorder(borderInfo, worksheet)
    setImages(thesheet, worksheet, workbook);
    setHyperlink(thesheet.hyperlink,worksheet)
    setFrozen(thesheet.frozen,worksheet)
    setConditions(thesheet.luckysheet_conditionformat_save,worksheet)
    setFilter(thesheet.filter_select,worksheet)
    setDataValidation(thesheet.dataVerification,worksheet)
    //开启显示数据透视表
    if(thesheet.isPivotTable){
      worksheet.pivotTables = true
      let rmax = 0
      let cmax = 0
      //得到行与列的最大值
      thesheet.celldata.forEach(itemCell=>{
        if(rmax<itemCell.r) rmax = itemCell.r
        if(cmax<itemCell.c) cmax = itemCell.c
      })
      // 循环遍历添加边框
      for(let i=0;i<=rmax;i++){
        for(let j=0;j<=cmax;j++){
          // 添加边框
          worksheet.getCell(i+1,j+1).border = {
            top: {style:'thin'},
            left: {style:'thin'},
            bottom: {style:'thin'},
            right: {style:'thin'}
          }
        }
      }
    }
    return true
  })
 
  // return
  // 4.写入 buffer
  const buffer = workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], {
      type: 'application/vnd.ms-excel;charset=utf-8'
    })
    //保存文件
    FileSaver.saveAs(blob, `${value}.xlsx`)
  })
  return buffer
}
//设置合并数据
var setMerge = function(luckyMerge = {}, worksheet) {
  const mergearr = Object.values(luckyMerge)
  mergearr.forEach(function(elem) {
    // elem格式：{r: 0, c: 0, rs: 1, cs: 2}
    // 按开始行，开始列，结束行，结束列合并（相当于 K10:M12）
    worksheet.mergeCells(
      elem.r + 1,
      elem.c + 1,
      elem.r + elem.rs,
      elem.c + elem.cs
    )
  })
}
//重新设置边框
var setBorder = function(luckyBorderInfo, worksheet) {
  if (!Array.isArray(luckyBorderInfo)) return
  luckyBorderInfo.forEach(function(elem) {
    // 现在只兼容到borderType 为range的情况
    if (elem.rangeType === 'range') {
      let border = borderConvert(elem.borderType, elem.style, elem.color)
      let rang = elem.range[0]
      let row = rang.row
      let column = rang.column
      for (let i = row[0] + 1; i < row[1] + 2; i++) {
        for (let y = column[0] + 1; y < column[1] + 2; y++) {
          worksheet.getCell(i, y).border = border
        }
      }
    }
    if (elem.rangeType === 'cell') {
      // col_index: 2
      // row_index: 1
      // b: {
      //   color: '#d0d4e3'
      //   style: 1
      // }
      const { col_index, row_index } = elem.value
      const borderData = Object.assign({}, elem.value)
      delete borderData.col_index
      delete borderData.row_index
      let border = addborderToCell(borderData, row_index, col_index)
      worksheet.getCell(row_index + 1, col_index + 1).border = border
    }
  })
}
// 设置单元格样式和值
var setStyleAndValue = function (cellArr, worksheet) {
  if (!Array.isArray(cellArr)) return;
 
  cellArr.forEach(function (row, rowid) {
    const dbrow = worksheet.getRow(rowid + 1);
    //设置单元格行高,默认乘以0.8倍
    dbrow.height = luckysheet.getRowHeight([rowid])[rowid] * 0.8;
    row.every(function (cell, columnid) {
      if (!cell) return true;
      if (rowid == 0) {
        const dobCol = worksheet.getColumn(columnid + 1);
        //设置单元格列宽除以8
        dobCol.width = luckysheet.getColumnWidth([columnid])[columnid] / 8;
      }
      let fill = fillConvert(cell.bg);
      let font = fontConvert(
        cell.ff||'Times New Roman',
        cell.fc,
        cell.bl,
        cell.it,
        cell.fs,
        cell.cl,
        cell.ul
      );
      let alignment = alignmentConvert(cell.vt, cell.ht, cell.tb, cell.tr);
      let value;
 
      var v = "";
      if (cell.ct && cell.ct.t == "inlineStr") {
        var s = cell.ct.s;
        s.forEach(function (val, num) {
          v += val.v;
        });
      } else {
        //导出后取显示值
        v = cell.m;
      }
      if (cell.f) {
        value = { formula: cell.f, result: v };
      } else {
        value = v;
      }
      let target = worksheet.getCell(rowid + 1, columnid + 1);
      //添加批注
      if(cell.ps){
        let ps = cell.ps
        target.note = ps.value
      }
      //单元格填充
      target.fill = fill;
      //单元格字体
      target.font = font;
      target.alignment = alignment;
      target.value = value;
      return true;
    });
  });
};
/**
 * *数据验证
 */
// var setDataValidation =  function(verify,worksheet){
//   if(!verify) return
//   for (const key in verify) {
//     const row_col = key.split('_')
//     let cell =worksheet.getCell(Number(row_col[0])+1,Number(row_col[1])+1)
 
//   }
// }
/**
 * *数据透视图
 */
/**
 * *筛选导出
 */
var setFilter = function(filter,worksheet){
  if(!filter) return
  worksheet.autoFilter = createCellRange(filter.row,filter.column)
}
/**
 * *冻结视图
 */
var setFrozen = function(frozen,worksheet){
  //不存在冻结或取消冻结，则不执行后续代码
  if(!frozen||frozen.type=='cancel') return
  //执行冻结操作代码
  let views = []
  switch (frozen.type) {
    //冻结首行
    case 'row':
      views = [
        {state: 'frozen', xSplit: 0, ySplit: 1}
      ];
      break;
    //冻结首列
    case 'column':
      views = [
        {state: 'frozen', xSplit: 1, ySplit: 0}
      ];
      break;
    //冻结首行首列
    case 'both':
      views = [
        {state: 'frozen', xSplit: 1, ySplit: 1}
      ];
      break;
    //冻结行至选区
    case 'rangeRow':
      views = [
        {state: 'frozen', xSplit: 0, ySplit: frozen.range.row_focus+1}
      ];
      break;
    //冻结列至选区
    case  'rangeColumn':
      views = [
        {state: 'frozen', xSplit: frozen.range.column_focus+1, ySplit: 0}
      ];
      break;
    //冻结至选区
    case 'rangeBoth':
      views = [
        {state: 'frozen', xSplit: frozen.range.column_focus+1, ySplit: frozen.range.row_focus+1}
      ];
      break;
  }
  worksheet.views = views
}
/**
 * *设置超链接
 */
var setHyperlink = function(hyperlink,worksheet){
  if (!hyperlink) return;
  for (const key in hyperlink) {
    const row_col = key.split('_')
    let cell =worksheet.getCell(Number(row_col[0])+1,Number(row_col[1])+1)
    let font = cell.style.font
    //设置导出后超链接的样式
    // cell.font= fontConvert(font.name,'#0000ff',font.bold,font.italic,font.size,font.strike,true)
    cell.font= fontConvert(font.name,'#0000ff',0,0,font.size,0,true)
    if(hyperlink[key].linkType=="external"){
      //外部链接
      cell.value = {
        text: cell.value,
        hyperlink: hyperlink[key].linkAddress,
        tooltip: hyperlink[key].linkTooltip
      }
    }else{
      // 内部链接
      const linkArr = hyperlink[key].linkAddress.split('!')
      let hyper = '#\\'+linkArr[0]+'\\'+'!'+linkArr[1]
      cell.value = {
        text: cell.value,
        hyperlink:hyper,
        tooltip: hyperlink[key].linkTooltip
      }
    }
    
  }
}
/**
 * *设置图片
 */
var setImages = function (thesheet, worksheet, workbook) {
  let {
    images,//图片对象或者数组
    visibledatacolumn, //所有行的位置
    visibledatarow, //所有列的位置
  } = { ...thesheet };
  if (typeof images != "object") return;
  for (let key in images) {
    // 通过 base64  将图像添加到工作簿
    const myBase64Image = images[key].src;
    //开始行 开始列 结束行 结束列
    const item = images[key];
    const imageId = workbook.addImage({
      base64: myBase64Image,
      extension: "png",
    });
 
    const col_st = getImagePosition(item.default.left, visibledatacolumn);
    const row_st = getImagePosition(item.default.top, visibledatarow);
 
    //模式1，图片左侧与luckysheet位置一样，像素比例保持不变，但是，右侧位置可能与原图所在单元格不一致
    worksheet.addImage(imageId, {
      tl: { col: col_st, row: row_st },
      ext: { width: item.default.width, height: item.default.height },
    });
    //模式2,图片四个角位置没有变动，但是图片像素比例可能和原图不一样
    // const w_ed = item.default.left+item.default.width;
    // const h_ed = item.default.top+item.default.height;
    // const col_ed = getImagePosition(w_ed,visibledatacolumn);
    // const row_ed = getImagePosition(h_ed,visibledatarow);
    // worksheet.addImage(imageId, {
    //   tl: { col: col_st, row: row_st},
    //   br: { col: col_ed, row: row_ed},
    // });
  }
};
//获取图片在单元格的位置
var getImagePosition = function (num, arr) {
  let index = 0;
  let minIndex;
  let maxIndex;
  for (let i = 0; i < arr.length; i++) {
    if (num < arr[i]) {
      index = i;
      break;
    }
  }
 
  if (index == 0) {
    minIndex = 0;
    maxIndex = 1;
    return Math.abs((num - 0) / (arr[maxIndex] - arr[minIndex])) + index;
  } else if (index == arr.length - 1) {
    minIndex = arr.length - 2;
    maxIndex = arr.length - 1;
  } else {
    minIndex = index - 1;
    maxIndex = index;
  }
  let min = arr[minIndex];
  let max = arr[maxIndex];
  let radio = Math.abs((num - min) / (max - min)) + index;
  return radio;
};
//单元格背景填充色处理
var fillConvert = function (bg) {
  if (!bg) {
    return null;
    // return {
    // 	type: 'pattern',
    // 	pattern: 'solid',
    // 	fgColor:{argb:'#ffffff'.replace('#','')}
    // }
  }
  bg = bg.indexOf("rgb") > -1 ? rgb2hex(bg) : bg;
  let fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: bg.replace("#", "") },
  };
  return fill;
};
//转换颜色
var rgb2hex = function (rgb) {
  if (rgb.charAt(0) == "#") {
    return rgb;
  }
 
  var ds = rgb.split(/\D+/);
  var decimal = Number(ds[1]) * 65536 + Number(ds[2]) * 256 + Number(ds[3]);
  return "#" + zero_fill_hex(decimal, 6);
 
  function zero_fill_hex(num, digits) {
    var s = num.toString(16);
    while (s.length < digits) s = "0" + s;
    return s;
  }
};
//字体转换处理
var fontConvert = function (
  ff = 0,
  fc = "#000000",
  bl = 0,
  it = 0,
  fs = 10,
  cl = 0,
  ul = 0
) {
  // luckysheet：ff(样式), fc(颜色), bl(粗体), it(斜体), fs(大小), cl(删除线), ul(下划线)
  const luckyToExcel = {
    0: "微软雅黑",
    1: "宋体（Song）",
    2: "黑体（ST Heiti）",
    3: "楷体（ST Kaiti）",
    4: "仿宋（ST FangSong）",
    5: "新宋体（ST Song）",
    6: "华文新魏",
    7: "华文行楷",
    8: "华文隶书",
    9: "Arial",
    10: "Times New Roman",
    11: "Tahoma ",
    12: "Verdana",
    num2bl: function (num) {
      return num === 0||false ? false : true;
    },
  };
  // let color = fc ? "" : (fc + "").indexOf("rgb") > -1 ? util.rgb2hex(fc) : fc;
  // let color = fc ? fc : (fc + "").indexOf("rgb") > -1 ? util.rgb2hex(fc) : fc;
 
  let font = {
    name: ff,
    family: 1,
    size: fs,
    color: { argb: fc.replace("#", "") },
    bold: luckyToExcel.num2bl(bl),
    italic: luckyToExcel.num2bl(it),
    underline: luckyToExcel.num2bl(ul),
    strike: luckyToExcel.num2bl(cl),
  };
 
  return font;
};
//对齐转换
var alignmentConvert = function(
  vt = 'default',
  ht = 'default',
  tb = 'default',
  tr = 'default'
) {
  // luckysheet:vt(垂直), ht(水平), tb(换行), tr(旋转)
  const luckyToExcel = {
    vertical: {
      0: 'middle',
      1: 'top',
      2: 'bottom',
      default: 'top'
    },
    horizontal: {
      0: 'center',
      1: 'left',
      2: 'right',
      default: 'left'
    },
    wrapText: {
      0: false,
      1: false,
      2: true,
      default: false
    },
    textRotation: {
      0: 0,
      1: 45,
      2: -45,
      3: 'vertical',
      4: 90,
      5: -90,
      default: 0
    }
  }
 
  let alignment = {
    vertical: luckyToExcel.vertical[vt],
    horizontal: luckyToExcel.horizontal[ht],
    wrapText: luckyToExcel.wrapText[tb],
    textRotation: luckyToExcel.textRotation[tr]
  }
  return alignment
}
//边框转换
var borderConvert = function(borderType, style = 1, color = '#000') {
  // 对应luckysheet的config中borderinfo的的参数
  if (!borderType) {
    return {}
  }
  const luckyToExcel = {
    type: {
      'border-all': 'all',
      'border-top': 'top',
      'border-right': 'right',
      'border-bottom': 'bottom',
      'border-left': 'left'
    },
    style: {
      0: 'none',
      1: 'thin',
      2: 'hair',
      3: 'dotted',
      4: 'dashDot', // 'Dashed',
      5: 'dashDot',
      6: 'dashDotDot',
      7: 'double',
      8: 'medium',
      9: 'mediumDashed',
      10: 'mediumDashDot',
      11: 'mediumDashDotDot',
      12: 'slantDashDot',
      13: 'thick'
    }
  }
  let template = {
    style: luckyToExcel.style[style],
    color: { argb: color.replace('#', '') }
  }
  let border = {}
  if (luckyToExcel.type[borderType] === 'all') {
    border['top'] = template
    border['right'] = template
    border['bottom'] = template
    border['left'] = template
  } else {
    border[luckyToExcel.type[borderType]] = template
  }
  return border
}
//向单元格添加边框
function addborderToCell(borders, row_index, col_index) {
  let border = {}
  const luckyExcel = {
    type: {
      l: 'left',
      r: 'right',
      b: 'bottom',
      t: 'top'
    },
    style: {
      0: 'none',
      1: 'thin',
      2: 'hair',
      3: 'dotted',
      4: 'dashDot', // 'Dashed',
      5: 'dashDot',
      6: 'dashDotDot',
      7: 'double',
      8: 'medium',
      9: 'mediumDashed',
      10: 'mediumDashDot',
      11: 'mediumDashDotDot',
      12: 'slantDashDot',
      13: 'thick'
    }
  }
  for (const bor in borders) {
    if (borders[bor].color.indexOf('rgb') === -1) {
      border[luckyExcel.type[bor]] = {
        style: luckyExcel.style[borders[bor].style],
        color: { argb: borders[bor].color.replace('#', '') }
      }
    } else {
      border[luckyExcel.type[bor]] = {
        style: luckyExcel.style[borders[bor].style],
        color: { argb: borders[bor].color }
      }
    }
  }
 
  return border
}
export default exportExcel