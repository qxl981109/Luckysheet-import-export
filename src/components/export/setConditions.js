/**
 * *条件格式设置
*/
import { createCellRange } from './utils'
export const setConditions = function(conditions,worksheet){
    //条件格式不存在，则不执行后续代码
    if(conditions==undefined) return
 
    //循环遍历规则列表
    conditions.forEach(item => {
        let ruleObj = {
            ref:createCellRange(item.cellrange[0].row,item.cellrange[0].column),
            rules:[]
        }
        //lucksheet对应的为----突出显示单元格规则和项目选区规则
        if(item.type=='default'){
            //excel中type为cellIs的条件下
            if(item.conditionName=='equal'||'greaterThan'||'lessThan'||'betweenness'){
                    ruleObj.rules = setDefaultRules({
                        type:'cellIs',
                        operator:item.conditionName=='betweenness'?'between':item.conditionName,
                        condvalue:item.conditionValue,
                        colorArr:[item.format.cellColor,item.format.textColor]
                    })
                    worksheet.addConditionalFormatting(ruleObj)
                }
            //excel中type为containsText的条件下
            if(item.conditionName=='textContains'){
                ruleObj.rules = [
                        {
                            type:'containsText',
                            operator:'containsText', //表示如果单元格值包含在text 字段中指定的值，则应用格式
                            text:item.conditionValue[0],
                            style: setStyle([item.format.cellColor,item.format.textColor])
                        }
                    ]
                worksheet.addConditionalFormatting(ruleObj)
            }
            //发生日期--时间段
            if(item.conditionName=='occurrenceDate'){
                ruleObj.rules = [
                        {
                            type:'timePeriod',
                            timePeriod:'today', //表示如果单元格值包含在text 字段中指定的值，则应用格式
                            style: setStyle([item.format.cellColor,item.format.textColor])
                        }
                    ]
                worksheet.addConditionalFormatting(ruleObj)
            }
            //重复值--唯一值
            // if(item.conditionName=='duplicateValue'){
            //     ruleObj.rules = [
            //             {
            //                 type:'expression',
            //                 formulae:'today', //表示如果单元格值包含在text 字段中指定的值，则应用格式
            //                 style: setStyle([item.format.cellColor,item.format.textColor])
            //             }
            //         ]
            //     worksheet.addConditionalFormatting(ruleObj)
            // }
            //项目选区规则--top10前多少项的操作
            if(item.conditionName=='top10'||'top10%'||'last10'||'last10%'){
                ruleObj.rules = [
                        {
                            type:'top10',
                            rank:item.conditionValue[0], //指定格式中包含多少个顶部（或底部）值
                            percent:item.conditionName=='top10'||'last10'?false:true,
                            bottom:item.conditionName=='top10'||'top10%'?false:true,
                            style: setStyle([item.format.cellColor,item.format.textColor])
                        }
                    ]
                worksheet.addConditionalFormatting(ruleObj)
            }
            //项目选区规则--高于/低于平均值的操作
            if(item.conditionName=='AboveAverage'||'SubAverage'){
                ruleObj.rules = [
                        {
                            type:'aboveAverage',
                            aboveAverage:item.conditionName=='AboveAverage'?true:false,
                            style: setStyle([item.format.cellColor,item.format.textColor])
                        }
                    ]
                worksheet.addConditionalFormatting(ruleObj)
            }
            return
        }
            
        //数据条
        if(item.type == 'dataBar'){
            ruleObj.rules = [
                {
                    type:'dataBar',
                    style:{}
                }
            ]
            worksheet.addConditionalFormatting(ruleObj)
            return
        }
        //色阶
        if(item.type == 'colorGradation'){
            ruleObj.rules = [
                {
                    type:'colorScale',
                    color:item.format,
                    style:{}
                }
            ]
            worksheet.addConditionalFormatting(ruleObj)
            return
        }
        //图标集
        if(item.type == 'icons'){
            ruleObj.rules = [
                    {
                        type:'iconSet',
                        iconSet:item.format.len
                    }
                ]
            worksheet.addConditionalFormatting(ruleObj)
            return
        }
    });
  }
/**
 * 
 * @param {
 *  type:lucketsheet对应的条件导出类型；
 *  operator：excel对应的条件导入类型；
 *  condvalue：1个公式字符串数组，返回要与每个单元格进行比较的值；
 *  colorArr：颜色数组，第一项为单元格填充色，第二项为单元格文本颜色
 * } obj 
 * @returns 
 */
function setDefaultRules(obj){
    let rules = [
        {
            type:obj.type,
            operator:obj.operator,
            formulae:obj.condvalue,
            style:setStyle(obj.colorArr)
        }
    ]
    return rules
}
/**
 * 
 * @param {颜色数组，第一项为单元格填充色，第二项为单元格文本颜色} colorArr 
 */
function setStyle(colorArr){
    return {
        fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: colorArr[0].replace("#", "")}},
        font: {color:{ argb: colorArr[1].replace("#", "")},}
    }
}