/**
 * *数据验证
 */
export const setDataValidation = function (verify, worksheet) {
  // 不存在不执行
  if (!verify) return;
  // 存在则有以下逻辑
  for (const key in verify) {
    const row_col = key.split("_");
    let cell = worksheet.getCell(
      Number(row_col[0]) + 1,
      Number(row_col[1]) + 1
    );
    let { type, type2, value1, value2 } = verify[key];
    //下拉框--list
    if (type == "dropdown") {
      cell.dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [`${value1}`],
      };
      continue;
    }
    //整数--whole
    if (type == "number_integer") {
      cell.dataValidation = {
        type: "whole",
        operator: setOperator(type2),
        showErrorMessage: true,
        formulae: value2 ? [Number(value1), Number(value2)] : [Number(value1)],
        errorStyle: "error",
        errorTitle: "警告",
        error: errorMsg(type2, type, value1, value2),
      };
      continue;
    }
    //小数-数字--decimal
    if (type == "number_decimal" || "number") {
      cell.dataValidation = {
        type: "decimal",
        operator: setOperator(type2),
        allowBlank: true,
        showInputMessage: true,
        formulae: value2 ? [Number(value1), Number(value2)] : [Number(value1)],
        promptTitle: "警告",
        prompt: errorMsg(type2, type, value1, value2),
      };
      continue;
    }
    //长度受控的文本--textLength
    if (type == "text_length") {
      cell.dataValidation = {
        type: "textLength",
        operator: setOperator(type2),
        showErrorMessage: true,
        allowBlank: true,
        formulae: value2 ? [Number(value1), Number(value2)] : [Number(value1)],
        promptTitle: "错误提示",
        prompt: errorMsg(type2, type, value1, value2),
      };
      continue;
    }
    //文本的内容--text_content
    if (type == "text_content") {
      cell.dataValidation = {};
      continue;
    }
    //日期--date
    if (type == "date") {
      cell.dataValidation = {
        type: "date",
        operator: setDateOperator(type2),
        showErrorMessage: true,
        allowBlank: true,
        promptTitle: "错误提示",
        prompt: errorMsg(type2, type, value1, value2),
        formulae: value2
          ? [new Date(value1), new Date(value2)]
          : [new Date(value1)],
      };
      continue;
    }
    //有效性--custom;type2=="phone"/"card"
    if (type == "validity") {
      
      // cell.dataValidation = {
      //   type: 'custom',
      //   allowBlank: true,
      //   formulae: [type2]
      // };
      continue;
    }
    //多选框--checkbox
    if (type == "checkbox") {
      cell.dataValidation = {};
      continue;
    }
  }
};
//类型type值为"number"/"number_integer"/"number_decimal"/"text_length"时，type2值可为
function setOperator(type2) {
  let transToOperator = {
    bw: "between",
    nb: "notBetween",
    eq: "equal",
    ne: "notEqual",
    gt: "greaterThan",
    lt: "lessThan",
    gte: "greaterThanOrEqual",
    lte: "lessThanOrEqual",
  };
  return transToOperator[type2];
}
//数字错误性提示语
function errorMsg(type2, type, value1 = "", value2 = "") {
  const tip = "你输入的不是";
  const tip1 = "你输入的不是长度";
  let errorTitle = {
    bw: `${
      type == "text_length" ? tip1 : tip
    }介于${value1}和${value2}之间的${numType(type)}`,
    nb: `${
      type == "text_length" ? tip1 : tip
    }不介于${value1}和${value2}之间的${numType(type)}`,
    eq: `${type == "text_length" ? tip1 : tip}等于${value1}的${numType(type)}`,
    ne: `${type == "text_length" ? tip1 : tip}不等于${value1}的${numType(
      type
    )}`,
    gt: `${type == "text_length" ? tip1 : tip}大于${value1}的${numType(type)}`,
    lt: `${type == "text_length" ? tip1 : tip}小于${value1}的${numType(type)}`,
    gte: `${type == "text_length" ? tip1 : tip}大于等于${value1}的${numType(
      type
    )}`,
    lte: `${type == "text_length" ? tip1 : tip}小于等于${value1}的${numType(
      type
    )}`,
    //日期
    bf: `${type == "text_length" ? tip1 : tip}早于${value1}的${numType(type)}`,
    nbf: `${type == "text_length" ? tip1 : tip}不早于${value1}的${numType(type)}`,
    af: `${type == "text_length" ? tip1 : tip}晚于${value1}的${numType(
      type
    )}`,
    naf: `${type == "text_length" ? tip1 : tip}不晚于${value1}的${numType(
      type
    )}`,
  };
 
  return errorTitle[type2];
}
// 数字类型（整数，小数，十进制数）
function numType(type) {
  let num = {
    number_integer: "整数",
    number_decimal: "小数",
    number: "数字",
    text_length: "文本",
    date:'日期'
  };
  return num[type];
}
//类型type值为date时
function setDateOperator(type2) {
  let transToOperator = {
    bw: "between",
    nb: "notBetween",
    eq: "equal",
    ne: "notEqual",
    bf: "greaterThan",
    nbf: "lessThan",
    af: "greaterThanOrEqual",
    naf: "lessThanOrEqual",
  };
  return transToOperator[type2];
}