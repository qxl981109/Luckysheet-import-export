/**
 * *创建单元格所在列的列的字母
 * @param {列数的index值} n 
 * @returns 
 */
export const createCellPos = function(n) {
  let ordA = 'A'.charCodeAt(0)

  let ordZ = 'Z'.charCodeAt(0)
  let len = ordZ - ordA + 1
  let s = ''
  while (n >= 0) {
    s = String.fromCharCode((n % len) + ordA) + s

    n = Math.floor(n / len) - 1
  }
  return s
}
/**
 * *创建单元格范围，期望得到如：A1:D6
 * @param {单元格行数组（例如：[0,3]）} rowArr 
 * @param {单元格列数组（例如：[5,7]）} colArr 
 * */
export const createCellRange = function(rowArr,colArr){
  const startCell = createCellPos(colArr[0])+(rowArr[0]+1)
  const endCell = createCellPos(colArr[1])+(rowArr[1]+1)

  return startCell+':'+endCell
}