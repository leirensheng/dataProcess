const fs = require('fs')
const utils = require('./utils')
const config = JSON.parse(fs.readFileSync('./config.json'))
let json = JSON.parse(fs.readFileSync('./excel.json'))

let key='表格id（如有多个附件，请插入1行，在原表格id上加.1,如1.1,1.2，不要合并任何单元格）'
let ids = json.map(one=>one[key])
console.log(ids.length);
console.log(new Set(ids).size);


// let i = 0
// let obj = json[i]
// while( obj['公告snapshot_id']!=='19364'){
//   obj['处理人'] = config.userInfo.handler
//   obj['处理时间'] = utils.formatDate(new Date())
//   console.log(i);
//   i++
//   obj = json[i]
// }
// console.log('完成');
// fs.writeFileSync('./excel2.json',JSON.stringify(json,null,4))

