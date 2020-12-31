const XLSX = require("xlsx");
const path =require('path')
// let keys=[
//   '表格id（如有多个附件，请插入1行，在原表格id上加.1,如1.1,1.2，不要合并任何单元格）',
//   '数据源名称',
//   '数据源类型【0-未分类，1-政府网站，2-政府文件，3-企业网站，4-企业报告等文档，5-媒体】',
//   '所属网站',
//   '所属网站URL【爬取进度的数据源URL】',
//   '网址【公告URL】',
//   '采集时间',
//   '采集人',
//   '公告id',
//   '公告snapshot_id',
//   '发布时间',
//   '发布主体',
//   '公告标题',
//   '处理时间',
//   '处理人',
//   '信息完整性【完整、不完整】',
//   '数据说明',
//   '源文件类型【1-EXCEL，2-PDF，3-WORD，4-HTML，5-jpg】',
//   '获取方式【1-爬取，2-导入，3-录入】',
//   '数据置信度【1-政府公告/文件，2-企业公开报告，3-企业密报，4-自采，5-可信数据源，6-参考数据源】',
//   '公布表格表名',
//   '公布表格合格条目数【抽检表格条目数】',
//   '公布表格不合格条目数【抽检表格条目数】',
//   '公布总抽检批次数',
//   '公布合格率',
//   '处理的表格是否齐全（1-是，2-否）',
//   '备注（缺少表格的，需要注明表格）',
//   '公告附件数量（仅填EXCEL、PDF、WORD附件的数量，无附件不需要填）',
//   '是否存在合并单元格(1-是，2-否)'
// ]


function outputNewSheet(res,config) {
  const wb = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(res);
  XLSX.utils.book_append_sheet(wb, sheet, config.sheetName);
  XLSX.writeFile(wb, path.resolve(__dirname, "./output.xlsx"));
}

module.exports =outputNewSheet