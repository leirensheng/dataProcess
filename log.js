const chalk = require('chalk');

const log = console.log



log.green =(val)=>{log(chalk.green(typeof val==='object'?JSON.stringify(val,null,4):val))} 
log.red =(val)=>{log(chalk.red(typeof val==='object'?JSON.stringify(val,null,4):val))} 
log.yellow =(val)=>{log(chalk.yellow(typeof val==='object'?JSON.stringify(val,null,4):val))} 
log.pink =(val)=>{log(chalk.magentaBright(typeof val==='object'?JSON.stringify(val,null,4):val))} 

module.exports = log