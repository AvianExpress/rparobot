const fs=require('fs');
const args = process.argv.slice(2);
const pathToEx=args[0];
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const util = require('util');
let start = Date.now();
let start1= new Date();
let logname = ('/'+ start1.getHours()+start1.getMinutes()+start1.getSeconds()+ '.log');
let log_file = fs.createWriteStream(__dirname + logname, {flags : 'w'});
let log_stdout = process.stdout;
let log_filed = fs.createWriteStream(__dirname + '/debug'+ Date.now()+'.log', {flags : 'w'});
console.log = function(d, param1=undefined) {
  if (param1!=undefined){
    start1=new Date();
  param1.write('['+start1.getHours() +':'+start1.getMinutes()+':'+start1.getSeconds()+'] '+util.format(d) + '\n');
  }
  log_stdout.write('['+start1.getHours() +':'+start1.getMinutes()+':'+start1.getSeconds()+'] '+util.format(d) + '\n');
 
};
(async () =>{
console.log('Кушать охота', log_file);
await delay(5000);
console.log('Щас бы сальца');
console.log('Дебаг ежже', log_filed);

})();





