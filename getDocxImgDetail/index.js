
const R = require('ramda')
const fs = require('fs')
const path = require('path')

// 获取文件
const buffer = fs.readFileSync(path.join(__dirname, '入学检测卷_副本.docx'));



const mammoth = require('mammoth')
const htmlparser2 = require('htmlparser2')

mammoth.convertToHtml({ buffer }).then((result) => {
    let html = result.value;
    let message = result.messages;
    // console.log(html,'html')
    let parser = new htmlparser2.parseDOM(result.value);
    console.log(parser)
})






