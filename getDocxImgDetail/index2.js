const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const fs = require('fs')
const path = require('path')
const sizeOf = require('image-size');
const mammoth = require("mammoth");

const filePath = path.join(__dirname, '入学检测卷_副本.docx');
// Load the docx file
const buffer = fs.readFileSync(filePath);
const zip = new PizZip(buffer);

const doc = new Docxtemplater().loadZip(zip);

console.log('doc===>', doc)

// Get all image elements from the document
const images = doc.getZip().file(/word\/media\/image\d+\.[a-zA-Z]+/);

// console.log(images, 'images')

let imagesSize = []

images.forEach(image => {

    // console.log(image, 'image')
    const data = image._data.getCompressedContent();


    // console.log(data, 'data')
    const buffer = Buffer.from(data, 'binary'); // convert to buffer

    // console.log(buffer, 'buffer')
    const dimensions = sizeOf(buffer);

    console.log(dimensions, 'dimensions')

    console.log(`Image ${image.name}: width=${dimensions.width}, height=${dimensions.height}`);

    imagesSize.push(dimensions)

    // const data = image._data.compressedContent;
    // const img = new Image();
    // img.src = 'data:image/png;base64,' + btoa(data);

    // console.log(`Image ${image.name}: width=${img.width}, height=${img.height}`);
});

// // Get all images in the document
// const images = doc.getImageData();

// // Loop through each image and log its dimensions
// images.forEach(image => {
//     console.log(`Image size: ${image.width} x ${image.height}`);
// });


const getHtmlTemplate = (content) => {
    const html = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Document</title>
    </head>
    <body>
        ${content}
    </body>
    </html>
    `;

    return html;
};

// ========
// 转为 html 
// ========
; (async () => {
    const result = await mammoth.convertToHtml({ path: filePath });
    let html = result.value;

    console.log(imagesSize,'imagesSize')
    const imgTags = html.match(/<img[^>]+>/g);

    imgTags.forEach((imgTag, i) => {
        const imgDimensions = imagesSize[i];
        if(imgDimensions){
            console.log(imgDimensions, 'imgDimensions')

            const width = imgDimensions.width;
            const height = imgDimensions.height;
    
            // 将width和height属性添加到img标签中
            const newImgTag = imgTag.replace(/(alt=".*?")?>/gi, `$1 width="${width}" height="${height}">`);
    
            // 将原始img标签替换为新的带有width和height属性的img标签
            html = html.replace(imgTag, newImgTag);
        }
    });



    fs.writeFile(path.join(__dirname, 'a.html'), getHtmlTemplate(html), (err) => {
        if (err) {
            console.error(err);
            return;
        }
        // console.log("内容已成功写入文件");
    });
})();






