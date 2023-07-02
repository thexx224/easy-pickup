const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const _ = require('lodash');
const path = require('path');
const fs = require('fs');

const app = express();

// 设置服务器上的上传目录绝对路径
const uploadDir = '/var/www/uploads';

try {
    // 检查目录是否存在，如果不存在，尝试创建目录
    if (!fs.existsSync(uploadDir)) {
        fs.mkdirSync(uploadDir, {recursive: true, mode: '0777'});
    } else {
        // 如果目录已经存在，尝试设置权限
        fs.chmodSync(uploadDir, '0777');
    }
} catch (error) {
    console.error(`Could not setup upload directory at ${uploadDir}:`, error);
    process.exit(1);
}

// 设置上传目录和文件名
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        // 在每次上传新文件之前，如果已经存在一个文件，就先将其删除
        const existingFiles = fs.readdirSync(uploadDir);
        for (let file of existingFiles) {
            fs.unlinkSync(path.join(uploadDir, file));
        }

        // 使用服务器上的绝对路径
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        cb(null, `${Date.now()}-${file.originalname}`);
    },
});

// 初始化上传中间件
const upload = multer({storage: storage});

// 上传路由，POST到 /upload，文件参数名为 file
app.post('/upload', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    // 读取Excel文件
    const workbook = XLSX.readFile(req.file.path);
    const sheetNameList = workbook.SheetNames;
    const worksheet = workbook.Sheets[sheetNameList[0]];
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 1});

    // 获取所有第一列的数据
    const firstColumnData = data.map(row => row[0]);

    // 从请求中获取用户选择的工作日
    const workdays = req.body.workdays;

    // 对每天进行随机抽取，结果不重复
    const results = new Map();
    for (let day of workdays) {
        const chosen = _.sample(firstColumnData);
        firstColumnData.splice(firstColumnData.indexOf(chosen), 1);
        results.set(day, chosen)
    }

    // 拼接结果字符串
    let resultString = `
    <!DOCTYPE html>
    <html>
    <head>
        <style>
         body { 
              font-family: Arial; 
         }
         ul {
               list-style-type: none;
                margin: 0;
                padding: 0;
         }
         </style>
         <title>结果</title>
        </head>
        <body>
           
         <ul>
         <li>随机选择结果</li><br>
    `;

    for (let [day, chosen] of results.entries()) {
        resultString += `<li>${day}: ${chosen}</li>`;
    }
    resultString += '</ul></body></html>';

    res.send(resultString);
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(5280, () => console.log('Server started at http://localhost:5280'));
