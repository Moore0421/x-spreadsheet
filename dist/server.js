const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.json());

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  next();
});

// 生成文件名方法
function getFileName(id) {
  return id ? `spreadsheet-data-${id}.json` : 'spreadsheet-data.json';
}

// 保存数据接口
app.post('/api/save', (req, res) => {
  const { id } = req.query;
  const data = JSON.stringify(req.body);
  const fileName = getFileName(id);
  
  try {
    fs.writeFileSync(fileName, data);
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 获取数据接口
app.get('/api/getData', (req, res) => {
  const { id } = req.query;
  const fileName = getFileName(id);

  try {
    if (fs.existsSync(fileName)) {
      const data = fs.readFileSync(fileName, 'utf8');
      res.json(JSON.parse(data));
    } else {
      res.json({}); // 文件不存在返回空对象
    }
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.listen(3000, () => {
  console.log('Server running on port 3000');
});