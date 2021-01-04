# 说明

- 原理

  1. 首次开始读取 excel 文件，生成 excel.json 保存在本地
  2. 使用 puppeteer 调起本地浏览器，打开相应地址
  3. 每处理完成一条，更新 json 数据，更新生成的 excel
  4. 自动跳转下一条

- 优点
  - 不需要在 excel 中处理，不需要在 excel 中插入行
  - 处理完成自动打开下一条公告详情
  - 遇到标题和上一条重复的，自动写入备注 **与 xxx 重复**
  - 可以看到当前进度、处理用时
  - 不需要担心多了空格或者回车，保证格式正确
  - 正文识别，能减少输入的字段
  - 处理时间自动添加
- 使用
  1. 将处理的 excel 放入 exe 可执行文件相同目录，
  2. 修改 config.json
  3. 本地已安装 chrome 浏览器
  4. 运行 exe 文件

- 配置说明

```javascript
{
    "index": 109, // 将要处理的那条数据在excel.json中的index
    "lastId": 1113,// 上一行id，不需要配置
    "lastSnapshotId": "18918",//上一条公告的snapshotId，不需要配置
    "sheetName": "Sheet7", // excel文件工作表名
    "fileName": "DCQA-0.XLSX", //excel文件名
    "defaultViewport": { //浏览器大小
        "width": 800,
        "height": 1080
    },
    "origin": "http://ssl.gdcers.com", //监测平台地址
    "userInfo": {
        "handler": "雷任胜", // 表格中的【处理人】
        "username": "ostest", // 账号
        "password": "ostest123"//密码
    }
}
```

- 相关问题
  - 什么时候需要登录？
    `只有第一次使用自动打开登录页面，输入验证码后跳转公告详情`
  - 从哪里开始？
    `在config.json中的index就是代表要处理的json索引，也可以从导出的excel文件中获取到`
  - 发现当前公告之前的输入有误怎么办？
    `ctrl+c结束掉，重新开始`
  - 怎么样修改之前处理过的某一条公告？
    `从导出的excel中获取要修改的公告的index，修改config.json中的index，完成后修改回来`
  - 想跳过某一条，如果处理？
    `问题内容填空，直接在备注写**待定**`
- 注意
  - 表头不能修改
