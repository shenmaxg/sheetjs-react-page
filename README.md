# 通用 Excel 数据导入

## 特性

1. 支持 Excel 中表格位置不固定。
2. 支持表头列合并。
3. 支持数据行合并。
4. 支持 Excel 预览。
5. Excel 数据转换为后端友好的 JSON 格式。

## 如何开始

安装依赖，需要 node 10+ 。

```bash
$ npm install
```

启动服务,

```bash
$ npm run start
```

## 效果展示

Excel 文件

![excel](https://raw.githubusercontent.com/shenmaxg/sheetjs-react-page/master/image/excel.png)

转换成的 antd table 预览图

![table](https://raw.githubusercontent.com/shenmaxg/sheetjs-react-page/master/image/antd-table.png)

生成的 JSON 数据
```
[
  {
    "name": "Tom",
    "age": "29",
    "birthday": "1991-09-01",
    "hobby": [
      "music",
      "reading",
      "piano"
    ]
  },
  {
    "name": "Jerry",
    "age": "27",
    "birthday": "1994-12-11",
    "hobby": [
      "reading",
      "draw",
      "piano"
    ]
  }
]
```

## 原理

使用 sheetjs 库解析 Excel ，对解析后的数据定制化处理，转化为 antd table 需要的参数格式，进行预览。同时转换为后端需要的 JSON 格式。

具体的使用规则在文章 [通用 Excel 数据导入方案](https://zhuanlan.zhihu.com/p/289347583)中定义。

## 相关文章

1. [JavaScript 是如何解析 Excel 文件的？](https://zhuanlan.zhihu.com/p/180074383)
2. [通用 Excel 数据导入方案](https://zhuanlan.zhihu.com/p/289347583)

