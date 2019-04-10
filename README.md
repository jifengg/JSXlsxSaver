# JSXlsxSaver
A xlsx file saver base on JSXlsxCore for node or browser.

基于`JSXlsxCore`([git](https://github.com/jifengg/JSXlsxCore)/[npm](https://www.npmjs.com/package/js-xlsx-core))实现的生成`微软Office Excel xlsx文件格式`的库。可以在node或者现代浏览器中使用。

本代码使用了`jszip`([npm](https://www.npmjs.com/package/jszip))用于生成zip文件（xlsx本质上是一个zip文件）;

如果要在浏览器在使用，推荐使用`file-saver`（[npm](https://www.npmjs.com/package/file-saver)）来保存文件到本地。jszip中已经包含了这个文件（./vendor/FileSaver.js）

# use in node

shell:
```shell
npm i xlsx-saver
```

js:

```javascript
const XlsxCore = require('js-xlsx-core');
require('xlsx-saver');
const {
    Book,
    HorizontalAlignment,
    VerticalAlignment
} = XlsxCore;
var book = new Book();
var sheet = book.CreateSheet("第一页");
sheet.AddText('一个普通文本', 0, 0);
//...
//get buffer,you can send to http respone,or save to localfile
var buffer = await book.SaveAsBuffer();
//save to localfile
fs.writeFileSync('./out.xlsx', buffer);
```


# use in browser

shell:
```shell
npm i xlsx-saver file-saver
```

html:

```html
<script src="node_modules/jszip/dist/jszip.js"></script>
<script src="node_modules/jszip/vendor/FileSaver.js"></script>
<script src="node_modules/js-xlsx-core/xlsxcore.js"></script>
<script src="node_modules/xlsx-saver/xlsxsaver.js"></script>
```

js:
```javascript
var test = async () => {
    const {
        Book,
        Sheet,
        Cell,
        ShareString,
        CellStyle,
        CellAlignment,
        NumberFormat,
        Image,
        ImageOption,
        HorizontalAlignment,
        VerticalAlignment
    } = XlsxCore;
    var book = new Book();
    var sheet = book.CreateSheet("第一页");
    sheet.AddText('一个普通文本', 0, 0);
    //...
    var bolb = await book.SaveAsBolb();
    // see FileSaver.js
    saveAs(bolb, "test.xlsx");
};

test();
```

# support 支持的操作

- 可在单元格中存放字符串、数字和图片。（`时间类型暂时不支持，请使用字符串或数字存储。`）
- 设置单元格字体样式：字体名称、字号、文字颜色、是否粗体、是否斜体、是否下划线；
- 设置单元格纯颜色填充；
- 设置单元格超链接；
- 设置单元格水平垂直对齐方式；
- 设置单元格是否支持换行；
- 合并单元格；
- 设置默认字体样式、默认行高、默认列宽；
- 设置指定行的高度（行高所使用单位为磅，1厘米=28.6磅）；
- 设置指定列的宽度（列宽使用单位为1/10英寸，既1个单位为2.54毫米）；
- 设置单元格内图片大小（图片大小不能超过单元格，所以如果要设置比较大的图片，请设置合适的单元格大小）；
- 设置单元格中数字的显示格式，如`百分比`,`千分符`等等，具体的格式码可以参照[微软文档](https://support.office.com/zh-cn/article/%E6%9F%A5%E7%9C%8B%E6%9C%89%E5%85%B3%E8%87%AA%E5%AE%9A%E4%B9%89%E6%95%B0%E5%AD%97%E6%A0%BC%E5%BC%8F%E7%9A%84%E5%87%86%E5%88%99-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)。
- 通过`共享文本`、`共享字体`、`共享填充`、`共享样式`、`共享格式码`、`共享图片`来达到一处定义多处使用，可以大大减少最终的文件体积。

# best practice 最佳实践

- 查看[`Demo`](https://github.com/jifengg/JSXlsxDemo)获得更多使用方式；
- 在xlsx中，当你有一个文本需要在多个单元格中显示时，可以使用`Book.prototype.CreateShareString()`获得一个`共享文本`，多处使用时对文件体积的影响微乎其微。可以在Demo中运行`node shareStringDemo.js`查看两个输出文件的大小进行比较；
- 不要修改任何下划线开头的变量，会造成不可预估的错误。