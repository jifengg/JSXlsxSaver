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


# best practice 最佳实践

- 查看[`Demo`](https://github.com/jifengg/JSXlsxDemo)获得更多使用方式；
- 在xlsx中，当你有一个文本需要在多个单元格中显示时，可以使用`Book.prototype.CreateShareString()`获得一个`共享文本`，多处使用时对文件体积的影响微乎其微。可以在Demo中运行`node shareStringDemo.js`查看两个输出文件的大小进行比较；
- 不要修改任何下划线开头的变量，会造成不可预估的错误。