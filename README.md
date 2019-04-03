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