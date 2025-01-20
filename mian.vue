<template>
  <div class="container">
    <div class="toolbar">
      <input type="file" @change="loadExcelFile"
        accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" />
      <button @click="exportJson">导出JSON</button>
      <button @click="exportExcel">导出xlsx</button>
    </div>
    <!--web spreadsheet组件-->
    <div ref="sheetContainer" class="grid" id="x-spreadsheet-demo"></div>
  </div>
</template>

<script>
// 在spreadsheet.js中初始化json数据 https://blog.csdn.net/CBGCampus/article/details/125366246
// 带样式导出 https://blog.csdn.net/weixin_42302145/article/details/121476579
// 引入依赖包
import Spreadsheet from "x-data-spreadsheet";
import zhCN from 'x-data-spreadsheet/src/locale/zh-cn'
import _ from "lodash";
import * as XLSX from "xlsx";
import * as Excel from 'exceljs/dist/exceljs'
import * as tinycolor from "tinycolor2";
export default {
  name: "xspreadsheet-demo",
  data() {
    return {
      xs: null,
      jsondata: {
        type: "",
        label: "",
      },
    };
  },
  mounted() {
    this.init();
  },
  methods: {
    init() {
      //设置中文
      Spreadsheet.locale("zh-cn", zhCN);
      this.xs = new Spreadsheet("#x-spreadsheet-demo", {
        mode: "edit",
        showToolbar: true,
        showGrid: true,
        showContextmenu: true,
        showBottomBar: true,
        view: {
          height: () => (_.isNumber(this.$refs.sheetContainer.offsetHeight) && this.$refs.sheetContainer) ? this.$refs.sheetContainer.offsetHeight : 0,
          width: () => (_.isNumber(this.$refs.sheetContainer.offsetWidth) && this.$refs.sheetContainer) ? this.$refs.sheetContainer.offsetWidth : 0,
        },
        formats: [],
        fonts: [],
        formula: [],
        row: {
          len: 9999,
          height: 25,
        },
        col: {
          len: 26,
          width: 100,
          indexWidth: 60,
          minWidth: 60,
        },
        style: {
          bgcolor: "#ffffff",
          align: "left",
          valign: "middle",
          textwrap: false,
          textDecoration: "normal",
          strikethrough: false,
          color: "#0a0a0a",
          align: "center",
          valign: "middle",
          font: {
            name: "Helvetica",
            size: 10,
            bold: false,
            italic: false,
          }
        },
      })
        .loadData([
          {
            name: "调配单变量",
            styles: [
              {
                /*表头样式*/
                bgcolor: '#f4f5f8',
                textwrap: true,
                color: '#900b09',
                border: {
                  top: ['thin', '#1E1E1E'],
                  bottom: ['thin', '#1E1E1E'],
                  right: ['thin', '#1E1E1E'],
                  left: ['thin', '#1E1E1E'],
                },
                font: {
                  bold: true
                }
              }
            ],
            rows: [
              {
                cells: [{
                  text: '序号',
                  editable: false,
                  style: 0
                }, {
                  text: '订单号',
                  editable: false,
                  style: 0
                }]
              }
            ],
            merges: []
          }
        ])
        .change((cdata) => {
        });

      this.xs
        .on("cell-selected", (cell, ri, ci) => {
          console.log("cell:", cell, ", ri:", ri, ", ci:", ci);
        })
        .on("cell-edited", (text, ri, ci) => {
          console.log("text:", text, ", ri: ", ri, ", ci:", ci);
        });
    },
    transRgba(rgba) {
      let result = '';
      let reg = /\(.*\)/;  //字符串匹配括号内的子串
      let arr = []
      result = reg.exec(rgba)[0];  // 截取'(255,255,255,0.6)'
      result = result.substr(1, result.length - 2); //截取十进制的'255,255,255,0.6'
      arr = result.split(','); //字符串切割 ['255','255','255','0.6']
      for (let i = 0; i < arr.length; i++) {
        arr[i] = parseFloat(arr[i]); //字符串类型转为浮点数类型
        if (i == arr.length - 1) {  //对于最后一个透明度数据，需要先将0-1的数值*255
          arr[i] = arr[i] * 255;
        }
        arr[i] = trans10to16(arr[i]);
      }
      return arr.join('');
    },
    // 导入excel
    loadExcelFile(e) {
      const wb = new Excel.Workbook();
      const reader = new FileReader()
      reader.readAsArrayBuffer(e.target.files[0])
      reader.onload = () => {
        const buffer = reader.result;
        // 微软的 Excel ColorIndex 一个索引数字对应一个颜色
        const indexedColors = [
          '000000',
          'FFFFFF',
          'FF0000',
          '00FF00',
          '0000FF',
          'FFFF00',
          'FF00FF',
          '00FFFF',
          '000000',
          'FFFFFF',
          'FF0000',
          '00FF00',
          '0000FF',
          'FFFF00',
          'FF00FF',
          '00FFFF',
          '800000',
          '008000',
          '000080',
          '808000',
          '800080',
          '008080',
          'C0C0C0',
          '808080',
          '9999FF',
          '993366',
          'FFFFCC',
          'CCFFFF',
          '660066',
          'FF8080',
          '0066CC',
          'CCCCFF',
          '000080',
          'FF00FF',
          'FFFF00',
          '00FFFF',
          '800080',
          '800000',
          '008080',
          '0000FF',
          '00CCFF',
          'CCFFFF',
          'CCFFCC',
          'FFFF99',
          '99CCFF',
          'FF99CC',
          'CC99FF',
          'FFCC99',
          '3366FF',
          '33CCCC',
          '99CC00',
          'FFCC00',
          'FF9900',
          'FF6600',
          '666699',
          '969696',
          '003366',
          '339966',
          '003300',
          '333300',
          '993300',
          '993366',
          '333399',
          '333333',
        ];
        wb.xlsx.load(buffer).then(workbook => {
          let workbookData = []
          console.log(workbook)
          workbook.eachSheet((sheet, sheetIndex) => {
            // 构造x-data-spreadsheet 的 sheet 数据源结构
            let sheetData = { name: sheet.name, styles: [], rows: {}, merges: [] }
            // 收集合并单元格信息
            let mergeAddressData = []
            for (let mergeRange in sheet._merges) {
              sheetData.merges.push(sheet._merges[mergeRange].shortRange)
              let mergeAddress = {}
              // 合并单元格起始地址
              mergeAddress.startAddress = sheet._merges[mergeRange].tl
              // 合并单元格终止地址
              mergeAddress.endAddress = sheet._merges[mergeRange].br
              // Y轴方向跨度
              mergeAddress.YRange = sheet._merges[mergeRange].model.bottom - sheet._merges[mergeRange].model.top
              // X轴方向跨度
              mergeAddress.XRange = sheet._merges[mergeRange].model.right - sheet._merges[mergeRange].model.left
              mergeAddressData.push(mergeAddress)
            }
            sheetData.cols = {}
            for (let i = 0; i < sheet.columns.length; i++) {
              sheetData.cols[i.toString()] = {}
              if (sheet.columns[i].width) {
                // 不知道为什么从 exceljs 读取的宽度显示到 x-data-spreadsheet 特别小, 这里乘以8
                sheetData.cols[i.toString()].width = sheet.columns[i].width * 8
              } else {
                // 默认列宽
                sheetData.cols[i.toString()].width = 100
              }
            }

            // 遍历行
            sheet.eachRow((row, rowIndex) => {
              sheetData.rows[(rowIndex - 1).toString()] = { cells: {} }
              //includeEmpty = false 不包含空白单元格
              row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
                let cellText = ''
                if (cell.value && cell.value.result) {
                  // Excel 单元格有公式
                  cellText = cell.value.result
                } else if (cell.value && cell.value.richText) {
                  // Excel 单元格是多行文本
                  for (let text in cell.value.richText) {
                    // 多行文本做累加
                    cellText += cell.value.richText[text].text
                  }
                }
                else {
                  // Excel 单元格无公式
                  cellText = cell.value
                }

                //解析单元格,包含样式
                //*********************单元格存在背景色******************************
                // 单元格存在背景色
                let backGroundColor = null
                if (cell.style.fill && cell.style.fill.fgColor && cell.style.fill.fgColor.argb) {
                  // 8位字符颜色先转rgb再转16进制颜色
                  backGroundColor = ((val) => {
                    val = val.trim().toLowerCase();  //去掉前后空格
                    let color = {};
                    try {
                      let argb = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(val);
                      color.r = parseInt(argb[2], 16);
                      color.g = parseInt(argb[3], 16);
                      color.b = parseInt(argb[4], 16);
                      color.a = parseInt(argb[1], 16) / 255;
                      return tinycolor(`rgba(${color.r}, ${color.g}, ${color.b}, ${color.a})`).toHexString()
                    } catch (e) {
                      console.log(e)
                    }
                  })(cell.style.fill.fgColor.argb)
                }

                if (backGroundColor) {
                  cell.style.bgcolor = backGroundColor
                }
                //*************************************************************************** */

                //*********************字体存在背景色******************************
                // 字体颜色
                let fontColor = null
                if (cell.style.font && cell.style.font.color && cell.style.font.color.argb) {
                  // 8位字符颜色先转rgb再转16进制颜色
                  fontColor = ((val) => {
                    val = val.trim().toLowerCase();  //去掉前后空格
                    let color = {};
                    try {
                      let argb = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(val)
                      color.r = parseInt(argb[2], 16);
                      color.g = parseInt(argb[3], 16);
                      color.b = parseInt(argb[4], 16);
                      color.a = parseInt(argb[1], 16) / 255;
                      return tinycolor(`rgba(${color.r}, ${color.g}, ${color.b}, ${color.a})`).toHexString()
                    } catch (e) {
                      console.log(e)
                    }
                  })(cell.style.font.color.argb)
                }
                if (fontColor) {
                  //console.log(fontColor)
                  cell.style.color = fontColor
                }
                //************************************************************************ */

                // exceljs 对齐的格式转成 x-date-spreedsheet 能识别的对齐格式 
                if (cell.style.alignment && cell.style.alignment.horizontal) {
                  cell.style.align = cell.style.alignment.horizontal
                  cell.style.valign = cell.style.alignment.vertical
                }

                //处理合并单元格
                let mergeAddress = _.find(mergeAddressData, function (o) { return o.startAddress == cell._address })
                if (mergeAddress) {
                  // 遍历的单元格属于合并单元格
                  if (cell.master.address != mergeAddress.startAddress) {
                    // 不是合并单元格中的第一个单元格不需要计入数据源
                    return
                  }
                  // 说明是合并单元格区域的起始单元格
                  sheetData.rows[(rowIndex - 1).toString()].cells[(colNumber - 1).toString()] = { text: cellText, style: 0, merge: [mergeAddress.YRange, mergeAddress.XRange] }
                  sheetData.styles.push(cell.style)
                  //对应的style存放序号
                  sheetData.rows[(rowIndex - 1).toString()].cells[(colNumber - 1).toString()].style = sheetData.styles.length - 1
                }
                else {
                  // 非合并单元格
                  sheetData.rows[(rowIndex - 1).toString()].cells[(colNumber - 1).toString()] = { text: cellText, style: 0 }
                  //解析单元格,包含样式
                  sheetData.styles.push(cell.style)
                  //对应的style存放序号
                  sheetData.rows[(rowIndex - 1).toString()].cells[(colNumber - 1).toString()].style = sheetData.styles.length - 1
                }
              });
            })
            workbookData.push(sheetData)
          })
          this.xs.loadData(workbookData);
        })
      }
    },
    // 导出excel
    exportExcel() {
      const exceljsWorkbook = new Excel.Workbook();
      exceljsWorkbook.modified = new Date();
      this.xs.getData().forEach(function (xws) {
        let rowobj = xws.rows;
        // 构造exceljs文档结构
        const exceljsSheet = exceljsWorkbook.addWorksheet(xws.name)
        // 读取列宽
        let sheetColumns = []
        let colIndex = 0
        for (let col in xws.cols) {
          if (xws.cols[col].width) {
            sheetColumns.push({ header: colIndex + '', key: colIndex + '', width: xws.cols[col].width / 8 })
          }
          colIndex++
        }
        exceljsSheet.columns = sheetColumns
        for (let ri = 0; ri < rowobj.len; ++ri) {
          let row = rowobj[ri];
          if (!row) continue;
          // 构造exceljs的行(如果尚不存在，则将返回一个新的空对象)
          const exceljsRow = exceljsSheet.getRow(ri + 1)
          Object.keys(row.cells).forEach(function (k) {
            let idx = +k;
            if (isNaN(idx)) return;
            const exceljsCell = exceljsRow.getCell(Number(k) + 1)
            exceljsCell.value = row.cells[k].text
            //console.log(row.cells[k])
            //console.log(xws.styles[row.cells[k].style])
            // 水平对齐方式
            if (xws.styles[row.cells[k].style] && xws.styles[row.cells[k].style].alignment) {
              if (xws.styles[row.cells[k].style].alignment.vertical) {
                if (exceljsCell.alignment == undefined) {
                  exceljsCell.alignment = {}
                }
                exceljsCell.alignment.vertical = xws.styles[row.cells[k].style].alignment.vertical
              }
            }
            // 垂直对齐方式
            if (xws.styles[row.cells[k].style] && xws.styles[row.cells[k].style].valign) {
              if (!exceljsCell.alignment == undefined) {
                exceljsCell.alignment = {}
              }
              exceljsCell.alignment.horizontal = xws.styles[row.cells[k].style].valign
            }
            // 边框
            exceljsCell.border = xws.styles[row.cells[k].style].border
            // 背景色
            if (xws.styles[row.cells[k].style].bgcolor) {
              let rgb = tinycolor(xws.styles[row.cells[k].style].bgcolor).toRgb()
              let rHex = parseInt(rgb.r).toString(16).padStart(2, '0')
              let gHex = parseInt(rgb.g).toString(16).padStart(2, '0')
              let bHex = parseInt(rgb.b).toString(16).padStart(2, '0')
              let aHex = parseInt(rgb.a).toString(16).padStart(2, '0')
              let _bgColor = aHex + rHex + gHex + bHex
              // 设置exceljs背景色
              exceljsCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: _bgColor }
              }
            }
            // 字体
            exceljsCell.font = xws.styles[row.cells[k].style].font
            // 字体颜色
            if (xws.styles[row.cells[k].style].color) {
              let rgb = tinycolor(xws.styles[row.cells[k].style].color).toRgb()
              let rHex = parseInt(rgb.r).toString(16).padStart(2, '0')
              let gHex = parseInt(rgb.g).toString(16).padStart(2, '0')
              let bHex = parseInt(rgb.b).toString(16).padStart(2, '0')
              let aHex = parseInt(rgb.a).toString(16).padStart(2, '0')
              let _fontColor = aHex + rHex + gHex + bHex
              exceljsCell.font.color = { argb: _fontColor }
            }
            // 合并单元格
            if (row.cells[k].merge) {
              // 开始行
              let startRow = ri + 1
              // 结束行,加上Y轴跨度
              let endRow = startRow + row.cells[k].merge[0]
              // 开始列
              let startColumn = Number(k) + 1
              // 结束列,加上X轴跨度
              let endColumn = startColumn + row.cells[k].merge[1]
              // 按开始行，开始列，结束行，结束列合并
              exceljsSheet.mergeCells(startRow, startColumn, endRow, endColumn);
            }
          });
        }
      });
      // writeBuffer 把写好的excel 转换成 ArrayBuffer 类型
      exceljsWorkbook.xlsx.writeBuffer().then((data) => {
        const link = document.createElement('a');
        // Blob 实现下载excel
        const blob = new Blob([data], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
        });
        link.href = window.URL.createObjectURL(blob);
        link.download = '调配单变量配置.xlsx';
        link.click();
      });
    },
    // 导出为 JSON
    exportJson() {
      let sheetsData = this.xs.getData();
      let rows = Object.entries(sheetsData[0].rows);
      let objectProperties = [
        "Index",
        "OrderIndex",
        "OrderNo",
        "ProductName",
        "OrderStatus",
      ];
      let jsonData = [];
      // 遍历数据，跳过第一行表头
      for (let i = 1; i < rows.length; i++) {
        if (rows[i] && rows[i][1] && rows[i][1].cells) {
          let row = Object.entries(rows[i][1].cells);
          // 构造行对象
          let JsonRow = {
            Index: null,
            OrderIndex: null,
            OrderNo: null,
            ProductName: null,
            OrderStatus: null,
          };
          for (let k = 0; k < row.length; k++) {
            let cells = row[k];
            JsonRow[objectProperties[k]] = cells[1].text;
          }
          jsonData.push(JsonRow);
        }
      }
      console.log(jsonData);
    },
    stox(wb) {
      var out = [];
      wb.SheetNames.forEach(function (name) {
        var o = { name: name, rows: {} };
        var ws = wb.Sheets[name];
        var aoa = XLSX.utils.sheet_to_json(ws, { raw: false, header: 1 });
        aoa.forEach(function (r, i) {
          var cells = {};
          r.forEach(function (c, j) {
            cells[j] = { text: c };
          });
          o.rows[i] = { cells: cells };
        });
        out.push(o);
      });
      return out;
    },
    fixData(data) {
      var o = "",
        l = 0,
        w = 10240;
      for (; l < data.byteLength / w; ++l)
        o += String.fromCharCode.apply(
          null,
          new Uint8Array(data.slice(l * w, l * w + w))
        );
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
      return o;
    },
  },
};
</script>
<style lang="less" scoped>
.container {
  width: 100%;
  height: 100%;
  display: flex;
  flex-direction: column;

  .toolbar {
    width: 100%;
    height: 50px;
  }

  .grid {
    width: 100%;
    height: calc(100% - 80px);
  }

  /deep/ .x-spreadsheet-toolbar {
    padding: 0px;
    width: calc(100% - 2px) !important;
  }
}
</style>