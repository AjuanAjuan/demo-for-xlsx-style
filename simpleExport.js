import XLSX from 'xlsx-style'

function dealDataSource (headerName = '未命名', dataName = [], data = []) {
  let header = []
  dataName.forEach((item) => {
    if (!item.title || !item.props) return
    header.push({title: item.title, dataIndex: item.props, width: 120})
  })
  let dataSource = [...data]
  let workSheetConfig = {
    merges: []
  }
  let cellConfig = {
    headerStyle: { // 表头区域样式配置
    border: {},
    font: {},
    alignment: {},
    fill: {}
    },
    dataStyle: { // 内容样式配置
    border: {},
    font: {},
    alignment: {},
    fill: {}
    } 
  }
  let sheetName = headerName // sheet名
  return {header, dataSource, workSheetConfig, cellConfig, sheetName}
}

const defaultWorkBook = {
  bookType: 'xlsx',
  bookSST: false,
  type: 'binary'
}
// 默认样式配置
const borderAll = {
  top: {
    style: 'thin'
  },
  bottom: {
    style: 'thin'
  },
  left: {
    style: 'thin'
  },
  right: {
    style: 'thin'
  }
}

const defaultCellStyle = {
  // 表头区域样式配置
  headerStyle: {
    border: borderAll,
    font: { name: '宋体', sz: 16, italic: false, underline: false, bold: true },
    alignment: { vertical: 'center', horizontal: 'center' },
    fill: { fgColor: { rgb: 'FFFFFF' } }
  },
  // 内容区域样式配置
  dataStyle: {
    border: borderAll,
    font: { name: '宋体', sz: 11, italic: false, underline: false },
    alignment: { vertical: 'center', horizontal: 'center', wrapText: true },
    fill: { fgColor: { rgb: 'FFFFFF' } }
  }
}

// 表头样式
const headerStyle_SheetTwo = {
  border: borderAll,
  font: { name: '宋体', sz: 12, italic: false, underline: false, bold: true },
  alignment: { vertical: 'center', horizontal: 'center' },
  fill: { fgColor: { rgb: 'D0CECE' } }
}

function exportAllData (fileName = '测试用例', titleProps, exportData, workBookConfig = defaultWorkBook) {
  let excelObj = dealDataSource(fileName, titleProps, exportData)
  console.log(excelObj);
  // 定义工作簿对象
  const wb = { SheetNames: [], Sheets: {} }

  // 处理sheet表头
  const _header = excelObj.header.map((item, i) =>
    Object.assign({}, {
      key: item.dataIndex,
      title: item.title,
      // 定位单元格
      position: getCharCol(i) + 1,
      // 设置表头样式
      // s: data.cellConfig && data.cellConfig.headerStyle ? data.cellConfig.headerStyle : defaultCellStyle.headerStyle,
      s: headerStyle_SheetTwo
    })
  ).reduce((prev, next) =>
    Object.assign({}, prev, {
      [next.position]: { v: next.title, key: next.key, s: next.s },
    }), {}
  )
  // 处理sheet内容
  const _data = {}
  excelObj.dataSource.forEach((item, i) => {
    excelObj.header.forEach((obj, index) => {
      const key = getCharCol(index) + (i + 2)
      const key_t = obj.dataIndex
      _data[key] = {
        v: item[key_t],
        // s: data.cellConfig && data.cellConfig.dataStyle ? data.cellConfig.dataStyle : defaultCellStyle.dataStyle,
        s: defaultCellStyle.dataStyle
      }
    })
  })

  const output = Object.assign({}, _header, _data)
  const outputPos = Object.keys(output)
  // 设置单元格宽度
  const colWidth = excelObj.header.map(item => { return { wpx: item.width || 80 } })

  const merges = excelObj.workSheetConfig && excelObj.workSheetConfig.merges

  const freeze = excelObj.workSheetConfig && excelObj.workSheetConfig.freeze

  // 处理sheet名
  wb.SheetNames[0] = excelObj.sheetName ? excelObj.sheetName : 'Sheet' + 1

  // 处理sheet数据
  wb.Sheets[wb.SheetNames[0]] = Object.assign({},
    output, // 导出的内容
    {
      '!ref': `${outputPos[0]}:${outputPos[outputPos.length - 1]}`,
      '!cols': [...colWidth],
      '!merges': merges ? [...merges] : undefined,
      '!freeze': freeze ? [...freeze] : undefined
    }
  )

  // 转成二进制对象
  const tmpDown = new Blob(
    [s2ab(XLSX.write(wb, workBookConfig))],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
  )

  // 下载表格
  downExcel(tmpDown, `${fileName + '.'}${workBookConfig.bookType === 'biff2' ? 'xls' : workBookConfig.bookType}`)
}

/**
 * 生成ASCll值 从A开始
* @param {*} n
*/
function getCharCol (n) {
  if (n > 25) {
    let s = ''
    let m = 0
    while (n > 0) {
      m = n % 26 + 1
      s = String.fromCharCode(m + 64) + s
      n = (n - m) / 26
    }
    return s
  }
  return String.fromCharCode(65 + n)
}

// 字符串转字符流---转化为二进制的数据流
function s2ab (s) {
  if (typeof ArrayBuffer !== 'undefined') {
    const buf = new ArrayBuffer(s.length)
    const view = new Uint8Array(buf)
    for (let i = 0; i !== s.length; ++i) { view[i] = s.charCodeAt(i) & 0xff }
    return buf
  } else {
    const buf = new Array(s.length)
    for (let i = 0; i !== s.length; ++i) { buf[i] = s.charCodeAt(i) & 0xff }
    return buf
  }
}

function downExcel (obj, fileName) {
  const a_node = document.createElement('a')
  a_node.download = fileName
  // 兼容ie
  if ('msSaveOrOpenBlob' in navigator) {
    window.navigator.msSaveOrOpenBlob(obj, fileName)
  } else {
    // 新的对象URL指向执行的File对象或者是Blob对象.
    a_node.href = URL.createObjectURL(obj)
  }
  a_node.click()
  setTimeout(() => {
    URL.revokeObjectURL(obj)
  }, 100)
}

// let ArrayData = [
//   {
//     EndTime: "20:00",
//     StartTime: "08:00",
//     GroupLeaderName: "",
//     GroupLeaderNo: "",
//     WorkGroupName: "接机班组A0001",
//     WorkGroupType: 2,
//     GroupMemberName: "吴伟,操作测试员",
//   },
// ]

// let dataName = [
//   {title: '班组名称', props: 'WorkGroupName'},
//   {title: '班组成员', props: 'GroupMemberName'},
//   {title: '计划开始时段', props: 'StartTime'},
//   {title: '计划结束时段', props: 'EndTime'},
//   {title: '班组类型', props: 'WorkGroupType'},
// ]

export {exportAllData}
