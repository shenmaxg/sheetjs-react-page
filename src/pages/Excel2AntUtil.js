import * as xlsx from 'xlsx';

const rc = {
  r: '行',
  c: '列',
};

const xlsxUtils = xlsx.utils;
const {
  decode_range,
  encode_row,
  decode_row,
  encode_col,
  decode_col,
  format_cell,
  split_cell,
} = xlsxUtils;

/**
 *  合并参数
 *  @param arg1 可以是对象，可以是数组，也可以是空
 *  @param arg2 可以是对象，可以是数组，也可以是空
 *  @return 数组或对象或空
 */
function merge_arguments(arg1, arg2) {
  if (Array.isArray(arg1)) {
    if (Array.isArray(arg2)) {
      return arg1.concat(arg2);
    } else {
      arg1.push(arg2);

      return arg1;
    }
  } else if (Array.isArray(arg2)) {
    arg2.push(arg1);

    return arg2;
  } else if (arg1 && arg2) {
    return [arg1, arg2];
  }

  return arg1 || arg2;
}

/**
 *  解析 excel 文档中单元格范围
 *  @param sheet 文档
 *  @return {object} s 表示 start e 表示 end
 */
function patch_decode_range(sheet) {
  const range = decode_range(sheet['!ref']);
  const ref = sheet['!ref'].split(':');

  // 说明系统解析出的 ref 是错误的
  if (!sheet[ref[0]] || !sheet[ref[1]]) {
    const rawNumber = [];
    const colNumber = [];

    for (let key in sheet) {
      if (!key.startsWith('!')) {
        const cellCodes = split_cell(key);

        rawNumber.push(decode_row(cellCodes[1]));
        colNumber.push(decode_col(cellCodes[0]));
      }
    }
    rawNumber.sort((a, b) => {
      return a - b;
    });
    colNumber.sort((a, b) => {
        return a - b;
    });

    if (!sheet[ref[0]]) {
      range.s = {
        r: rawNumber[0],
        c: colNumber[0],
      };
    }

    if (!sheet[ref[1]]) {
      range.e = {
        r: rawNumber[rawNumber.length - 1],
        c: colNumber[colNumber.length - 1],
      };
    }
  }

  return range;
}

/**
 *  多行数据合并一行，如果存在 key 相同，则 value 拼接为数组
 *  @param keyList {Array} 参数列表
 *  @param prev {object} 上个返回值
 *  @param cur {object} 当前数据
 *  @return {object} 合并后的数据
 */
function merge_ant_data_reducer(keyList, prev, cur) {
  if (prev) {
    for (let key of keyList) {
      prev[key] = merge_arguments(prev[key], cur[key]);
    }

    return prev;
  } else {
    return cur;
  }
}

/**
 *  计算 Excel 列合并或者行合并数
 *  @param merge eg: [{e: {c: 0, r: 2}, s: {c: 0, r: 1}}]
 *  @param rawOrCol r 表示行，c 表示列
 *  @return span {number} 合并数，1 表示没有合并
 */
function get_merge_num(merge, rawOrCol) {
  let span = 1;
  const validateField = rawOrCol === 'r' ? 'c' : 'r';

  if (merge) {
    if (merge.s[validateField] !== merge.e[validateField]) {
      throw `第${merge.s[validateField] + 1}${
        rc[validateField]
      }解析错误!【表头不允许行合并，数据不允许列合并】`;
    } else {
      span = merge.e[rawOrCol] - merge.s[rawOrCol] + 1;
    }
  }

  return span;
}

/**
 *  merges 是 sheet 中的一个属性，表示文档中需要合并单元格的部分
 *  @param merges eg: [{e: {c: 0, r: 2}, s: {c: 0, r: 1}}]
 *  @return mergedMap {Map} 将 s 中的 c 和 r 转换成 Excel 标准坐标（A2）,作为 key
 */
function merges_to_map(merges) {
  const mergedMap = new Map();

  if (merges) {
    merges.forEach(merge => {
      const index = encode_col(merge.s.c) + encode_row(merge.s.r);
      mergedMap.set(index, merge);
    });
  }

  return mergedMap;
}

/**
 *  生成 ant table 中某个字段的配置对象
 *  @param title 列头显示文字
 *  @param dataIndex 列数据在数据项中对应的路径
 *  @param colSpan 表头列合并,设置为 0 时，不渲染
 *  @param colNum 对应 Excel 中的列数
 *  @return {object} 装入参拼接为 ant column 指定的格式
 */
function make_ant_column(title, dataIndex, colSpan, colNum) {
  return {
    title,
    dataIndex,
    colSpan,
    _colIndex: encode_col(colNum),
    _rawSpanList: [],
  };
}

/**
 *  为 column 字段增加 render 方法
 *  @param range excel 数据范围
 *  @param column ant table 中某个字段的配置对象
 */
function patch_column_render(range, column) {
  const rawSpanList = column['_rawSpanList'];

  column.render = (value, row, index) => {
    const obj = {
      children: value,
      props: {},
    };

    rawSpanList.forEach(({ rawIndex, rawSpan }) => {
      const index2RawIndex = index + range.s.r + 2;

      if (index2RawIndex === rawIndex) {
        obj.props.rowSpan  = rawSpan;
      } else if (
        index2RawIndex > rawIndex &&
        index2RawIndex < rawIndex + rawSpan
      ) {
        obj.props.rowSpan  = 0;
      }
    });

    return obj;
  };
}

/**
 *  blob 转为 workbook，由于是文件的读取，返回值是一个 Promise
 *  @param blob {blob} 文件二进制流
 *  @return {Promise} 可获得 workbook 对象
 */
export function blob_to_wb(blob) {
  return new Promise((resolve, reject) => {
    const fileReader = new FileReader();

    fileReader.onload = event => {
      try {
        const { result } = event.target;
        const workbook = xlsx.read(result, {
          type: 'binary',
          dateNF: 'yyyy-mm-dd',
        });

        resolve(workbook);
      } catch (e) {
        reject('文件类型不正确');
      }
    };

    fileReader.readAsBinaryString(blob);
  });
}

/**
 *  根据 sheet 数据解析生成 ant table 需要的 columns 数组。
 *  @param sheet excel 工作薄对象
 *  @return columns {Array} ant table 需要的 columns 数组
 */
export function sheet_to_ant_columns(sheet) {
  if (sheet == null || sheet['!ref'] == null) return [];

  const merges = sheet['!merges'];
  const mergedMap = merges_to_map(merges);
  const range = patch_decode_range(sheet);
  const headerRow = encode_row(range.s.r);
  const columns = [];

  // 处理表头列合并
  for (let i = range.s.c; i <= range.e.c; i++) {
    const cellCol = encode_col(i);
    const cellIndex = cellCol + headerRow;
    const headerCell = sheet[cellIndex];

    if (headerCell) {
      const value = format_cell(headerCell);
      const colSpan = get_merge_num(mergedMap.get(cellIndex), 'c');

      for (let j = 1; j <= colSpan; j++) {
        const dataIndex = colSpan === 1 ? value : `${value}_${j}`;
        const columnSpan = j === 1 ? colSpan : 0;
        const column = make_ant_column(value, dataIndex, columnSpan, i + j - 1);

        columns.push(column);
      }
    }
  }

  // 处理数据行合并
  for (let i = range.s.r + 1; i <= range.e.r; i++) {
    const cellRowIndex = encode_row(i);

    columns.forEach(column => {
      const cellIndex = column['_colIndex'] + cellRowIndex;
      const dataCell = sheet[cellIndex];

      if (dataCell) {
        const rawSpan = get_merge_num(mergedMap.get(cellIndex), 'r');

        column['_rawSpanList'].push({
          rawIndex: parseInt(cellRowIndex, 10),
          startRaw: range.s.r,
          rawSpan,
        });
      }
    });
  }

  // 根据 column 的 _rawSpanList 属性生成 column render 方法
  columns.forEach(patch_column_render.bind(null, range));

  return columns;
}

/**
 *  根据 sheet 数据解析生成 ant table 需要的 dataSource，空数据默认 null。
 *  @param sheet excel 工作薄对象
 *  @return dataSource {Array} ant table 需要的 dataSource 数组
 */
export function sheet_to_ant_dataSource(sheet, columns) {
  const range = patch_decode_range(sheet);
  const dataSource = [];

  for (let i = range.s.r + 1; i <= range.e.r; i++) {
    const cellRowIndex = encode_row(i);
    const data = {};

    data.key = i;
    columns.forEach(column => {
      const cellIndex = column['_colIndex'] + cellRowIndex;
      const cell = sheet[cellIndex];

      if (cell) {
        data[column.dataIndex] = format_cell(cell);
      } else {
        data[column.dataIndex] = '';
      }
    });

    dataSource.push(data);
  }

  return dataSource;
}

/**
 *  将 ant table 数据转换为 json 格式
 *  @param columns table 表头定义
 *  @param dataSource table 数据
 *  @return jsonArray {Array} json 数据，该数据用来上传后台
 */
export function ant_table_to_json(columns, dataSource) {
  const jsonArray = [];

  if (columns && dataSource) {
    const keySet = new Set();
    const dataIndexToTitle = {};
    const jsonArrayAfterColSpan = [];

    columns.forEach(column => {
      dataIndexToTitle[column.dataIndex] = column.title;
      keySet.add(column.title);
    });

    // 处理数据头部列合并
    dataSource.forEach(data => {
      const json = {};

      for (let key in data) {
        if (key !== 'key') {
          const title = dataIndexToTitle[key];

          if (data[key]) {
            json[title] = merge_arguments(json[title], data[key]);
          }
        }
      }

      jsonArrayAfterColSpan.push(json);
    });

    // 处理数据中的行合并
    const firstColumn = columns[0];
    const firstRawSpanList = firstColumn['_rawSpanList'];

    firstRawSpanList.forEach(rawSpanObj => {
      let json = {};
      const { rawIndex, rawSpan, startRaw } = rawSpanObj;
      const dataIndex = rawIndex - startRaw - 2;

      if (rawSpan > 1) {
        const mergedData = jsonArrayAfterColSpan.slice(dataIndex, dataIndex + rawSpan);

        json = mergedData.reduce(
          merge_ant_data_reducer.bind(null, keySet),
          null,
        );
      } else {
        json = jsonArrayAfterColSpan[dataIndex];
      }

      jsonArray.push(json);
    });

    return jsonArray;
  }

  return null;
}
