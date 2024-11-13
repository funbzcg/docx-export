/*
 * @Description:读取xml文件的数据和样式
 * @Date: 2024-11-12 13:27:17
 */
/**
 * 异步获取书签样式映射
 *
 * 该函数接收一个包含文档和样式的XML映射对象，然后解析文档以提取书签，并将每个书签与其对应的样式映射
 * 它首先检查文档是否存在，如果不存在则返回一个空对象如果文档存在，它将解析XML文档并获取所有书签起始标签
 * 对于每个书签，除了名称为'_GoBack'的书签外，它都会获取书签的父节点，并使用父节点和样式数据来确定书签的样式
 * 最后，它将每个书签与其样式作为一个键值对添加到结果对象中并返回该对象
 *
 * @param {Object} xmlMap - 包含'document'和'styles'属性的对象，分别代表文档的XML字符串和样式信息
 * @returns {Promise<Object>} - 返回一个Promise，解析为一个对象，其中键是书签名称，值是对应的样式对象
 */
const getBookmarkStyleMap = async (xmlMap) => {
  // 从xmlMap参数中解构出document和styles属性
  const { document, styles } = xmlMap;
  // 初始化一个空对象来存储书签名称和对应的样式
  const result = {};
  // 如果没有文档，则直接返回空对象
  if (!document) return result;
  // 使用DOMParser解析XML文档字符串，以便于后续操作
  const xmlDom = new DOMParser().parseFromString(document, 'text/xml');
  // 获取所有书签起始标签元素
  const domList = xmlDom.getElementsByTagName('w:bookmarkStart');

  // 调用getStyleData函数处理样式信息，以方便后续查找样式
  const styleData = getStyleData(styles);

  // 遍历所有书签起始标签元素
  for (let i = 0; i < domList.length; i++) {
    // 获取当前遍历到的书签节点
    const bmNode = domList[i];
    // 获取书签的名称属性
    const name = bmNode.getAttribute('w:name');
    // 如果书签名称是'_GoBack'，则跳过当前循环迭代
    if (name === '_GoBack') continue;

    // 获取当前书签的父节点
    const pNode = bmNode.parentNode;
    // 使用父节点和样式数据来获取书签的样式，并将书签名称和样式作为键值对添加到结果对象中
    result[name] = getStyle(pNode, styleData);
  }

  // 返回包含所有书签和对应样式的对象
  return result;
};

/**
 * 将docx文件作为zip文件打开并处理
 * 此函数旨在从docx文件中提取文档内容、样式和图片，并以对象形式返回
 * @param {Object} docx - 包含docx文件的zip对象
 * @returns {Promise<Object>} - 返回一个Promise，解析为包含文档、样式和图片的字符串对象
 * @throws {Error} - 如果文件处理过程中发生错误，抛出错误
 */
const openDocxAsZip = async (docx) => {
  try {
    // 获取zip文件中的所有文件
    const { files } = docx;
    // 初始化结果对象，用于存储处理后的文档、样式和图片
    const result = {
      document: '',
      styles: '',
    };
    // 遍历zip文件中的每个文件
    for (const key in files) {
      // 确保当前文件是zip文件的一部分
      if (Object.prototype.hasOwnProperty.call(files, key)) {
        // 检查并提取文档内容
        if (key.includes('word/document.xml')) {
          result['document'] = await zipFileToString(files[key] || null);
        }
        // 检查并提取样式内容
        if (key.includes('word/styles.xml')) {
          result['styles'] = await zipFileToString(files[key] || null);
        }
      }
    }
    // 返回处理后的结果对象
    return result;
  } catch (e) {
    // 如果发生错误，抛出带有错误信息的异常
    throw new Error('this file is not a docx file!', { cause: 'fileError' });
  }
};

/**
 * 从styles.xml文件中提取paragraph段落的样式表
 * @param {string | Document} stylesXml styles.xml文档or提取的xml字符串
 */
const getStyleData = (stylesXml) => {
  if (typeof stylesXml === 'string') {
    stylesXml = new DOMParser().parseFromString(stylesXml, 'text/xml');
  }
  const styleList = queryAll(stylesXml, 'w:style');
  return styleList.reduce(
    (res, curNode) => {
      const type = curNode.getAttribute('w:type');
      const id = curNode.getAttribute('w:styleId');
      switch (type) {
        case 'paragraph': {
          const data = {
            default: curNode.getAttribute('w:default'),
            name: getVal(curNode, 'w:name'),
            uiPriority: getVal(curNode, 'w:uiPriority'),
            style: getStyle(curNode),
          };
          res.paragraph[id] = data;
          if (truthyList.includes(data.default)) {
            Object.defineProperty(res.paragraph, 'default', {
              value: data,
              configurable: false,
            });
          }
        }
      }

      return res;
    },
    {
      paragraph: {},
      character: {},
      table: {},
    }
  );
};

const zipFileToString = async (zipFile) => {
  if (!zipFile) return '';
  const ab = await zipFile.async('arraybuffer');
  const decoder = new TextDecoder('utf-8');
  return decoder.decode(ab);
};

const query = (node, prop) => node?.getElementsByTagName(prop)?.[0] || null;
const queryAll = (node, prop) =>
  node ? [...node.getElementsByTagName(prop)] : [];
const getVal = (node, prop) => {
  const el = query(node, prop);
  if (!el) return;
  return el.getAttribute('w:val');
};
const setVal = (node, val) => {
  node.setAttribute('w:val', val);
};

const getStyle = (node, styleData) => {
  const pPr = query(node, 'w:pPr');
  const rPr = query(node, 'w:rPr');
  const pStyle = getVal(pPr, 'w:pStyle');
  const style = {};
  if (styleData) {
    const defaultStyle = pStyle
      ? styleData.paragraph[pStyle]?.style
      : styleData.paragraph['default']?.style;

    Object.assign(style, defaultStyle || {});
  }

  if (pPr) {
    // 首行缩进
    const ind = query(pPr, 'w:ind');
    if (ind) {
      const flc = ind.getAttribute('w:firstLineChars');
      if (flc && +flc !== 0) {
        style.textIndent = flc / 100 + 'em';
      } else {
        const fl = ind.getAttribute('w:firstLine');
        if (fl && +fl !== 0) style.textIndent = flc / 567 + 'cm';
      }
    }

    // 水平对齐方式
    const jc = query(pPr, 'w:jc');
    if (jc) {
      const val = jc.getAttribute('w:val');
      if (val && val !== 'both') {
        const map = {
          left: 'left',
          right: 'right',
          center: 'center',
          distribute: 'justify',
        };
        style.textAlign = map[val];
      }
    }

    // 行间距and段落间距
    const spacing = query(pPr, 'w:spacing');
    if (spacing) {
      // 行距
      const lineRule = spacing.getAttribute('w:lineRule');
      const line = spacing.getAttribute('w:line');
      if (line && +line !== 0) {
        style.lineHeight =
          lineRule === 'exact' || lineRule === 'atLeast'
            ? point(line) // 固定值or最小值
            : line / 2.4 + '%'; // auto -> 单倍行距or多倍行距
      } else {
        style.lineHeight = 1.15;
      }
      // 段前段后
      /** before or after / 20 pt
       *  beforeLines or afterLines / 100 * line
       */
      const before = spacing.getAttribute('w:before');
      if (before && +before !== 0) style.marginTop = point(before);
      const after = spacing.getAttribute('w:after');
      if (after && +after !== 0) style.marginBottom = point(after);
    }
  }

  if (rPr) {
    // 字体
    /* 此处暂不考虑 <w:cs> 和 styles.xml */
    const rFonts = query(rPr, 'w:rFonts');
    if (rFonts) {
      const hint = rFonts.getAttribute('w:hint');
      const fontMap = {
        ascii: rFonts.getAttribute('w:ascii'),
        hAnsi: rFonts.getAttribute('w:hAnsi'),
        eastAsia: rFonts.getAttribute('w:eastAsia'),
        cs: rFonts.getAttribute('w:cs'),
      };
      const fonts = [];
      if (hint in fontMap) fonts.push(fontMap[hint]);
      for (const key in fontMap) {
        if (Object.hasOwnProperty.call(fontMap, key)) {
          const font = fontMap[key];
          if (font && !fonts.includes(font)) fonts.push(font);
        }
      }
      const val = fonts.join(',');
      if (val) {
        style.fontFamily = val;
      }
    }

    // 字号
    const sz = query(rPr, 'w:sz');
    const szCs = query(rPr, 'w:szCs');
    if (sz) {
      const val = sz.getAttribute('w:val');
      if (val && +val !== 0) style.fontSize = val / 2 + 'pt';
      else if (szCs) {
        const val = szCs.getAttribute('w:val');
        if (val && +val !== 0) style.fontSize = val / 2 + 'pt';
      }
    }

    // 加粗
    const b = query(rPr, 'w:b');
    // 暂不考虑bCs
    // const bCs = query(rPr, "w:bCs");
    if (b) {
      const val = b.getAttribute('w:val');
      if (!val || (val !== '0' && val !== 'false')) style.fontWeight = 'bold';
      // else if (bCs) {
      //   const val = bCs.getAttribute("w:val");
      //   if (!val || (val !== "0" && val !== "false"))
      //     style.fontWeight = "bold";
      // }
    }

    // 斜体
    const i = query(rPr, 'w:i');
    // const iCs = query(rPr, "w:iCs");
    if (i) {
      const val = i.getAttribute('w:val');
      if (!val || (val !== '0' && val !== 'false')) style.fontStyle = 'italic';
      // else if (iCs) {
      //   const val = iCs.getAttribute("w:val");
      //   if (!val || (val !== "0" && val !== "false"))
      //     style.fontStyle = "italic";
      // }
    }

    // 下划线
    // const u = query(rPr, "w:u");
    // if (u) {
    //
    // }

    // 颜色
    const color = query(rPr, 'w:color');
    if (color) {
      const val = color.getAttribute('w:val');
      // 暂不考虑主题色
      // const themeColor = color.getAttribute("w:themeColor");
      if (val) {
        if (val !== 'auto')
          style.color = hexColorReg.test(val) ? `#${val}` : val;
      }
    }

    // 背景色
    const highlight = query(rPr, 'w:highlight');
    if (highlight) {
      const val = color.getAttribute('w:val');
      if (val) {
        style.backgroundColor = hexColorReg.test(val) ? `#${val}` : val;
      }
    }
  }

  return style;
};

const point = (val) => val / 20 + 'pt';
const truthyList = ['1', 'on', 'true'];
const hexColorReg = /^[0-9a-f]{6}$/i;

export { getBookmarkStyleMap, openDocxAsZip };
