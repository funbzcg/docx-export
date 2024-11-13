/*
 * @Description:
 * @Date: 2024-11-12 10:08:18
 */
/*
 * @Description:
 * @Date: 2024-11-11 13:33:53
 */
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { TIP_TAP_STYLE } from './tipTapStyle.js';
export class Html2Docx {
  constructor(html, option) {
    if (!html) throw new Error('html is required');
    if (!option) throw new Error('option is required');
    if (!option.styleMap) throw new Error('styleMap is required');
    this.html = html;
    this.xmlMap = option.xmlMap;
    this.styleMap = option.styleMap;
    const dom = new DOMParser().parseFromString(html, 'text/html');
    this.domList = [...dom.body.children];
  }
  /**
   * 当有模板传入时初始化样式 暂停方案
   */
  initStyle() {
    const styles = {};
    if (this.styleMap && '标题' in this.styleMap) {
      styles['M_Title'] = this.getTemplateStyle(this.styleMap['标题']);
    }
    if (this.styleMap && '正文' in this.styleMap) {
      styles['M_Text'] = this.getTemplateStyle(this.styleMap['正文']);
      styles['M_H2'] = this.getTemplateStyle(this.styleMap['正文']);
      styles['M_H3'] = this.getTemplateStyle(this.styleMap['正文']);
      styles['M_H4'] = this.getTemplateStyle(this.styleMap['正文']);
      styles['M_H5'] = this.getTemplateStyle(this.styleMap['正文']);
      styles['M_H6'] = this.getTemplateStyle(this.styleMap['正文']);
    }
    return styles;
  }
  /**
   * 根据dom 实时生成样式:用在内部生成样式
   * @param {Object} tempStyle 模板样式，优先使用
   * @param {Object} domStyle dom 样式 当无模板样式时使用
   * @returns {{run:Object,paragraph:Object}} 样式对象
   */
  getStyle(tempStyle, domStyle) {
    const resultStyle = {};
    // 首先获取模板样式
    for (const key in tempStyle) {
      if (Object.prototype.hasOwnProperty.call(tempStyle, key)) {
        switch (key) {
          case 'fontSize':
            resultStyle['run']['size'] = parseInt(tempStyle[key]) * 2;
            break;
          case 'color':
            resultStyle['run']['color'] = tempStyle[key];
            break;
          case 'fontWeight':
            resultStyle['run']['bold'] = tempStyle[fontWeight] === 'bold';
            break;
          case 'fontFamily':
            resultStyle['run']['font'] = {
              name: tempStyle[key],
            };
            break;
          case 'textAlign':
            resultStyle['paragraph']['alignment'] = tempStyle[key];
            break;
          case 'textIndent':
            resultStyle['paragraph']['indent'] = {
              start: tempStyle[key],
            };
            break;
        }
      }
    }
    // 获取dom样式
    for (const key in domStyle) {
      if (Object.prototype.hasOwnProperty.call(domStyle, key)) {
        switch (key) {
          case 'fontSize':
            resultStyle['run']['size'] =
              resultStyle['run']['size'] || parseInt(domStyle[key]) * 2;
            break;
          case 'color':
            resultStyle['run']['color'] =
              resultStyle['run']['color'] || domStyle[key];
            break;
          case 'fontWeight':
            resultStyle['run']['bold'] =
              resultStyle['run']['bold'] || domStyle[key] === 'bold';
            break;
          case 'fontFamily':
            resultStyle['run']['font'] = resultStyle['run']['font'] || {
              name: domStyle[key],
            };
            break;
          case 'textAlign':
            resultStyle['paragraph']['alignment'] =
              resultStyle['paragraph']['alignment'] || domStyle[key];
            break;
          case 'textIndent':
            resultStyle['paragraph']['indent'] = resultStyle['paragraph'][
              'indent'
            ] || {
              start: domStyle[key],
            };
            break;
        }
      }
    }

    return resultStyle;
  }
  /**
   * 统一样式，匹配 TipTap 中的样式
   * @param {Object} tempStyle 模板样式，优先使用
   * @param {Object} defaultStyle dom 样式 当无模板样式时使用
   * @returns {{run:Object,paragraph:Object}} 样式对象
   */
  getTemplateStyle(tempStyle, defaultStyle) {
    return {
      run: {
        color: tempStyle.color || defaultStyle.color,
        size:
          parseInt(tempStyle.fontSize?.match(/\d+/)[0], 10) * 2 ||
          defaultStyle.size,
        bold: tempStyle.fontWeight === 'bold',
        font: {
          name: tempStyle.fontFamily,
        },
      },
      paragraph: {
        alignment: tempStyle.textAlign,
        indent: {
          start: tempStyle.textIndent,
        },
      },
    };
  }

  parseTitle(dom) {
    //优先模板样式 覆盖dom样式
    if (this.styleMap && '标题' in this.styleMap) {
      return new Paragraph({
        text: dom.innerText,
        heading: 'Title',
      });
    }
  }
  parseHTag(dom) {
    //优先模板样式 覆盖dom样式
    if (this.styleMap && '正文' in this.styleMap) {
      return new Paragraph({
        text: dom.innerText,
        heading: 'Heading' + dom.tagName.slice(1),
      });
    }
  }
  parsePTag(dom) {
    if ('正文' in this.styleMap) {
      return new Paragraph({
        text: dom.innerText,
      });
    }
  }
  /**
   * 解析DOM元素
   * 此函数旨在根据DOM元素的类型进行特定处理当前只处理H1元素
   * @param {HTMLElement} dom - 需要解析的DOM元素
   */
  parseDom(dom) {
    // 检查DOM元素是否为H1标签
    if (dom.nodeName === 'H1') {
      return this.parseTitle(dom);
    } else if (/^H[2-6]$/i.test(dom.nodeName)) {
      return this.parseHTag(dom);
    } else {
      return this.parsePTag(dom);
    }
  }
  // 导出DocxBlob
  async exportDocx() {
    const children = [];
    this.domList.forEach((dom) => {
      children.push(this.parseDom(dom));
    });
    const doc = new Document({
      sections: [
        {
          children: [...children],
        },
      ],
      styles: {
        default: {
          document: TIP_TAP_STYLE.p,
          title: TIP_TAP_STYLE.h1,
          heading1: TIP_TAP_STYLE.h1,
          heading2: TIP_TAP_STYLE.h2,
          heading3: TIP_TAP_STYLE.h3,
          heading4: TIP_TAP_STYLE.h4,
          heading5: TIP_TAP_STYLE.h5,
          heading6: TIP_TAP_STYLE.h6,
        },
      },
    });

    // doc.addSection({
    //   children: [...children],
    // });
    console.log(doc.Styles);

    return await Packer.toBlob(doc);
  }
}
