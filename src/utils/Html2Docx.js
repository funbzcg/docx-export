/*
 * @Description:
 * @Date: 2024-11-11 13:33:53
 */
import { Document, Packer, Paragraph, TextRun } from 'docx';
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
  // 引入模板时会先初始化样式，无模板时采用自身Dom样式
  // initStyle() {
  //   if (this.styleMap && '标题' in this.styleMap) {
  //     const style = this.styleMap['标题'];
  //     this.heading1Style = {
  //       run: {
  //         color: style.color,
  //         size: parseInt(style.fontSize?.match(/\d+/)[0], 10) * 2,
  //         bold: style.fontWeight === 'bold',
  //         font: {
  //           name: style.fontFamily,
  //         },
  //       },
  //       paragraph: {
  //         alignment: style.textAlign,
  //         indent: {
  //           start: style.textIndent,
  //         },
  //       },
  //     };
  //   }
  //   if (this.styleMap && '正文' in this.styleMap) {
  //     const style = this.styleMap['正文'];
  //     this.contentStyle = {
  //       run: {
  //         color: style.color,
  //         size: parseInt(style.fontSize?.match(/\d+/)[0], 10) * 2,
  //         bold: style.fontWeight === 'bold',
  //         font: {
  //           name: style.fontFamily,
  //         },
  //       },
  //       paragraph: {
  //         alignment: style.textAlign,
  //         indent: {
  //           start: style.textIndent,
  //         },
  //       },
  //     };
  //     this.heading = {
  //       run: {
  //         color: style.color,
  //         size: parseInt(style.fontSize?.match(/\d+/)[0], 10) * 2,
  //         bold: style.fontWeight === 'bold',
  //         font: {
  //           name: style.fontFamily,
  //         },
  //       },
  //       alignment: 'start',
  //       indent: {
  //         start: style.textIndent,
  //       },
  //     };
  //   }
  // }
  parseTitle(dom) {
    let style = {};
    const domStyle = dom.getAttribute('style');
    console.log(domStyle);
    if (domStyle) {
      Object.assign(style, {
        run: {
          color: domStyle.color,
          size: parseInt(domStyle.fontSize?.match(/\d+/)[0], 10) * 2,
          bold: domStyle.fontWeight === 'bold',
          font: {
            name: domStyle.fontFamily,
          },
        },
        alignment: domStyle.textAlign,
        indent: {
          start: domStyle.textIndent,
        },
      });
    }
    //优先模板样式 覆盖dom样式
    if (this.styleMap && '标题' in this.styleMap) {
      const tempStyle = this.styleMap['标题'];
      Object.assign(style, {
        run: {
          ...style.run,
          color: tempStyle.color,
          size: parseInt(tempStyle.fontSize?.match(/\d+/)[0], 10) * 2,
          bold: tempStyle.fontWeight === 'bold',
          font: {
            name: tempStyle.fontFamily,
          },
        },
        alignment: tempStyle.textAlign,
        indent: {
          start: tempStyle.textIndent,
        },
      });
    }
    return new Paragraph({
      text: dom.innerText,
      ...style,
    });
  }
  parseH(dom) {
    let style = {};
    const domStyle = dom.getAttribute('style');
    console.log(domStyle);
    if (domStyle) {
      style = {
        run: {
          color: domStyle.color,
          size: parseInt(domStyle.fontSize?.match(/\d+/)[0], 10) * 2,
          bold: domStyle.fontWeight === 'bold',
          font: {
            name: domStyle.fontFamily,
          },
        },
        alignment: domStyle.textAlign,
        indent: {
          start: domStyle.textIndent,
        },
      };
    }
    //优先模板样式 覆盖dom样式
    if (this.styleMap && '正文' in this.styleMap) {
      const tempStyle = this.styleMap['正文'];
      style = {
        ...style,
        run: {
          ...style.run,
          color: tempStyle.color,
          size: parseInt(tempStyle.fontSize?.match(/\d+/)[0], 10) * 2,
          bold: tempStyle.fontWeight === 'bold',
          font: {
            name: tempStyle.fontFamily,
          },
        },
        alignment: tempStyle.textAlign,
        indent: {
          start: tempStyle.textIndent,
        },
      };
    }
    if ('正文' in this.styleMap) {
      return new Paragraph({
        text: dom.innerText,
        ...style,
      });
    }
  }
  parseP(dom) {
    let style = {};
    const domStyle = dom.getAttribute('style');
    console.log(domStyle);
    if (domStyle) {
      style = {
        run: {
          color: domStyle.color,
          size: parseInt(domStyle.fontSize?.match(/\d+/)[0], 10) * 2,
          bold: domStyle.fontWeight === 'bold',
          font: {
            name: domStyle.fontFamily,
          },
        },
        alignment: domStyle.textAlign,
        indent: {
          start: domStyle.textIndent,
        },
      };
    }
    //优先模板样式 覆盖dom样式
    if (this.styleMap && '正文' in this.styleMap) {
      const tempStyle = this.styleMap['正文'];
      style = {
        ...style,
        run: {
          ...style.run,
          color: tempStyle.color,
          size: parseInt(tempStyle.fontSize?.match(/\d+/)[0], 10) * 2,
          bold: tempStyle.fontWeight === 'bold',
          font: {
            name: tempStyle.fontFamily,
          },
        },
        alignment: tempStyle.textAlign,
        indent: {
          start: tempStyle.textIndent,
        },
      };
    }
    if ('正文' in this.styleMap) {
      return new Paragraph({
        text: dom.innerText,
        ...style,
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
      return this.parseH(dom);
    } else {
      return this.parseP(dom);
    }
  }
  // 导出DocxBlob
  async exportDocx() {
    const children = [];
    this.domList.forEach((dom) => {
      children.push(this.parseDom(dom));
    });
    const doc = new Document({
      sections: [],
    });
    doc.addSection({
      children: [...children],
    });
    console.log(doc.Styles);

    return await Packer.toBlob(doc);
  }
}
