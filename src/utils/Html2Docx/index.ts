/*
 * @Description: 将 html 转换为 docx.js 配置项
 * @Date: 2024-11-11 13:33:53
 */
import { Document, Packer, Paragraph, TextRun, } from 'docx';
import { TIP_TAP_STYLE } from './tipTapStyle';
import type { OptionInterface, StyleMapInterface, TagName, TipTapStyleType, HeadingLevel, ResultStyleInterface, AlignmentTypeString, IBaseParagraphStyleOptions } from './Html2Docx.type';
export class Html2Docx {
  html: string;
  domList: HTMLElement[];
  styleMap: OptionInterface['styleMap'];
  defaultStyle: TipTapStyleType | undefined;
  constructor(html: string, option: OptionInterface) {
    if (!html) throw new Error('html is required');
    if (!option) throw new Error('option is required');
    if (!option.styleMap) throw new Error('styleMap is required');
    this.html = html;
    this.styleMap = option.styleMap;
    const dom = new DOMParser().parseFromString(html, 'text/html');
    this.domList = Array.from(dom.body.children) as HTMLElement[];
    this.initDefaultStyle();
  }
  initDefaultStyle() {
    const copy = deepCopy<TipTapStyleType>(TIP_TAP_STYLE);
    for (let key in copy) {
      if (key === 'H1') {
        if (this.styleMap && '标题' in this.styleMap) {
          copy[key] = this.getTemplateStyle(this.styleMap['标题'] as StyleMapInterface, key);
        } else {
          continue;
        }
      } else {
        copy[key as TagName] = this.getTemplateStyle(this.styleMap['正文'] as StyleMapInterface, key as TagName);
      }
    }
    this.defaultStyle = copy;
  }
  /**
   * 统一样式，匹配 TipTap 中的样式
   * @param {Object} tempStyle 模板样式，优先使用
   * @param {Object} defaultStyle dom 样式 当无模板样式时使用
   * @returns {{run:Object,paragraph:Object}} 样式对象
   */
  getTemplateStyle(tempStyle: StyleMapInterface, tag: TagName): ResultStyleInterface {
    const resultStyle = TIP_TAP_STYLE[tag];
    // 确保 resultStyle 的 run 和 paragraph 存在
    if (!(resultStyle!.run)) {
      resultStyle!.run = {};
    }
    if (!(resultStyle!.paragraph)) {
      resultStyle!.paragraph = {};
    }
    // 首先获取模板样式
    for (const key in tempStyle) {
      if (tempStyle.hasOwnProperty(key)) {
        const value = tempStyle[key];
        if (value === undefined || value === null) {
          continue; // 跳过无效值
        }

        switch (key) {
          case 'fontSize':
            if (!isNaN(parseInt(value))) {
              resultStyle!.run.size = parseInt(value) * 2;
            }
            break;
          case 'color':
            resultStyle!.run.color = value;
            break;
          case 'fontWeight':
            resultStyle!.run.bold = value === 'bold';
            break;
          case 'fontFamily':
            resultStyle!.run.font = value;
            break;
          case 'textAlign':
            resultStyle!.paragraph.alignment = value as AlignmentTypeString;
            break;
        }
      }
    }
    return resultStyle as ResultStyleInterface;
  }
  parseTitle(dom: HTMLElement) {
    return new Paragraph({
      text: dom.innerText,
      heading: 'Title',
    });
  }
  parseHTag(dom: HTMLElement) {
    const tag = dom.tagName;
    return new Paragraph({
      text: dom.innerText,
      heading: ('Heading' + dom.tagName.slice(1)) as (typeof HeadingLevel)[keyof typeof HeadingLevel],
    });
  }
  parsePTag(dom: HTMLElement) {
    return new Paragraph({
      text: dom.innerText,
    });
  }
  parseOlTag(dom: HTMLElement) { }
  /**
   * 解析DOM元素
   * 此函数旨在根据DOM元素的类型进行特定处理当前只处理H1元素
   * @param {HTMLElement} dom - 需要解析的DOM元素
   */
  parseDom(dom: HTMLElement) {
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
    const children: Paragraph[] = [];
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
          document: this.defaultStyle!.P as IBaseParagraphStyleOptions,
          title: this.defaultStyle!.H1 as IBaseParagraphStyleOptions,
          heading1: this.defaultStyle!.H1 as IBaseParagraphStyleOptions,
          heading2: this.defaultStyle!.H2 as IBaseParagraphStyleOptions,
          heading3: this.defaultStyle!.H3 as IBaseParagraphStyleOptions,
          heading4: this.defaultStyle!.H4 as IBaseParagraphStyleOptions,
          heading5: this.defaultStyle!.H5 as IBaseParagraphStyleOptions,
          heading6: this.defaultStyle!.H6 as IBaseParagraphStyleOptions,
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

function deepCopy<T>(obj: T): T {
  return JSON.parse(JSON.stringify(obj));
}