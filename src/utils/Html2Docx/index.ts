/*
 * @Description: 将 html 转换为 docx.js 配置项 接受的pt(镑) 或者 px 应换算为 twip(缇) 
 * @Date: 2024-11-11 13:33:53
 */
import { Document, Packer, Paragraph, LevelFormat, convertInchesToTwip, AlignmentType } from 'docx';
import { TIP_TAP_STYLE } from './tipTapStyle';
import type { OptionInterface, StyleMapInterface, TagName, TipTapStyleType, HeadingLevel, ResultStyleInterface, AlignmentTypeString, IBaseParagraphStyleOptions, ISectionOptions } from './Html2Docx.type';

export class Html2Docx {
  html: string;
  domList: HTMLElement[];
  styleMap: OptionInterface['styleMap'] | undefined = undefined;
  defaultStyle: TipTapStyleType | undefined;
  numberingConfig: any[];
  listIndex: number = 0;
  constructor(html: string, option: OptionInterface) {
    if (!html) throw new Error('html is required');
    this.styleMap = option.styleMap;
    this.html = html;
    const dom = new DOMParser().parseFromString(html, 'text/html');
    this.domList = Array.from(dom.body.children) as HTMLElement[];
    this.numberingConfig = [{
      reference: "unique",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "\u1F60",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) },
            },
          },
        },
      ],
    },]
  }
  initDefaultStyle() {
    /** 执行该函数时 this.styleMap 不为null */
    const copy = deepCopy<TipTapStyleType>(TIP_TAP_STYLE);
    if (!this.styleMap) {
      return copy
    }
    for (let key in copy) {
      if (key === 'H1') {
        if (this.styleMap && '标题' in this.styleMap) {
          copy[key] = this.getTemplateStyle(this.styleMap['标题'] as StyleMapInterface, key);
        } else {
          continue;
        }
      } else {
        copy[key as TagName] = this.getTemplateStyle(this.styleMap!['正文'] as StyleMapInterface, key as TagName);
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
    if (dom.children.length) {
      const lineDomList = Array.from(dom.children) as HTMLElement[];
      return new Paragraph({
        heading: 'Title',
        // children: [this.parseLineDomList(lineDomList)]
      })
    }
    return new Paragraph({
      text: dom.innerText,
      heading: 'Title',
    });
  }
  parseHTag(dom: HTMLElement) {
    return new Paragraph({
      text: dom.innerText,
      heading: ('Heading' + dom.tagName.slice(1)) as (typeof HeadingLevel)[keyof typeof HeadingLevel],
    });
  }
  /**
   * 
   * @param dom ol 列表dom
   * @param level { number } 层级，当出现嵌套列表时，level 会自动加 1
   * @returns {Paragraph}
   */
  parseOlTag(dom: HTMLElement, level: number = 0): Paragraph[] {
    this.listIndex++
    //通过 判断索引奇偶性 区分 reference
    const children = Array.from(dom.children) as HTMLElement[];
    this.addNumbering(this.listIndex);
    let paragraphList: Paragraph[] = [];
    for (const element of children) {
      if (element.nodeName === 'OL') {
        paragraphList = paragraphList.concat(this.parseOlTag(element, level + 1));
      } else if (element.nodeName === 'UL') {
        paragraphList = paragraphList.concat(this.parseUlTag(element, level + 1))
      } else if (element.nodeName === 'LI') {
        const text = element.innerText;
        paragraphList.push(new Paragraph({
          text,
          numbering: {
            reference: "number" + this.listIndex,
            level,
          }
        }))
      }
    }
    return paragraphList
  }
  /**
 * 
 * @param dom ol 列表dom
 * @param level { number } 层级，当出现嵌套列表时，level 会自动加 1
 * @returns {Paragraph}
 */
  parseUlTag(dom: HTMLElement, level: number = 0): Paragraph[] {
    //通过 判断索引奇偶性 区分 reference
    const children = Array.from(dom.children) as HTMLElement[];
    let paragraphList: Paragraph[] = [];
    for (const element of children) {
      if (element.nodeName === 'OL') {
        paragraphList = paragraphList.concat(this.parseOlTag(element, level + 1));
      } else if (element.nodeName === 'UL') {
        paragraphList = paragraphList.concat(this.parseUlTag(element, level + 1))
      } else if (element.nodeName === 'LI') {
        const text = element.innerText;
        paragraphList.push(new Paragraph({
          text,
          numbering: {
            reference: "unique",
            level,
          }
        }))
      }
    }
    return paragraphList
  }
  parsePTag(dom: HTMLElement) {
    return new Paragraph({
      text: dom.innerText,
    });
  }

  /**
   * 解析DOM元素
   * 此函数旨在根据DOM元素的类型进行特定处理当前只处理H1元素
   * @param {HTMLElement[]} domList - 需要解析的DOM元素
   */
  parseDomList(domList: HTMLElement[]): ISectionOptions[] {
    console.log(domList);

    // selection docx 的单页，默认不主动分页；需要分页时，selections push 个 {children:[]} index =+ 1
    let selections: { children: Paragraph[] }[] = [{ children: [] }];
    let index = 0;
    for (const dom of domList) {
      // 检查DOM元素是否为H1标签
      if (dom.nodeName === 'H1') {
        selections[index].children.push(this.parseTitle(dom));
      } else if (/^H[2-6]$/i.test(dom.nodeName)) {
        selections[index].children.push(this.parseHTag(dom));
      } else if (dom.nodeName === 'OL') {
        selections[index].children = selections[index].children.concat(this.parseOlTag(dom));
      } else if (dom.nodeName === 'UL') {
        selections[index].children = selections[index].children.concat(this.parseUlTag(dom));
      } else {
        selections[index].children.push(this.parsePTag(dom));
      }
    }
    return selections
  }
  /**
 * @description 相同 reference 的数字序列号会延续，因此出现了几个数字列表就需要几个reference
 * @param {number} index - 列表数量，默认为1
 */
  addNumbering(index: number) {
    this.numberingConfig.push({
      reference: "number" + index,
      levels: [
        {
          level: 0,
          format: LevelFormat.DECIMAL,
          text: "%1.",
          alignment: AlignmentType.START,
          start: 1,
          style: {
            paragraph: {
              indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.18) },
            },
          },
        },
      ],
    })
  }


  // 导出DocxBlob
  async exportDocx() {

    const doc = new Document({
      numbering: {
        config: this.numberingConfig
      },
      sections: this.parseDomList(this.domList),
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

    return await Packer.toBlob(doc);
  }
}

function deepCopy<T>(obj: T): T {
  return JSON.parse(JSON.stringify(obj));
}