/*
 * @Description: 将 json 转换为 docx.js 配置项 接受的pt(镑) 或者 px 应换算为 twip(缇) 
 * @Date: 2024-11-11 13:33:53
 */
import { Document, Packer, Paragraph, LevelFormat, convertInchesToTwip, AlignmentType, TextRun } from 'docx';
import { TIP_TAP_STYLE } from './tipTapStyle';
import type { OptionInterface, StyleMapInterface, TagName, TipTapStyleType, HeadingLevel, ResultStyleInterface, AlignmentTypeString, IBaseParagraphStyleOptions, ISectionOptions, TipTapJson } from './Json2Docx.type';

export class Html2Docx {
  json: TipTapJson;
  styleMap: OptionInterface['styleMap'] | undefined = undefined;
  defaultStyle: TipTapStyleType | undefined;
  numberingConfig: any[];
  listIndex: number = 0;
  selections: ISectionOptions[];
  pageNum: number = -1
  constructor(json: TipTapJson, option: OptionInterface) {
    if (!json) throw new Error('json is required');
    this.styleMap = option.styleMap;
    this.json = json;
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
    this.selections = [];
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
  parseHeader(obj: TipTapJson): Paragraph {
    if (obj.attrs.level === 1) {
      return new Paragraph({
        heading: 'Title',
        children: obj.content.map(item => {
          return this.parseJson(item)
        })
      });
    } else {
      return new Paragraph({})
    }



  }
  parseText(textInfo: TipTapJson): TextRun {
    if (textInfo.type === 'text') {
      return new TextRun({
        text: textInfo.text,
      })
    } else {
      return {
        children: textInfo.content.map(e => {
          return this.parseText(e)
        })
      }
    }
  }
  parseDoc(info: TipTapJson) {
    this.pageNum += 1;
    for (const element of info.content) {
      this.parseJson(element);
    }

  }
  /**
   * @description 解析 json 树形结构
   * @param {TipTapJson} json - json 树形结构
   */
  parseJson(json: TipTapJson) {
    switch (json.type) {
      case 'doc':
        return this.parseDoc(json);
      case 'text':
        return this.parseDoc(json);
      case 'heading':
        return this.parseHeader(json)
      default:

        break;
    }
  }

  // 导出DocxBlob
  async exportDocx() {
    const doc = new Document({
      numbering: {
        config: this.numberingConfig
      },
      sections: this.selections,
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