/*
 * @Description: 将 html 转换为 docx.js 配置项 接受的pt(镑) 或者 px 应换算为 twip(缇) 
 * @Date: 2024-11-11 13:33:53
 */
import { Document, Packer, Paragraph, FrameAnchorType, TextRun, BorderStyle, HorizontalPositionAlign, VerticalPositionAlign } from 'docx';

export class ToTextBox {
  html: string;
  domList: HTMLElement[];
  constructor(html: string) {
    if (!html) throw new Error('html is required');
    this.html = html;
    const dom = new DOMParser().parseFromString(html, 'text/html');
    this.domList = Array.from(dom.body.children) as HTMLElement[];
  }
  parseDom(dom: HTMLElement, x: number, y: number) {
    if (y === 9000) {
      return new Paragraph({
        text: dom.innerText,
      })
    }
    return new Paragraph({
      shading: {
        fill: '#FFFF00'
      },
      frame: {
        type: "absolute",
        position: {
          x,
          y,
        },
        width: 4000,
        height: 1000,
        anchor: {
          horizontal: FrameAnchorType.MARGIN,
          vertical: FrameAnchorType.MARGIN,
        },

      },
      border: {
        top: {
          color: "auto",
          space: 1,
          style: BorderStyle.NONE,
          size: 6,
        },
        bottom: {
          color: "auto",
          space: 1,
          style: BorderStyle.NONE,
          size: 6,
        },
        left: {
          color: "auto",
          space: 1,
          style: BorderStyle.NONE,
          size: 6,
        },
        right: {
          color: "auto",
          space: 1,
          style: BorderStyle.NONE,
          size: 6,
        },
      },
      children: [
        new TextRun(dom.innerText),
      ]

    });

  }
  /**
   * 解析DOM元素
   * 此函数旨在根据DOM元素的类型进行特定处理当前只处理H1元素
   * @param {HTMLElement[]} domList - 需要解析的DOM元素
   */
  parseDomList(domList: HTMLElement[]): any[] {
    // selection docx 的单页，默认不主动分页；需要分页时，selections push 个 {children:[]} index =+ 1
    let selections: { children: Paragraph[] }[] = [{ children: [] }];
    let index = 0;
    let distanceIndex = 0;
    /** 主要测试 */
    for (const i in domList) {
      let instance = (Number(i) - distanceIndex) * 3000;
      const dom = domList[i]
      if (instance > (16839 - 3000)) {
        index += 1;
        selections.push({ children: [] })
        instance = 0;
        distanceIndex = Number(i);
      }
      selections[index].children.push(this.parseDom(dom, 1000, instance))
    }
    return selections
  }
  // 导出DocxBlob
  async exportDocx() {
    console.log(this.domList);
    const doc = new Document({
      sections: this.parseDomList(this.domList),
    });
    return await Packer.toBlob(doc);
  }
}