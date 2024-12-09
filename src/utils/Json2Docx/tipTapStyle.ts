// TipTap 一些基础样式 待完善

import { TipTapStyleType } from "./Json2Docx.type";
import { convertPtToTwip, } from "./utils";

/**
 * @description: TipTap 样式 数字样式在导出的docx中会少一半，因此这里统一*2
 */
const TIP_TAP_STYLE: TipTapStyleType = {
  H1: {
    run: {
      color: '#000000',
      size: convertPtToTwip(22),
      bold: true,
      font: '宋体',
    },
    paragraph: {
      alignment: 'center',
      spacing: {
        before: 200 * 1.32,
        line: 200 * 2.4,
      },
    },
  },
  H2: {
    run: {
      color: '#000000',
      size: convertPtToTwip(16),
      bold: true,
      font: '宋体',
    },
    paragraph: {
      alignment: 'left',

      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.75,
      },
    },
  },
  H3: {
    run: {
      color: '#000000',
      size: convertPtToTwip(12),
      bold: true,
      font: '宋体',
    },
    paragraph: {
      alignment: 'left',

      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.75,
      },
    },
  },
  H4: {
    run: {
      color: '#000000',
      size: convertPtToTwip(12),
      bold: true,
      font: '黑体',
    },
    paragraph: {
      alignment: 'left',

      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.55,
      },
    },
  },
  H5: {
    run: {
      color: '#000000',
      size: convertPtToTwip(12),
      bold: true,
      font: '宋体',
    },
    paragraph: {
      alignment: 'left',

      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.55,
      },
    },
  },
  H6: {
    run: {
      color: '#000000',
      size: convertPtToTwip(12),
      bold: true,
      font: '黑体',
    },
    paragraph: {
      alignment: 'left',

      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.32,
      },
    },
  },
  P: {
    run: {
      color: '#000000',
      size: convertPtToTwip(12),
      bold: false,
      font: '宋体',
    },
    paragraph: {
      alignment: 'left',
      indent: {
        left: 200, //不起效果
      },
      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.5,
      },
    },
  },
  LI: {
    run: {
      color: '#000000',
      size: convertPtToTwip(12),
      bold: false,
      font: '宋体',
    },
    paragraph: {
      alignment: 'left',
      indent: {
        left: 200, //不起效果
      },
      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.5,
      },
    },
  }
};

export { TIP_TAP_STYLE };
