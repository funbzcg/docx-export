// TipTap 一些基础样式 待完善
/**
 * @description: TipTap 样式 数字样式在导出的docx中会少一半，因此这里统一*2
 */
const TIP_TAP_STYLE = {
  H1: {
    run: {
      color: '#000000',
      size: 2 * 22,
      bold: true,
      font: {
        name: '宋体',
      },
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
      size: 2 * 16,
      bold: true,
      font: {
        name: '宋体',
      },
    },
    paragraph: {
      alignment: 'left',
      indent: {
        start: 0,
      },
      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.75,
      },
    },
  },
  H3: {
    run: {
      color: '#000000',
      size: 2 * 12,
      bold: true,
      font: {
        name: '宋体',
      },
    },
    paragraph: {
      alignment: 'left',
      indent: {
        start: 0,
      },
      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.75,
      },
    },
  },
  H4: {
    run: {
      color: '#000000',
      size: 2 * 12,
      bold: true,
      font: {
        name: '黑体',
      },
    },
    paragraph: {
      alignment: 'left',
      indent: {
        start: 0,
      },
      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.55,
      },
    },
  },
  H5: {
    run: {
      color: '#000000',
      size: 2 * 12,
      bold: true,
      font: {
        name: '宋体',
      },
    },
    paragraph: {
      alignment: 'left',
      indent: {
        start: 0,
      },
      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.55,
      },
    },
  },
  H6: {
    run: {
      color: '#000000',
      size: 2 * 12,
      bold: true,
      font: {
        name: '黑体',
      },
    },
    paragraph: {
      alignment: 'left',
      indent: {
        start: 0,
      },
      spacing: {
        before: 200 * 1.32,
        line: 200 * 1.32,
      },
    },
  },
  P: {
    run: {
      color: '#000000',
      size: 2 * 12,
      bold: false,
      font: {
        name: '宋体',
      },
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
};

export { TIP_TAP_STYLE };
