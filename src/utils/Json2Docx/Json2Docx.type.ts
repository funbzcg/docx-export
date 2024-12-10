/*
 * @Description: 类型
 * @Date: 2024-11-22 15:07:16
 */
import { AlignmentType, ISectionOptions, ISpacingProperties, HeadingLevel, IBaseParagraphStyleOptions, Paragraph } from 'docx'

export type TagName = 'H1' | 'H2' | 'H3' | 'H4' | 'H5' | 'H6' | 'P' | 'OL' | 'UL' | 'LI' | 'BR' | 'A' | 'IMG' | 'TABLE' | 'TR' | 'TD' | 'TH' | 'B' | 'I' | 'U' | 'STRONG' | 'EM' | 'CODE' | 'PRE';


/**
 * @description:  样式映射
 */
export interface StyleMapInterface {
  fontSize: string; // 字体大小
  color: string; // 字体颜色
  fontWeight: string; // 字体粗细
  fontFamily: string; // 字体
  textAlign: string; // 对齐方式
  textIndent: string; // 首行缩进
  lineHeight: string; // 行高
  [key: string]: string;
}

export interface TipTapJson {
  type: string;
  marks?: any[];
  attrs?: any;
  content: TipTapJson[];
  text?: string;
}

enum StyleMapKeyType {
  Title = '标题',
  Body = '正文',
  Footer = '页脚',
}

/**@description:  构造函数配置项*/
export interface OptionInterface {
  styleMap: Partial<Record<StyleMapKeyType, Partial<StyleMapInterface>>>;
}
/** @description:  字体样式 */
interface RunStyleInterface {
  size: number | string;
  color: string;
  bold: boolean;
  font: string; // 字体 docx中接受三种类型 这里只接受字符串
}

export type AlignmentTypeString = (typeof AlignmentType)[keyof typeof AlignmentType];

/** @description:  段落样式*/
interface ParagraphStyleInterface {
  alignment: AlignmentTypeString;
  spacing: ISpacingProperties;
  indent: Object;
}
/** @description:  交由docx接收的样式 */
export interface ResultStyleInterface {
  run?: Partial<RunStyleInterface>,
  paragraph?: Partial<ParagraphStyleInterface>
}



export type TipTapStyleType = Partial<Record<TagName, ResultStyleInterface>>;

export { HeadingLevel, }

export type { IBaseParagraphStyleOptions, ISectionOptions }