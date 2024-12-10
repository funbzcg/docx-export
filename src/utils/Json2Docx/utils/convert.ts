/*
 * @Description: 单位转换工具
 * @Date: 2024-12-02 14:09:38
 */
import { convertInchesToTwip } from 'docx';


function convertPtToTwip(pt: number) {
  // 1 pt = 1/720 英寸
  return convertInchesToTwip(pt / 720);
}
function convertPxToTwip(px: number) {
  // 1 px 的大小取决于设备的分辨率，但在Web标准中通常认为 1 px ≈ 1/96 英寸
  return convertInchesToTwip(px / 960);
}

export {
  convertPtToTwip, // 点值转换
  convertPxToTwip, // 像素值转换
  convertInchesToTwip, // 英寸转换
}