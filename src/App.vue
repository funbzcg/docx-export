<!--
* @Description  测试docx导出
* @FileName  App
* @Date 2024-11-22 14:45:27
!-->
<template>
  <div id="app">
    <input type="file" @input="handleFileUpload" />
    <br />
    <button v-if="show" @click="exportDocx">导出docx</button>
    <CanvasEditor :value="html"></CanvasEditor>
  </div>
</template>
<script setup lang="js">
import { onBeforeMount, ref } from 'vue';
import { Html2Docx } from '@/utils/Html2Docx/index.ts';
import { openDocxAsZip, getBookmarkStyleMap } from '@/utils/readDocx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import CanvasEditor from './components/CanvasEditor.vue';
const jsZip = new JSZip();

const html =
  '<h1 style="line-height: 2.4;text-align: center;margin-top: 17pt;margin-bottom: 16.5pt;"><span style="font-family: 宋体;color: #000000;font-size: 22pt;">培训工作计划</span></h1><h2 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 黑体;color: #000000;font-size: 16pt;">一、引言</span></h2><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">随着企业不断发展壮大，对员工的培训需求也日益增长。为了提高员工的业务能力和综合素质，提升企业整体竞争力，特制定本培训工作计划。本计划旨在明确培训目标、内容、方式及时间安排，以确保培训工作的有序进行。</span></p><h2 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 黑体;color: #000000;font-size: 16pt;">二、培训目标</span></h2><ol>\n<li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">增强员工的业务技能，提高工作效率；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">培养员工的团队协作能力，优化工作流程；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">提升员工的职业素养，树立企业形象；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">激发员工的创新意识，推动企业持续发展。</span></p></li></ol>\n<h2 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 黑体;color: #000000;font-size: 16pt;">三、培训内容</span></h2><h3 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">（一）业务技能培训</span></h3><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">针对员工所在岗位的具体业务需求，开展专业技能培训。包括但不限于：市场营销策略、客户服务技巧、财务管理流程、生产操作技能等。通过培训，使员工熟练掌握业务技能，提高工作效率。</span></p><h3 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">（二）团队协作培训</span></h3><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">通过团队协作培训，增强员工之间的沟通与合作能力。培训内容包括：团队建设理念、沟通技巧、协作方法、冲突解决等。旨在帮助员工建立良好的工作关系，提高团队协作能力。</span></p><h3 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">（三）职业素养培训</span></h3><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">职业素养培训旨在提升员工的职业道德、职业态度和职业行为。培训内容包括：职业礼仪、职场文化、时间管理、自我管理等。通过培训，使员工具备高度的职业素养，为企业树立良好的形象。</span></p><h3 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">（四）创新意识培训</span></h3><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">为了激发员工的创新意识，推动企业持续发展，将开展创新意识培训。培训内容包括：创新思维方法、创新实践案例、创新氛围营造等。旨在帮助员工培养创新意识，提高创新能力。</span></p><h2 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 黑体;color: #000000;font-size: 16pt;">四、培训方式</span></h2><ol>\n<li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">线下培训：组织专业讲师进行现场授课，员工参与互动交流；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">线上培训：利用网络平台，提供视频教程、在线课程等学习资源；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">实践操作：安排员工在实际工作场景中进行操作练习，巩固培训成果；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">分享交流：鼓励员工分享工作经验和心得，促进知识共享。</span></p></li></ol>\n<h2 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 黑体;color: #000000;font-size: 16pt;">五、培训时间安排</span></h2><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">本培训工作计划将按照年度进行规划，具体分为以下几个阶段：</span></p><ol>\n<li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">第一季度：主要开展业务技能培训，针对各岗位员工的具体需求，制定详细的培训计划；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">第二季度：进行团队协作培训，加强员工之间的沟通与合作，提升团队协作能力；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">第三季度：开展职业素养培训，提升员工的职业道德和职业素养，树立企业形象；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">第四季度：进行创新意识培训，激发员工的创新意识，为企业的持续发展注入新的活力。</span></p></li></ol>\n<p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">每个季度的培训时间将根据具体情况进行安排，确保员工能够充分参与并完成培训任务。</span></p><h2 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 黑体;color: #000000;font-size: 16pt;">六、培训效果评估</span></h2><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">为了确保培训工作的有效性，将对培训效果进行评估。评估方式包括：</span></p><ol>\n<li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">考试测评：针对培训内容设置考试题目，检测员工对知识的掌握程度；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">实操考核：对员工在实际工作中的操作进行考核，评估其技能提升情况；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">反馈调查：向员工发放培训反馈问卷，收集员工对培训工作的意见和建议；</span></p></li><li><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">绩效评估：结合员工的工作绩效变化，评估培训成果对企业的影响。</span></p></li></ol>\n<h2 style="line-height: 1.75;margin-top: 13pt;margin-bottom: 13pt;"><span style="font-family: 黑体;color: #000000;font-size: 16pt;">七、总结与展望</span></h2><p style="line-height: 1.5;"><span style="font-family: 宋体;color: #000000;font-size: 12pt;">本培训工作计划旨在全面提高员工的业务能力和综合素质，为企业的发展提供有力的人才保障。通过系统的培训内容和多样化的培训方式，相信能够取得良好的培训效果。未来，我们将根据实际情况对培训计划进行持续优化和完善，以满足企业不断发展的需求。</span></p>';

const show = ref(false);
const templateDocx = ref(null);
const exportDocx = async () => {
  const xmlMap = await openDocxAsZip(templateDocx.value);
  const styleMap = await getBookmarkStyleMap(xmlMap);
  const o = new Html2Docx(html, {
    xmlMap: xmlMap,
    styleMap: styleMap,
  });

  const blob = await o.exportDocx();
  saveAs(blob, 'test.docx');
};
const handleFileUpload = async (event) => {
  const file = event.target.files[0]; // 获取用户选择的文件
  templateDocx.value = await jsZip.loadAsync(file);
  show.value = true;
};
</script>
<style scoped></style>
