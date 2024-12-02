<!--
* @Description  canvas 编辑器
* @FileName  CanvasEditor
* @Date 2024-11-27 16:30:52
!-->
<template>
  <div id="canvasEditor" class="canvas-editor"></div>
</template>
<script setup lang="ts">
import { onBeforeMount, onMounted, ref } from 'vue';
import Editor, { getElementListByHTML } from '@hufe921/canvas-editor';
let editor: Editor;
const props = defineProps({
  value: {
    type: String,
    default: '',
  },
});
onMounted(() => {
  const dom = document.getElementById('canvasEditor') as HTMLDivElement;
  const elementList = getElementListByHTML(props.value, {
    innerWidth: 800,
  });
  editor = new Editor(dom, elementList, {});
  console.log(editor);
});
onBeforeMount(() => {});
</script>
<style scoped>
.canvas-editor {
  width: 100vw;
  height: 100vh;
}
</style>
