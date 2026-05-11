<script setup lang="ts">
import { computed } from 'vue'
import { ElMessage } from 'element-plus'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()

const stateLabel = computed(() => {
  if (store.saveState === 'saving') return '保存中...'
  if (store.saveState === 'saved') return '已保存'
  if (store.saveState === 'error') return '保存失败'
  if (store.dirty) return '未保存'
  return store.currentFileId ? '已保存' : ''
})

const stateColor = computed(() => {
  if (store.saveState === 'error') return '#f56c6c'
  if (store.dirty || store.saveState === 'saving') return '#e6a23c'
  if (store.currentFileId) return '#67c23a'
  return '#909399'
})

// 占位 handlers，task 5.4 onUpload / task 5.5 onDownload
async function onSave() {
  try {
    await store.save()
    ElMessage.success('保存成功')
  } catch (e: any) {
    ElMessage.error(`保存失败：${e.message ?? e}`)
  }
}
function onUpload() { /* task 5.4 */ }
function onDownload() { /* task 5.5 */ }
</script>

<template>
  <div class="toolbar">
    <div class="toolbar-left">
      <span class="filename">{{ store.currentFileName || '未打开文件' }}</span>
      <span class="dot" :style="{ background: stateColor }" />
      <span class="state">{{ stateLabel }}</span>
    </div>
    <div class="toolbar-right">
      <el-button size="small" :disabled="!store.currentFileId" @click="onSave">保存</el-button>
      <el-button size="small" @click="onUpload">上传 .xlsx</el-button>
      <el-button size="small" :disabled="!store.currentFileId" @click="onDownload">下载</el-button>
    </div>
  </div>
</template>

<style scoped>
.toolbar {
  height: 48px;
  padding: 0 16px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-bottom: 1px solid #e4e7ed;
}
.toolbar-left { display: flex; align-items: center; gap: 8px; }
.filename { font-weight: 600; }
.dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
.state { font-size: 12px; color: #606266; }
.toolbar-right { display: flex; gap: 8px; }
</style>
