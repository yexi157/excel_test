<script setup lang="ts">
import { computed, ref } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import type { FileTreeNode } from '@/api/types'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
const treeData = computed(() => store.treeNodes)
const treeProps = { children: 'children', label: 'name' }

async function onNodeClick(node: FileTreeNode) {
  if (node.type !== 'file') return
  try {
    await store.openFile(node.id)
  } catch (e) {
    console.error('[openFile] failed:', e)
  }
}

// 新建文件（在选中文件夹下；如果选中是文件，则在其父文件夹下）
async function newFile(parentNode: FileTreeNode | null) {
  const parentId = parentNode?.type === 'folder' ? parentNode.id : (parentNode?.parentId ?? null)
  try {
    const { value } = await ElMessageBox.prompt('文件名（自动加 .xlsx 后缀）', '新建文件', {
      inputPattern: /^[^\\/:*?"<>|]+$/,
      inputErrorMessage: '名称含非法字符',
    })
    await store.createNewSheet(parentId, value)
    ElMessage.success('已创建')
  } catch (e: any) {
    if (e === 'cancel' || e === 'close') return
    ElMessage.error(`创建失败：${e.message ?? e}`)
  }
}

const contextNode = ref<FileTreeNode | null>(null)
const menuVisible = ref(false)
const menuStyle = ref<{ top: string; left: string }>({ top: '0px', left: '0px' })

function onContext(e: MouseEvent, _data: FileTreeNode, node: any) {
  e.preventDefault()
  contextNode.value = node.data
  menuStyle.value = { top: `${e.clientY}px`, left: `${e.clientX}px` }
  menuVisible.value = true
}

function closeMenu() { menuVisible.value = false }

function newFileFromContext() {
  closeMenu()
  newFile(contextNode.value)
}
</script>

<template>
  <div class="tree-panel" @click="closeMenu">
    <div class="tree-panel-header">
      <span>文件</span>
      <el-button size="small" text @click="newFile(null)">+ 新建</el-button>
    </div>
    <el-tree
      :data="treeData"
      :props="treeProps"
      node-key="id"
      highlight-current
      @node-click="onNodeClick"
      @node-contextmenu="onContext"
    >
      <template #default="{ node, data }">
        <span class="tree-node">
          <span :class="data.type === 'folder' ? 'icon-folder' : 'icon-file'" />
          {{ node.label }}
        </span>
      </template>
    </el-tree>

    <div v-if="menuVisible" class="ctx-menu" :style="menuStyle" @click.stop>
      <div class="ctx-item" @click="newFileFromContext">新建文件</div>
      <!-- task 5.7 增加：重命名 / 删除 / 新建文件夹 -->
    </div>
  </div>
</template>

<style scoped>
.tree-panel { height: 100%; display: flex; flex-direction: column; position: relative; }
.tree-panel-header {
  height: 48px; padding: 0 16px;
  display: flex; align-items: center; justify-content: space-between;
  border-bottom: 1px solid #e4e7ed; font-weight: 600;
}
.tree-node { display: inline-flex; align-items: center; gap: 6px; }
.icon-folder::before { content: '📁'; }
.icon-file::before { content: '📄'; }
.ctx-menu {
  position: fixed; background: white; border: 1px solid #e4e7ed; border-radius: 4px;
  box-shadow: 0 2px 12px rgba(0,0,0,0.1); z-index: 9999; min-width: 120px;
}
.ctx-item { padding: 8px 16px; cursor: pointer; }
.ctx-item:hover { background: #f5f7fa; }
</style>
