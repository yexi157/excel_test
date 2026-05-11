<script setup lang="ts">
import { computed } from 'vue'
import type { FileTreeNode } from '@/api/types'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
const treeData = computed(() => store.treeNodes)

const treeProps = {
  children: 'children',
  label: 'name',
}

function onNodeClick(node: FileTreeNode) {
  if (node.type !== 'file') return
  // task 5.1 实现 openFile
  console.log('[FileTreePanel] click', node.id, node.name)
}

// task 5.7 实现右键菜单（重命名/删除/新建文件夹/新建文件）
</script>

<template>
  <div class="tree-panel">
    <div class="tree-panel-header">
      <span>文件</span>
    </div>
    <el-tree
      :data="treeData"
      :props="treeProps"
      node-key="id"
      highlight-current
      @node-click="onNodeClick"
    >
      <template #default="{ node, data }">
        <span class="tree-node">
          <span :class="data.type === 'folder' ? 'icon-folder' : 'icon-file'" />
          {{ node.label }}
        </span>
      </template>
    </el-tree>
  </div>
</template>

<style scoped>
.tree-panel { height: 100%; display: flex; flex-direction: column; }
.tree-panel-header {
  height: 48px;
  padding: 0 16px;
  display: flex;
  align-items: center;
  border-bottom: 1px solid #e4e7ed;
  font-weight: 600;
}
.tree-node { display: inline-flex; align-items: center; gap: 6px; }
.icon-folder::before { content: '📁'; }
.icon-file::before { content: '📄'; }
</style>
