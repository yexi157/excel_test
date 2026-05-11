<script setup lang="ts">
import { onMounted } from 'vue'
import FileTreePanel from '@/components/FileTreePanel.vue'
import EditorPanel from '@/components/EditorPanel.vue'
import { useFileStore } from '@/stores/fileStore'

const store = useFileStore()
onMounted(async () => {
  await store.refreshTree()
})
</script>

<template>
  <el-container class="app-root">
    <el-aside width="280px" class="aside">
      <FileTreePanel />
    </el-aside>
    <el-main class="main">
      <EditorPanel />
    </el-main>
  </el-container>
</template>

<style scoped>
.app-root { height: 100vh; }
.aside { border-right: 1px solid #e4e7ed; padding: 0; }
.main { padding: 0; }
</style>
