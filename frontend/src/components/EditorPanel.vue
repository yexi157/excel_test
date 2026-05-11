<script setup lang="ts">
import ToolbarBar from './ToolbarBar.vue'
import UniverHost from './UniverHost.vue'
</script>

<template>
  <div class="editor-panel">
    <ToolbarBar />
    <div class="editor-host">
      <UniverHost />
    </div>
  </div>
</template>

<style scoped>
.editor-panel { height: 100%; display: flex; flex-direction: column; }
.editor-host { flex: 1; min-height: 0; }
</style>
