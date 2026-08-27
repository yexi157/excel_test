import { ElMessageBox } from 'element-plus'
import { useFileStore } from '@/stores/fileStore'

export type GuardResult = 'continue' | 'cancel'

/**
 * 切换/关闭前的 dirty 守卫（spec §6.2）
 * 返回 'continue' 表示用户已处理（保存或放弃），可以继续
 * 返回 'cancel' 表示用户取消，应中止当前操作
 */
export async function guardDirty(): Promise<GuardResult> {
  const store = useFileStore()
  if (!store.dirty) return 'continue'

  try {
    const action = await ElMessageBox.confirm(
      `当前文件「${store.currentFileName ?? ''}」有未保存的改动`,
      '继续操作？',
      {
        distinguishCancelAndClose: true,
        confirmButtonText: '保存',
        cancelButtonText: '放弃',
        type: 'warning',
      },
    )
    if (action === 'confirm') {
      await store.save()
      return 'continue'
    }
    return 'continue'
  } catch (action) {
    if (action === 'cancel') return 'continue'   // 放弃改动
    return 'cancel'                              // close / esc
  }
}
