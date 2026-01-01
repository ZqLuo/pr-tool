<script setup>
import { ref } from 'vue'
import * as XLSX from 'xlsx'

// 数据状态
const excelData = ref([])
const columns = ref([])
const selectedUrl = ref('')
const selectedRowIndex = ref(-1) // 当前选中的行索引
const rowStatus = ref({}) // 行状态：{ index: 'confirm' | 'delete' | 'pending' }
const loading = ref(false)

// 文件选择处理
const handleFileChange = (event) => {
  const file = event.target.files[0]
  if (!file) return
  
  loading.value = true
  const reader = new FileReader()
  
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result)
      const workbook = XLSX.read(data, { type: 'array' })
      const firstSheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[firstSheetName]
      const jsonData = XLSX.utils.sheet_to_json(worksheet)
      
      excelData.value = jsonData
      if (jsonData.length > 0) {
        columns.value = Object.keys(jsonData[0])
      }
      // 重置状态
      selectedRowIndex.value = -1
      selectedUrl.value = ''
      rowStatus.value = {}
    } catch (error) {
      console.error('读取Excel文件失败:', error)
      alert('读取Excel文件失败，请检查文件格式')
    } finally {
      loading.value = false
    }
  }
  
  reader.readAsArrayBuffer(file)
}

// 处理链接点击
const handleLinkClick = (url, index) => {
  selectedUrl.value = url
  selectedRowIndex.value = index
}

// 处理行操作
const handleRowAction = (action) => {
  if (selectedRowIndex.value === -1) return
  
  if (action === 'delete') {
    // 删除行
    excelData.value.splice(selectedRowIndex.value, 1)
    // 更新状态映射
    const newStatus = {}
    Object.keys(rowStatus.value).forEach(key => {
      const index = parseInt(key)
      if (index < selectedRowIndex.value) {
        newStatus[index] = rowStatus.value[key]
      } else if (index > selectedRowIndex.value) {
        newStatus[index - 1] = rowStatus.value[key]
      }
    })
    rowStatus.value = newStatus
    selectedRowIndex.value = -1
    selectedUrl.value = ''
  } else {
    // 设置行状态
    rowStatus.value[selectedRowIndex.value] = action
  }
}

// 判断是否为链接
const isLink = (value) => {
  return typeof value === 'string' && /^https?:\/\//i.test(value)
}

// 获取行状态类名
const getRowClass = (index) => {
  const status = rowStatus.value[index]
  if (status === 'confirm') return 'row-confirm'
  if (status === 'pending') return 'row-pending'
  return ''
}

// 导出Excel
const exportExcel = () => {
  if (excelData.length === 0) {
    alert('没有数据可以导出')
    return
  }
  
  try {
    const worksheet = XLSX.utils.json_to_sheet(excelData.value)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')
    XLSX.writeFile(workbook, '导出结果.xlsx')
  } catch (error) {
    console.error('导出Excel失败:', error)
    alert('导出Excel失败')
  }
}
</script>

<template>
  <div class="app-container">
    <header class="header">
      <h1>Excel内容展示</h1>
      <div class="header-actions">
        <input 
          type="file" 
          accept=".xlsx, .xls" 
          @change="handleFileChange"
          class="file-input"
        />
        <button 
          @click="exportExcel"
          class="export-btn"
          :disabled="excelData.length === 0"
        >
          导出Excel
        </button>
      </div>
    </header>
    
    <main class="main-content">
      <!-- 左侧Excel表格 -->
      <div class="left-panel">
        <div v-if="loading" class="loading">加载中...</div>
        <div v-else-if="excelData.length === 0" class="empty-state">
          <p>请选择一个Excel文件</p>
        </div>
        <div v-else>
          <!-- 操作选项 -->
          <div v-if="selectedRowIndex !== -1" class="action-options">
            <h3>当前选择行：{{ selectedRowIndex + 1 }}</h3>
            <div class="action-buttons">
              <button 
                @click="handleRowAction('confirm')"
                class="action-btn confirm-btn"
              >
                确定
              </button>
              <button 
                @click="handleRowAction('delete')"
                class="action-btn delete-btn"
              >
                删除
              </button>
              <button 
                @click="handleRowAction('pending')"
                class="action-btn pending-btn"
              >
                待定
              </button>
            </div>
          </div>
          
          <!-- 表格容器 -->
          <div class="table-container">
            <!-- Excel表格 -->
            <table class="excel-table">
              <thead>
                <tr>
                  <th v-for="col in columns" :key="col">{{ col }}</th>
                </tr>
              </thead>
              <tbody>
                <tr 
                  v-for="(row, index) in excelData" 
                  :key="index"
                  :class="[getRowClass(index), { 'row-selected': index === selectedRowIndex }]"
                >
                  <td v-for="col in columns" :key="col">
                    <template v-if="isLink(row[col])">
                      <a 
                        href="#" 
                        @click.prevent="handleLinkClick(row[col], index)"
                        class="excel-link"
                      >
                        {{ row[col] }}
                      </a>
                    </template>
                    <span v-else>{{ row[col] }}</span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
      
      <!-- 右侧网页预览 -->
      <div class="right-panel">
        <div v-if="!selectedUrl" class="preview-empty">
          <p>点击左侧表格中的链接查看网页内容</p>
        </div>
        <iframe 
          v-else 
          :src="selectedUrl" 
          class="web-preview"
          title="网页预览"
          sandbox="allow-scripts allow-same-origin allow-popups allow-forms"
        ></iframe>
      </div>
    </main>
  </div>
</template>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
  background-color: #f5f7fa;
  color: #333;
}

.app-container {
  width: 100vw;
  height: 100vh;
  display: flex;
  flex-direction: column;
}

.header {
  background-color: #fff;
  padding: 1rem 2rem;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.header h1 {
  font-size: 1.5rem;
  font-weight: 600;
  color: #42b883;
}

.header-actions {
  display: flex;
  gap: 1rem;
  align-items: center;
}

.file-input {
  padding: 0.5rem 1rem;
  border: 2px solid #e1e4e8;
  border-radius: 4px;
  cursor: pointer;
  transition: border-color 0.2s;
}

.file-input:hover {
  border-color: #42b883;
}

.export-btn {
  padding: 0.5rem 1rem;
  background-color: #42b883;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-weight: 500;
  transition: background-color 0.2s;
}

.export-btn:hover:not(:disabled) {
  background-color: #36a675;
}

.export-btn:disabled {
  background-color: #a0e2c0;
  cursor: not-allowed;
}

.main-content {
  flex: 1;
  display: flex;
  overflow: hidden;
}

.left-panel {
  flex: 1;
  background-color: #fff;
  margin: 1rem;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  overflow: auto;
  padding: 1rem;
  min-width: 0;
  position: relative;
}

.loading {
  text-align: center;
  padding: 2rem;
  color: #666;
}

.empty-state {
  text-align: center;
  padding: 2rem;
  color: #999;
}

/* 操作选项样式 */
.action-options {
  background-color: #f9fafb;
  padding: 1rem;
  border-radius: 6px;
  margin-bottom: 1rem;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
  position: sticky;
  top: 0;
  z-index: 20;
  backdrop-filter: blur(4px);
  background-color: rgba(249, 250, 251, 0.95);
  width: fit-content;
  min-width: 100%;
}

/* 表格容器样式 */
.table-container {
  overflow-x: auto;
  border-radius: 4px;
}

.action-options h3 {
  font-size: 1rem;
  font-weight: 600;
  color: #374151;
  margin-bottom: 0.75rem;
}

.action-buttons {
  display: flex;
  gap: 0.75rem;
  flex-wrap: wrap;
}

.action-btn {
  padding: 0.5rem 1rem;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-weight: 500;
  transition: all 0.2s;
  min-width: 80px;
}

.confirm-btn {
  background-color: #10b981;
  color: white;
}

.confirm-btn:hover {
  background-color: #059669;
}

.delete-btn {
  background-color: #ef4444;
  color: white;
}

.delete-btn:hover {
  background-color: #dc2626;
}

.pending-btn {
  background-color: #f59e0b;
  color: white;
}

.pending-btn:hover {
  background-color: #d97706;
}

/* Excel表格样式 */
.excel-table {
  min-width: 100%;
  border-collapse: collapse;
  font-size: 0.9rem;
}

.excel-table th {
  background-color: #f9fafb;
  border-bottom: 2px solid #e1e4e8;
  padding: 0.75rem;
  text-align: left;
  font-weight: 600;
  color: #374151;
  position: sticky;
  top: 0;
  z-index: 10;
  white-space: nowrap;
  min-width: 100px;
}

.excel-table td {
  border-bottom: 1px solid #e5e7eb;
  padding: 0.75rem;
  color: #4b5563;
  white-space: nowrap;
  min-width: 100px;
}

/* 行状态样式 */
.excel-table tr:hover {
  background-color: #f9fafb;
}

.row-selected {
  background-color: #dbeafe !important;
  border-left: 4px solid #3b82f6;
}

.row-confirm {
  background-color: #d1fae5;
}

.row-pending {
  background-color: #fef3c7;
}

.excel-link {
  color: #3b82f6;
  text-decoration: none;
  cursor: pointer;
  transition: color 0.2s;
}

.excel-link:hover {
  color: #2563eb;
  text-decoration: underline;
}

.right-panel {
  width: 50%;
  background-color: #fff;
  margin: 1rem 1rem 1rem 0;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  overflow: hidden;
  display: flex;
  flex-direction: column;
}

.preview-empty {
  flex: 1;
  display: flex;
  justify-content: center;
  align-items: center;
  color: #999;
  background-color: #f9fafb;
}

.web-preview {
  width: 100%;
  height: 100%;
  border: none;
}
</style>
