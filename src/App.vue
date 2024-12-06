<script setup>
import { ref, onMounted, onUnmounted } from 'vue'
import * as XLSX from 'xlsx'

const fileInput = ref(null)
const tableData = ref([])
const tableHeaders = ref([])
const sheets = ref([])
const currentSheet = ref('')
const timer = ref(null)
const currentRole = ref('无')
const roles = ['无', '字幕', '音控', '放像']

const triggerFileInput = () => {
  fileInput.value.click()
}

const loadSheet = (workbook, sheetName) => {
  const sheet = workbook.Sheets[sheetName]
  const options = {
    header: 1,
    raw: false,
    dateNF: 'HH:mm:ss',
    defval: ''
  }
  
  const rawData = XLSX.utils.sheet_to_json(sheet, options)
  
  if (rawData.length > 0) {
    tableHeaders.value = rawData[0]
    
    const roleColumns = {
      '字幕': 5,  // 改为第7列 (索引从0开始，所以是5)
      '音控': 6,  // 改为第8列 (索引从0开始，所以是6)
      '放像': 7   // 改为第9列 (索引从0开始，所以是7)
    }
    
    tableData.value = rawData.slice(1).map(row => {
      const now = new Date()
      const currentTime = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()
      
      const parseTime = (timeStr) => {
        if (!timeStr) return null
        const [hours, minutes, seconds] = timeStr.split(':').map(Number)
        return hours * 3600 + minutes * 60 + (seconds || 0)
      }
      
      const startTime = parseTime(row[1])
      const endTime = parseTime(row[2])
      
      const isActive = startTime !== null && 
                      endTime !== null && 
                      currentTime >= startTime && 
                      currentTime <= endTime

      return {
        data: row.map((cell, index) => cell || ''),
        isActive,
        roleColumn: roleColumns[currentRole.value]
      }
    })
    
    setTimeout(scrollToActiveRow, 100)
  }
}

const handleSheetChange = (event) => {
  const selectedSheet = event.target.value
  currentSheet.value = selectedSheet
  loadSheet(window.currentWorkbook, selectedSheet)
}

const handleFileChange = (event) => {
  const file = event.target.files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result)
    const workbook = XLSX.read(data, { type: 'array' })
    
    // 保存workbook引用以便切换sheet时使用
    window.currentWorkbook = workbook
    
    // 获取所有工作表名称
    sheets.value = workbook.SheetNames
    
    // 默认选择第一个工作表
    if (sheets.value.length > 0) {
      currentSheet.value = sheets.value[0]
      loadSheet(workbook, currentSheet.value)
    }
  }
  
  reader.readAsArrayBuffer(file)
}

const updateActiveStatus = () => {
  if (!window.currentWorkbook || !currentSheet.value) return
  
  // 保存当前的活动行
  const previousActiveRow = document.querySelector('.active-row')
  const previousRowIndex = previousActiveRow ? 
    Array.from(previousActiveRow.parentElement.children).indexOf(previousActiveRow) : -1
  
  loadSheet(window.currentWorkbook, currentSheet.value)
  
  // 获取新的活动行
  const newActiveRow = document.querySelector('.active-row')
  const newRowIndex = newActiveRow ? 
    Array.from(newActiveRow.parentElement.children).indexOf(newActiveRow) : -1
  
  // 只有当活动行改变时才滚动
  if (previousRowIndex !== newRowIndex) {
    setTimeout(scrollToActiveRow, 100)
  }
}

// 添加滚动到活动行的函数
const scrollToActiveRow = () => {
  const activeRow = document.querySelector('.active-row')
  const container = document.querySelector('.table-container')
  if (activeRow && container) {
    const containerHeight = container.clientHeight
    const rowHeight = activeRow.clientHeight
    const activeRowTop = activeRow.offsetTop
    
    // 将高亮行位置调整到更靠近顶部，只留出少量空间
    const targetScroll = activeRowTop - (containerHeight * 0.1) // 改为容器高度的10%
    
    // 减小缓冲区大小，使滚动更精确
    const buffer = containerHeight / 8
    const isInTarget = Math.abs(container.scrollTop - targetScroll) < buffer
    
    if (!isInTarget) {
      container.scrollTo({
        top: Math.max(0, targetScroll),
        behavior: 'smooth'
      })
    }
  }
}

onMounted(() => {
  timer.value = setInterval(updateActiveStatus, 1000)
  setTimeout(scrollToActiveRow, 500)
})

onUnmounted(() => {
  if (timer.value) {
    clearInterval(timer.value)
  }
})
</script>

<template>
  <div>
    <div class="controls">
      <select 
        v-if="sheets.length" 
        v-model="currentSheet" 
        @change="handleSheetChange"
        class="sheet-select"
      >
        <option v-for="sheet in sheets" :key="sheet" :value="sheet">
          {{ sheet }}
        </option>
      </select>
      <select 
        v-if="tableData.length"
        v-model="currentRole"
        class="role-select"
        @change="updateActiveStatus"
      >
        <option v-for="role in roles" :key="role" :value="role">
          {{ role }}
        </option>
      </select>
      <button @click="triggerFileInput">选择 Excel 文件</button>
    </div>

    <input 
      type="file" 
      ref="fileInput" 
      style="display: none" 
      accept=".xlsx, .xls" 
      @change="handleFileChange"
    >

    <div v-if="tableData.length" 
         class="table-container"
         ref="tableContainer">
      <table>
        <thead>
          <tr>
            <th>序号/No.</th>
            <th>时间/TIME</th>
            <th>时长/DUR</th>
            <th>章节/Chapter</th>
            <th>内容/Description</th>
            <th>字幕包/Graphics</th>
            <th>音控/Audio</th>
            <th>放像/VCR</th>
            <th>备注</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(row, rowIndex) in tableData" 
              :key="rowIndex"
              :class="{ 'active-row': row.isActive }">
            <td v-for="(cell, cellIndex) in row.data" 
                :key="cellIndex"
                :class="{
                  'highlight-cell': row.isActive && 
                                  currentRole !== '无' && 
                                  cellIndex === row.roleColumn
                }">
              {{ cell }}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<style scoped>
.controls {
  display: flex;
  gap: 20px;
  align-items: center;
  justify-content: center;
  background: rgba(40, 40, 40, 0.4);
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  padding: 16px;
  border-radius: 12px;
  box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
  min-height: 60px;
  border: 1px solid rgba(255, 255, 255, 0.1);
  max-width: 500px;
  margin: 0 auto;
}

button {
  padding: 10px 20px;
  background: linear-gradient(145deg, #4CAF50, #45a049);
  color: white;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  font-weight: 500;
  transition: all 0.3s ease;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

button:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
}

.sheet-select, .role-select {
  padding: 10px 16px;
  border-radius: 8px;
  border: 1px solid rgba(255, 255, 255, 0.1);
  background: rgba(40, 40, 40, 0.6);
  color: #fff;
  font-size: 14px;
  cursor: pointer;
  transition: all 0.3s ease;
  min-width: 120px;
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
}

.sheet-select:hover, .role-select:hover {
  border-color: rgba(255, 255, 255, 0.2);
  background: rgba(50, 50, 50, 0.7);
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
}

body {
  background-color: #000 !important;
}

.table-container {
  height: 90vh;
  overflow-y: auto;
  margin-top: 1rem;
  background-color: rgba(40, 40, 40, 0.5);
  border-radius: 8px;
  padding: 0;
  position: relative;
}

table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  margin: 0;
  table-layout: fixed;
}

thead {
  position: sticky;
  top: 0;
  z-index: 2;
  background: #2c5282;
}

th:nth-child(1) { width: 6%; }
th:nth-child(2) { width: 8%; }
th:nth-child(3) { width: 8%; }
th:nth-child(4) { width: 10%; }
th:nth-child(5) { width: 25%; }
th:nth-child(6) { width: 10%; }
th:nth-child(7) { width: 11%; }
th:nth-child(8) { width: 11%; }
th:nth-child(9) { width: 11%; }

thead th {
  padding: 15px 8px;
  color: #fff;
  font-weight: 600;
  text-align: center;
  border: 1px solid rgba(255, 255, 255, 0.1);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

tbody tr:nth-child(odd) {
  background-color: rgba(60, 60, 60, 0.6);
}

tbody tr:nth-child(even) {
  background-color: rgba(50, 50, 50, 0.6);
}

td {
  padding: 12px 8px;
  border: 1px solid rgba(255, 255, 255, 0.05);
  color: rgba(255, 255, 255, 0.9);
  text-align: center;
  white-space: normal;
  word-break: break-word;
  min-height: 40px;
  max-height: none;
  vertical-align: middle;
}

.active-row {
  background: rgba(146, 45, 35, 0.8) !important;
}

.highlight-cell {
  background: rgba(255, 69, 58, 0.95) !important;
  color: #ffffff !important;
  font-weight: bold;
  text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3);
  box-shadow: inset 0 0 10px rgba(255, 255, 255, 0.1);
}

tbody tr:not(.active-row):hover {
  background-color: rgba(70, 70, 70, 0.8);
}

@keyframes highlight {
  0% { 
    background: linear-gradient(90deg, rgba(231, 76, 60, 0.4), rgba(192, 57, 43, 0.3));
  }
  50% { 
    background: linear-gradient(90deg, rgba(231, 76, 60, 0.5), rgba(192, 57, 43, 0.4));
  }
  100% { 
    background: linear-gradient(90deg, rgba(231, 76, 60, 0.4), rgba(192, 57, 43, 0.3));
  }
}

@keyframes cell-highlight {
  0% { 
    background: linear-gradient(145deg, #1890ff, #096dd9);
  }
  50% { 
    background: linear-gradient(145deg, #096dd9, #1890ff);
  }
  100% { 
    background: linear-gradient(145deg, #1890ff, #096dd9);
  }
}

@media (max-width: 768px) {
  .controls {
    flex-direction: column;
    padding: 12px;
  }
  
  .sheet-select, .role-select {
    width: 100%;
  }
  
  .table-container {
    padding: 8px;
    height: calc(100vh - 180px);
  }
}

.file-select {
  padding: 16px;
}

.centered {
  display: flex;
  justify-content: center;
  align-items: center;
  min-height: 90vh;
}

/* 修改第一行第一个单元格为直角 */
tr:first-child td:first-child {
  border-top-left-radius: 0;
}

/* 修改最后一行第一个单元格为直角 */
tr:last-child td:first-child {
  border-bottom-left-radius: 0;
}

/* 确保所有角都是直角 */
tr:first-child td:first-child,
tr:first-child td:last-child,
tr:last-child td:first-child,
tr:last-child td:last-child {
  border-radius: 0;
}

/* 移除单元格的圆角 */
td {
  border-radius: 0 !important;
}

/* 如果需要显示省略号，可以使用这个样式 */
td.truncate {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
</style>
