<template>
  <div class="app-container">
    <!-- Sidebar -->
    <aside class="sidebar">
      <div class="brand">
        <h1>BigBaby</h1>
        <span>Toolbox (Vue 3)</span>
      </div>
      <nav class="nav-menu">
        <div class="nav-item" :class="{ active: currentTab === 'bank-reconcile' }"
          @click="currentTab = 'bank-reconcile'">
          <span class="icon">💰</span>
          集团银行明细对账
        </div>
        <div class="nav-item disabled">
          <span class="icon">🚧</span>
          更多功能开发中...
        </div>
      </nav>
    </aside>

    <!-- Main Content -->
    <main class="main-content">
      <header class="top-bar">
        <h2>{{ currentTitle }}</h2>
      </header>

      <section class="content-body">
        <BankReconciliation v-if="currentTab === 'bank-reconcile'" :onLog="addLog" />
      </section>

      <!-- Console / Log Panel -->
      <footer class="console-panel">
        <div class="console-header">
          <span>运行日志</span>
          <button class="clear-console" @click="clearLogs">清空</button>
        </div>
        <div class="console-content" ref="consoleOutput">
          <div v-for="(log, index) in logs" :key="index" class="log-entry" :class="`log-${log.type}`">
            [{{ log.time }}] {{ log.message }}
          </div>
        </div>
      </footer>
    </main>
  </div>
</template>

<script setup>
import { ref, computed, nextTick, watch } from 'vue';
import BankReconciliation from './components/BankReconciliation.vue';

const currentTab = ref('bank-reconcile');
const logs = ref([{ time: new Date().toLocaleTimeString(), message: '[System] 准备就绪 (Vue 3)', type: 'info' }]);
const consoleOutput = ref(null);

const currentTitle = computed(() => {
  if (currentTab.value === 'bank-reconcile') return '集团银行明细对账';
  return '未命名功能';
});

const addLog = (message, type = 'info') => {
  logs.value.push({
    time: new Date().toLocaleTimeString(),
    message,
    type
  });
};

const clearLogs = () => {
  logs.value = [];
};

// Auto scroll to bottom
watch(logs.value, async () => {
  await nextTick();
  if (consoleOutput.value) {
    consoleOutput.value.scrollTop = consoleOutput.value.scrollHeight;
  }
});
</script>

<style>
/* Global Layout Styles (Migrated from index.css) */
:root {
  /* Modern Slate Palette */
  --bg-app: #f8fafc;
  /* Slate 50 - Lighter background */
  --bg-sidebar: #ffffff;
  /* White Sidebar for cleaner look */
  --text-sidebar: #334155;
  /* Slate 700 */
  --bg-card: #ffffff;
  --primary: #3b82f6;
  /* Blue 500 - Softer Blue */
  --primary-hover: #2563eb;
  /* Blue 600 */
  --secondary: #e2e8f0;
  /* Slate 200 */
  --border: #cbd5e1;
  /* Slate 300 */
  --text-main: #1e293b;
  /* Slate 800 */
  --text-muted: #64748b;
  /* Slate 500 */
  --console-bg: #2a3953;
  /* Slate 800 */
  --console-text: #f1f5f9;
  /* Slate 100 */
  --success: #10b981;
  /* Emerald 500 */
  --error: #ef4444;
  /* Red 500 */
}

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
  height: 100vh;
  overflow: hidden;
  color: var(--text-main);
  background-color: var(--bg-app);
  -webkit-font-smoothing: antialiased;
}

.app-container {
  display: flex;
  height: 100vh;
  width: 100vw;
}

/* Sidebar */
.sidebar {
  width: 280px;
  /* Increased width */
  background-color: var(--bg-sidebar);
  color: var(--text-sidebar);
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
  border-right: 1px solid var(--secondary);
  /* Add border instead of shadow for cleaner look */
  z-index: 10;
}

.brand {
  height: 70px;
  /* Increased height */
  padding: 0 28px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  border-bottom: 1px solid var(--secondary);
  background: transparent;
}

.brand h1 {
  font-size: 20px;
  font-weight: 700;
  letter-spacing: -0.5px;
  color: var(--primary);
  /* Brand color */
}

.brand span {
  font-size: 12px;
  opacity: 0.8;
  color: var(--text-muted);
  text-transform: uppercase;
  letter-spacing: 1.5px;
  font-weight: 500;
  margin-top: 2px;
}

.nav-menu {
  flex: 1;
  padding: 24px 16px;
  overflow-y: auto;
}

.nav-item {
  padding: 14px 18px;
  /* Larger padding */
  margin-bottom: 8px;
  border-radius: 8px;
  cursor: pointer;
  display: flex;
  align-items: center;
  transition: all 0.2s ease;
  font-size: 15px;
  font-weight: 500;
  color: var(--text-sidebar);
}

.nav-item:hover {
  background-color: #f1f5f9;
  /* Slate 100 */
  color: var(--primary);
}

.nav-item.active {
  background-color: #eff6ff;
  /* Blue 50 */
  color: var(--primary);
  box-shadow: none;
}

.nav-item.disabled {
  opacity: 0.4;
  cursor: not-allowed;
}

.nav-item .icon {
  margin-right: 12px;
  font-size: 16px;
}

/* Main Content */
.main-content {
  flex: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  position: relative;
}

.top-bar {
  height: 60px;
  background: white;
  padding: 0 32px;
  display: flex;
  align-items: center;
  border-bottom: 1px solid var(--border);
}

.top-bar h2 {
  font-size: 18px;
  font-weight: 600;
  color: var(--text-main);
}

.content-body {
  flex: 1;
  padding: 18px;
  overflow-y: auto;
}

/* Cards */
.card {
  background: var(--bg-card);
  border-radius: 12px;
  box-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1);
  border: 1px solid var(--border);
  margin-bottom: 24px;
  max-width: 900px;
}

.card-header {
  padding: 20px 24px;
  border-bottom: 1px solid var(--border);
  background-color: transparent;
}

.card-header h3 {
  font-size: 16px;
  font-weight: 600;
  color: var(--text-main);
}

.card-body {
  padding: 24px;
}

/* Forms */
.form-group {
  margin-bottom: 24px;
}

.form-group label {
  display: block;
  margin-bottom: 8px;
  font-size: 14px;
  font-weight: 600;
  color: var(--text-main);
}

.input-group {
  display: flex;
  gap: 12px;
}

input[type="text"] {
  flex: 1;
  padding: 10px 14px;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 14px;
  outline: none;
  background-color: #fff;
  transition: border-color 0.2s, box-shadow 0.2s;
  color: var(--text-main);
}

input[type="text"]:focus {
  border-color: var(--primary);
  box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
}

/* Buttons */
.btn {
  padding: 10px 20px;
  border-radius: 6px;
  border: none;
  cursor: pointer;
  font-size: 14px;
  font-weight: 600;
  transition: all 0.2s;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.btn-primary {
  background-color: var(--primary);
  color: white;
  box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
}

.btn-primary:hover {
  background-color: var(--primary-hover);
  transform: translateY(-1px);
  box-shadow: 0 4px 6px rgba(37, 99, 235, 0.25);
}

.btn-primary:active {
  transform: translateY(0);
}

.btn-primary:disabled {
  background-color: var(--secondary);
  box-shadow: none;
  cursor: not-allowed;
  transform: none;
}

.btn-secondary {
  background-color: white;
  color: var(--text-main);
  border: 1px solid var(--border);
}

.btn-secondary:hover {
  background-color: #f8fafc;
  border-color: #cbd5e1;
}

.help-text {
  margin-top: 8px;
  font-size: 12px;
  color: var(--text-muted);
}

/* Console Panel */
.console-panel {
  height: 200px;
  background-color: var(--console-bg);
  color: var(--console-text);
  display: flex;
  flex-direction: column;
  flex-shrink: 0;
  border-top: 1px solid rgba(255, 255, 255, 0.1);
}

.console-header {
  padding: 0 16px;
  height: 36px;
  background-color: rgba(0, 0, 0, 0.3);
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  color: #94a3b8;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.clear-console {
  background: none;
  border: none;
  color: #64748b;
  cursor: pointer;
  font-size: 11px;
  padding: 4px 8px;
  border-radius: 4px;
  transition: all 0.2s;
}

.clear-console:hover {
  color: #f1f5f9;
  background-color: rgba(255, 255, 255, 0.1);
}

.console-content {
  flex: 1;
  padding: 12px 16px;
  overflow-y: auto;
  font-family: 'JetBrains Mono', 'Menlo', 'Monaco', 'Courier New', monospace;
  font-size: 12px;
  line-height: 1.6;
}

.log-entry {
  margin-bottom: 4px;
  padding-bottom: 2px;
  border-bottom: 1px dashed rgba(255, 255, 255, 0.05);
}

.log-info {
  color: #e2e8f0;
}

.log-success {
  color: var(--success);
}

.log-error {
  color: var(--error);
}
</style>
