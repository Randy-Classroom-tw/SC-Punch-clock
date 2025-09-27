// App 全域總部
window.App = window.App || {};

// UI 模組 IIFE
(function(App) {
  'use strict';

  // SC辦公室出勤系統 UI模組 v2.0

  // 系統配置策略

  // 自包含配置策略，避免跨檔案依賴

  // 函式、工具、模組宣告

  // 初始化全域視窗追蹤物件
  if (!window.openedUploadWindows) {
    window.openedUploadWindows = {};
  }
  
  // 時區處理工具
  const TimeZoneUtils = {
    SYSTEM_TIMEZONE: 'Asia/Taipei',
    
    getTaiwanNow() {
      return new Date().toLocaleString('zh-TW', {
        timeZone: this.SYSTEM_TIMEZONE,
        hour12: false
      });
    },
    
    formatSystemTime(date = new Date()) {
      const options = {
        timeZone: this.SYSTEM_TIMEZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
      };
      
      const formatter = new Intl.DateTimeFormat('zh-TW', options);
      const parts = formatter.formatToParts(date);
      const values = {};
      
      parts.forEach(p => {
        if (p.type !== 'literal') values[p.type] = p.value;
      });
      
      return {
        date: `${values.year}-${values.month}-${values.day}`,
        time: `${values.hour}:${values.minute}`,
        full: `${values.year}-${values.month}-${values.day} ${values.hour}:${values.minute}`
      };
    }
  };

  /**
   * 智慧型 HTML 過濾器 (V4.4 新增)
   * 兼顧安全與體驗，只允許安全的 HTML 標籤與屬性
   */
  const DOMPurifyUtils = {
    /**
     * 核心過濾函數，並安全地設置 HTML
     * @param {HTMLElement} element - 目標元素
     * @param {string} dirtyHTML - 未經處理的 HTML 字串
     */
    safeSetHTML: function(element, dirtyHTML) {
      if (!element) return;
      if (typeof dirtyHTML !== 'string') {
        element.innerHTML = '';
        return;
      }

      // 使用瀏覽器內建的解析器來創建 DOM 樹，這比正則表達式更安全
      const parser = new DOMParser();
      const doc = parser.parseFromString(dirtyHTML, 'text/html');

      // 遞歸清理節點
      this._cleanNode(doc.body);

      // 將清理後的內容設置回去
      element.innerHTML = doc.body.innerHTML;
    },

    /**
     * 遞歸清理節點的輔助函數
     * @param {Node} node - 當前要清理的節點
     */
    _cleanNode: function(node) {
      // 安全標籤和屬性的白名單
      const ALLOWED_TAGS = new Set(['div', 'span', 'b', 'strong', 'i', 'em', 'br', 'p', 'h3', 'a', 'u', 'table', 'thead', 'tbody', 'tr', 'th', 'td']);
      const ALLOWED_ATTR = new Set(['style', 'class', 'id', 'href', 'target', 'colspan', 'rowspan']);

      // 遍歷所有子節點 (從後往前，方便刪除)
      for (let i = node.childNodes.length - 1; i >= 0; i--) {
        const child = node.childNodes[i];

        // 如果是元素節點
        if (child.nodeType === 1) {
          const tagName = child.tagName.toLowerCase();

          if (!ALLOWED_TAGS.has(tagName)) {
            while (child.firstChild) {
              node.insertBefore(child.firstChild, child);
            }
            node.removeChild(child);
            continue;
          }

          // 清理屬性
          for (let j = child.attributes.length - 1; j >= 0; j--) {
            const attr = child.attributes[j];
            const attrName = attr.name.toLowerCase();

            if (!ALLOWED_ATTR.has(attrName) || 
                (attrName === 'href' && (attr.value.startsWith('javascript:') || attr.value.startsWith('vbscript:'))) ||
                (attrName === 'style' && /expression|javascript|vbscript/i.test(attr.value))) {
              child.removeAttributeNode(attr);
            }
          }
          
          this._cleanNode(child);
        }
        else if (child.nodeType !== 3 && child.nodeType !== 8) { // 只保留文字節點和註解
           node.removeChild(child);
        }
      }
    }
  };

  /**
   * 安全 DOM 操作工具模組 - 防止 XSS 攻擊
   */
  const DOMUtils = {
    /**
     * 安全地設置元素文本內容
     * @param {HTMLElement} element - 目標元素
     * @param {string} text - 要設置的文本
     */
    safeSetText: function(element, text) {
      if (element && typeof text === 'string') {
        element.textContent = text; // 永遠使用 textContent，避免 XSS
      }
    },
    
    /**
     * 安全地設置 HTML 內容（僅在確實需要時使用）
     * @param {HTMLElement} element - 目標元素
     * @param {string} htmlString - HTML 字串
     */
    safeSetHTML: function(element, htmlString) {
      if (element && typeof htmlString === 'string') {
        const sanitized = this.sanitizeHTML(htmlString);
        element.innerHTML = sanitized;
      }
    },
    
    /**
     * 簡單的 HTML 清理函數
     * @param {string} html - 要清理的 HTML
     * @returns {string} 清理後的 HTML
     */
    sanitizeHTML: function(html) {
      if (typeof html !== 'string') return '';
      // 移除 script 標籤和其他危險標籤
      return html
        .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
        .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '')
        .replace(/<object\b[^<]*(?:(?!<\/object>)<[^<]*)*<\/object>/gi, '')
        .replace(/<embed\b[^<]*(?:(?!<\/embed>)<[^<]*)*<\/embed>/gi, '')
        .replace(/javascript:/gi, '')
        .replace(/on\w+\s*=/gi, '');
    },
    
    /**
     * 創建安全的文本節點
     * @param {string} text - 文本內容
     * @returns {Text} 文本節點
     */
    createTextNode: function(text) {
      return document.createTextNode(text || '');
    }
  };

  /**
   * 防抖動函數
   */
  function debounce(func, wait) {
    let timeout;
    return function (...args) {
      clearTimeout(timeout);
      timeout = setTimeout(() => func.apply(this, args), wait);
    };
  }

  /**
   * 檢查 4 位數字
   */
  function validateIdCode(idCode) {
    return /^\d{4}$/.test(idCode);
  }

  /**
   * 取消處理函式工廠 (P2: JSONP中斷機制)
   */
  function setupCancellationHandler(controller, modalSelector) {
    const modal = document.querySelector(modalSelector);
    if (!modal) return null;
    
    const abort = () => controller.abort();
    const keydownHandler = (e) => {
      if (e.key === 'Escape') abort();
    };
    const clickHandler = (e) => {
      if (e.target === modal || e.target.closest('.btn-cancel')) {
        abort();
      }
    };
    
    document.addEventListener('keydown', keydownHandler);
    modal.addEventListener('click', clickHandler);
    
    return () => {
      document.removeEventListener('keydown', keydownHandler);
      modal.removeEventListener('click', clickHandler);
    };
  }

  /**
   * 按鈕狀態管理器
   */
  const ButtonStateManager = {
    // 追蹤正在冷卻中的按鈕
    coolingButtons: new Set(),
    
    // 追蹤操作狀態
    operationInProgress: false,
    
    // 將按鈕添加到冷卻集合
    addToCooling: function(buttonId) {
      this.coolingButtons.add(buttonId);
    },
    
    // 從冷卻集合中移除按鈕
    removeFromCooling: function(buttonId) {
      this.coolingButtons.delete(buttonId);
    },
    
    // 檢查按鈕是否正在冷卻
    isButtonCooling: function(buttonId) {
      return this.coolingButtons.has(buttonId);
    },
    
    // 設置操作狀態
    setOperationStatus: function(inProgress) {
      this.operationInProgress = inProgress;
    },
    
    // 重置所有狀態（緊急情況使用）
    resetAll: function() {
      this.coolingButtons.clear();
      this.operationInProgress = false;
    }
  };

  /**
   * 瀏覽器相關工具
   */
  const BrowserUtils = {
    isSafari: () => {
      const ua = navigator.userAgent.toLowerCase();
      return /safari/.test(ua) && 
             !/chrome/.test(ua) && 
             !/crios/.test(ua) &&    // 排除 Chrome iOS
             !/firefox/.test(ua) && 
             !/fxios/.test(ua) &&    // 排除 Firefox iOS
             !/opr/.test(ua) && 
             !/edg/.test(ua);
    },
    showSafariNotice: () => {
      if (BrowserUtils.isSafari() && elements.safariNotice) {
        elements.safariNotice.classList.add('show');
      }
    },
    parseBrowser: (userAgent) => {
      const ua = userAgent.toLowerCase();
      // 優先檢查特定瀏覽器，避免誤判
      if (/edg\//.test(ua)) return "Edge";
      if (/opr\/|opera\//.test(ua)) return "Opera";
      
      // Chrome 檢測（包含 iOS 版本 CriOS）
      if (/chrome\//.test(ua) || /crios\//.test(ua)) return "Chrome";
      
      // Firefox 檢測（包含 iOS 版本 FxiOS）
      if (/firefox\//.test(ua) || /fxios\//.test(ua)) return "Firefox";
      
      // Safari 檢測（必須排除其他瀏覽器）
      if (/safari\//.test(ua) && 
          !/chrome\//.test(ua) && 
          !/crios\//.test(ua) && 
          !/firefox\//.test(ua) && 
          !/fxios\//.test(ua) && 
          !/opr\//.test(ua) && 
          !/edg\//.test(ua)) {
        return "Safari";
      }
      
      if (/trident\//.test(ua)) return "Internet Explorer";
      return "未知瀏覽器";
    },
    /**
     * 智慧型彈出視窗權限檢查 (v2.0)
     * @returns {Promise<boolean>} - 返回一個解析為布林值的 Promise，true 代表允許，false 代表被阻擋。
     */
    checkPopupPermission: async () => {
      return new Promise(resolve => {
        setTimeout(() => {
          let testPopup;
          const browserType = BrowserUtils.parseBrowser(navigator.userAgent);
          
          try {
            testPopup = window.open('', '_blank', 'width=1,height=1,left=9999,top=9999');
            
            if (testPopup) {
              // 對於非 Safari 瀏覽器，如果能打開視窗但無法立即訪問，也視為被部分阻擋
              if (browserType !== 'Safari' && (testPopup.closed === false && !testPopup.window)) {
                 testPopup.close(); // 嘗試關閉
                 resolve(false);
                 return;
              }
              // 如果能成功打開並關閉，權限正常
              testPopup.close();
              resolve(true);
            } else {
              // window.open 返回 null 或 undefined，明確表示被阻擋
              resolve(false);
            }
          } catch (e) {
            // 任何錯誤都視為被阻擋
            console.error("彈出視窗權限檢查出錯:", e);
            resolve(false);
          }
        }, 100);
      });
    },
    getWebGLFingerprint: () => {
      try {
        const canvas = document.createElement('canvas');
        const gl = canvas.getContext('webgl');
        if (!gl) return 'WebGL not supported';
        const debugInfo = gl.getExtension('WEBGL_debug_renderer_info');
        if (debugInfo) {
          const vendor = gl.getParameter(debugInfo.UNMASKED_VENDOR_WEBGL) || 'unknown_vendor';
          const renderer = gl.getParameter(debugInfo.UNMASKED_RENDERER_WEBGL) || 'unknown_renderer';
          return `${vendor}-${renderer}`;
        }
        return 'No debug info';
      } catch (err) {
        console.error("WebGL 指紋失敗:", err);
        return 'WebGL error';
      }
    },
    detectCommonFonts: () => {
      const fontList = [
        'Arial','Times New Roman','Courier New','Georgia',
        'Verdana','Helvetica','Tahoma','Trebuchet MS'
      ];
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (!context) return "canvas_unsupported";
      const testString = "abcdefghijklmnopqrstuvwxyz123456789";
      context.font = "72px monospace";
      const baselineSize = context.measureText(testString).width;
      const detectedFonts = fontList.filter(font => {
        context.font = `72px '${font}', monospace`;
        return context.measureText(testString).width !== baselineSize;
      });
      return detectedFonts.length > 0 ? detectedFonts.join('|').substring(0, 100) : "no_fonts_detected";
    }
  };

  /**
   * 儲存相關工具
   */
  const StorageUtils = {
    getCookie: (name) => {
      const nameEQ = name + "=";
      const ca = document.cookie.split(';');
      for (let i=0; i<ca.length; i++) {
        let c = ca[i].trim();
        if (c.indexOf(nameEQ) === 0) return c.substring(nameEQ.length);
      }
      return null;
    },
    setCookie: (name, value, days) => {
      try {
        const date = new Date();
        date.setTime(date.getTime() + (days*24*60*60*1000));
        const expires = "expires=" + date.toUTCString();
        document.cookie = `${name}=${value};${expires};path=/;SameSite=Strict`;
        return true;
      } catch(e) {
        console.error("設置 Cookie 失敗:", e);
        return false;
      }
    }
  };

  /**
   * 加密與識別碼工具
   */
  const CryptoUtils = {
    generateFallbackID: () => {
      if (crypto && typeof crypto.randomUUID === 'function') {
        return crypto.randomUUID();
      }
      return 'xxxxxxxxxxxx4xxxyxxxxxxxxxxxxxxx'.replace(/[xy]/g, c => {
        const r = Math.random()*16|0;
        const v = (c === 'x') ? r : (r & 0x3 | 0x8);
        return v.toString(16);
      });
    },
    hashString: async (str) => {
      try {
        if (!str) throw new Error("無法雜湊空字串");
        const encoder = new TextEncoder();
        const data = encoder.encode(str);
        const hashBuffer = await crypto.subtle.digest("SHA-256", data);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        return hashArray.map(b => b.toString(16).padStart(2,"0")).join("");
      } catch(err) {
        console.error("雜湊函數失敗:", err);
        return CryptoUtils.generateFallbackID();
      }
    }
  };

  /**
   * 獲取 Web App URL，包含防禦性檢查
   * @returns {string} Web App URL 或空字串
   */
  function getWebAppUrl() {
    // 防禦性檢查：確保 App.config 已初始化
    if (!window.App || !App.config || !App.config.WEB_APP_URL) {
      console.warn('[getWebAppUrl] App.config 尚未初始化或 WEB_APP_URL 未設定');
      return '';
    }
    return App.config.WEB_APP_URL;
  }

  /**
   * 網路相關工具 (JSONP)
   */
  const NetworkUtils = {
    getIPAddress: async () => {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 3000);
        const response = await fetch('https://api.ipify.org?format=json', {signal: controller.signal});
        clearTimeout(timeoutId);
        if (!response.ok) throw new Error(`網路錯誤: ${response.status}`);
        const data = await response.json();
        return data.ip;
      } catch(err) {
        console.error("無法獲取 IP:", err);
        return 'unknown';
      }
    },
    // 網路品質監控
    networkMetrics: {
      responseTimes: [],
      successCount: 0,
      totalCount: 0,
      lastCheck: Date.now()
    },
    
    // 記錄API回應時間
    recordNetworkMetrics: (responseTime, success = true) => {
      const metrics = NetworkUtils.networkMetrics;
      metrics.responseTimes.push(responseTime);
      metrics.totalCount++;
      if (success) metrics.successCount++;
      
      // 只保留最近10次記錄
      if (metrics.responseTimes.length > 10) {
        metrics.responseTimes.shift();
      }
      
      // 更新網路狀態顯示
      NetworkUtils.updateNetworkStatus();
    },
    
    // 評估網路品質
    assessNetworkQuality: () => {
      const metrics = NetworkUtils.networkMetrics;
      
      if (!navigator.onLine) return 'offline';
      if (metrics.responseTimes.length === 0) return 'good'; // 預設狀態
      
      const avgResponseTime = metrics.responseTimes.reduce((a, b) => a + b, 0) / metrics.responseTimes.length;
      const successRate = metrics.totalCount > 0 ? metrics.successCount / metrics.totalCount : 1;
      
      // 簡化為4種狀態
      if (avgResponseTime < 2000 && successRate > 0.90) return 'good';
      if (avgResponseTime < 5000 && successRate > 0.80) return 'medium';
      return 'poor';
    },
    
    updateNetworkStatus: () => {
      const quality = NetworkUtils.assessNetworkQuality();
      
      const statusConfigs = {
        good: { icon: 'wifi', text: '良好', class: 'network-good' },  // 使用 wifi 滿格圖標
        medium: { icon: 'network_wifi_3_bar', text: '一般', class: 'network-medium' },  // 改用3格，與1格有更明顯區別
        poor: { icon: 'network_wifi_1_bar', text: '緩慢', class: 'network-poor' },
        offline: { icon: 'wifi_off', text: '離線', class: 'network-offline' }
      };
      
      const config = statusConfigs[quality];
      elements.networkStatus.innerHTML = `
        <span class="material-icons">${config.icon}</span>
        <span class="status-text">${config.text}</span>
      `;
      elements.networkStatus.className = `network-status ${config.class}`;
      
      if (quality === 'offline') {
        // 移除自動彈出的網路錯誤對話框，因為右上角已有離線狀態顯示
        // UIUtils.displayError("網路連線已中斷");
      }
    },
    callGAS: async (functionName, parametersArray, signal = null) => {
      const startTime = Date.now(); // 記錄開始時間
      return new Promise((resolve, reject) => {
        let script, timeout, isAborted = false;
        let callbackName; // 提前宣告
        
        // AbortSignal 處理
        const cleanup = () => {
          if (script && script.parentNode) script.parentNode.removeChild(script);
          if (timeout) clearTimeout(timeout);
          if (callbackName) delete window[callbackName];
        };
        
        const abortHandler = () => {
          isAborted = true;
          cleanup();
          reject(new Error('AbortError'));
        };
        
        if (signal) {
          if (signal.aborted) return reject(new Error('AbortError'));
          signal.addEventListener('abort', abortHandler);
        }
        
        try {
          callbackName = 'jsonpCallback_' + Date.now() + '_' + Math.random().toString(36).substring(2,10);
          window[callbackName] = (data) => {
            if (isAborted) return;
            const responseTime = Date.now() - startTime;
            console.log(`JSONP回調成功: ${functionName}`, data.status||'未知狀態');
            
            // 記錄網路指標
            NetworkUtils.recordNetworkMetrics(responseTime, true);
            
            cleanup();
            if (signal) signal.removeEventListener('abort', abortHandler);
            resolve(data);
          };
          
          const params = new URLSearchParams({
            function: functionName,
            parameters: JSON.stringify(parametersArray),
            callback: callbackName
          });
          const fullUrl = `${getWebAppUrl()}?${params.toString()}`;
          console.log(`發起JSONP請求: ${functionName}`, parametersArray);
          
          script = document.createElement('script');
          script.src = fullUrl;
          script.async = true;
          
          script.onerror = (e) => {
            if (isAborted) return;
            const responseTime = Date.now() - startTime;
            console.error(`JSONP請求失敗: ${functionName}`, e);
            
            // 記錄失敗的網路指標
            NetworkUtils.recordNetworkMetrics(responseTime, false);
            
            cleanup();
            if (signal) signal.removeEventListener('abort', abortHandler);
            reject(new Error('JSONP 請求失敗 - 請檢查網路連接'));
          };
          
          timeout = setTimeout(() => {
            if (isAborted) return;
            const responseTime = Date.now() - startTime;
            console.warn(`JSONP請求超時: ${functionName}`);
            
            // 記錄超時的網路指標
            NetworkUtils.recordNetworkMetrics(responseTime, false);
            
            cleanup();
            if (signal) signal.removeEventListener('abort', abortHandler);
            reject(new Error('請求超時 - 伺服器回應過慢'));
          }, App.config.API_TIMEOUT);
          
          script.onload = () => clearTimeout(timeout);
          document.head.appendChild(script);
        } catch(err) {
          if (signal) signal.removeEventListener('abort', abortHandler);
          console.error("構建JSONP請求失敗:", err);
          reject(new Error(`請求建立失敗: ${err.message}`));
        }
      });
    },
    callGASWithRetry: async (functionName, parametersArray, retries, signal = null) => {
      const MAX_RETRIES = App.config.JSONP_RETRY_COUNT || 3;
      const INITIAL_DELAY = 1000; // 初始延遲1秒
      // 內部遞歸函式，用於處理重試邏輯
      const attempt = async (retryCount) => {
        try {
          // 首次或重試時呼叫核心API
          return await NetworkUtils.callGAS(functionName, parametersArray, signal);
        } catch (error) {
          // 如果是使用者主動取消，則直接拋出錯誤，不進行任何重試
          if (error.message === 'AbortError') {
            throw error;
          }
          // 如果重試次數已用盡，拋出最終的、更友善的錯誤
          if (retryCount >= MAX_RETRIES) {
            console.error(`網路請求 '${functionName}' 在 ${MAX_RETRIES} 次重試後最終失敗。`, error);
            throw new Error("網路連線不穩定，請檢查您的網路並稍後再試。");
          }
          // 核心升級：實施「漸進式退避」演算法
          // 計算下一次重試的延遲時間，每次重試延遲都會加倍，並加入少量隨機性以避免同時重試
          const delay = (INITIAL_DELAY * Math.pow(2, retryCount)) + (Math.random() * 1000);
          
          console.log(`請求 '${functionName}' 失敗 (第 ${retryCount + 1} 次)，將在 ${(delay / 1000).toFixed(1)} 秒後重試...`);
          
          // 在UI上提供更智慧的提示
          App.UI.UIUtils.displayLoading(`網路連線不穩... 第 ${retryCount + 1} 次重試中...`);
          
          // 等待計算出的延遲時間
          await new Promise(resolve => setTimeout(resolve, delay));
          
          // 遞歸呼叫下一次嘗試
          return attempt(retryCount + 1);
        }
      };
      // 啟動第一次嘗試
      return attempt(0);
    },
    checkJSONPError: (response) => {
      return !response || (response.error && response.status === 'error');
    }
  };

  /**
   * UI 介面工具 (v2.0 安全性強化版)
   */
  const UIUtils = {
    /**
     * 安全地顯示錯誤訊息
     * @param {string} msg - 要顯示的訊息
     */
    /**
     * 錯誤分類配置系統
     * 定義不同類型錯誤的處理方式和用戶指引
     */
    ERROR_CATEGORIES: {
      NETWORK: {
        icon: 'wifi_off',
        color: '#ff5722',
        backgroundColor: 'rgba(255, 87, 34, 0.1)',
        borderColor: '#ff5722',
        title: '網路連線問題',
        suggestions: [
          '請檢查網路連線狀態',
          '嘗試重新整理頁面',
          '確認網路設定是否正確'
        ]
      },
      VALIDATION: {
        icon: 'error_outline',
        color: '#ff9800',
        backgroundColor: 'rgba(255, 152, 0, 0.1)',
        borderColor: '#ff9800',
        title: '資料驗證錯誤',
        suggestions: [
          '請檢查必填欄位是否完整',
          '確認資料格式是否正確',
          '檢查日期範圍是否合理'
        ]
      },
      PERMISSION: {
        icon: 'block',
        color: '#e91e63',
        backgroundColor: 'rgba(233, 30, 99, 0.1)',
        borderColor: '#e91e63',
        title: '權限或授權問題',
        suggestions: [
          '請確認您有執行此操作的權限',
          '檢查瀏覽器彈出視窗設定',
          '請點擊 AI Bot 尋求協助'
        ]
      },
      SYSTEM: {
        icon: 'warning',
        color: '#f44336',
        backgroundColor: 'rgba(244, 67, 54, 0.1)',
        borderColor: '#f44336',
        title: '系統錯誤',
        suggestions: [
          '請稍後重新嘗試',
          '如問題持續，請點擊 AI Bot 尋求協助',
          '請記錄錯誤發生的具體操作步驟'
        ]
      },
      GENERAL: {
        icon: 'info',
        color: '#f44336',
        backgroundColor: 'rgba(244, 67, 54, 0.1)',
        borderColor: '#f44336',
        title: '操作失敗',
        suggestions: []
      }
    },

    /**
     * 智慧錯誤分類器
     * 根據錯誤訊息內容自動判斷錯誤類型
     * @param {string} message - 錯誤訊息
     * @returns {string} 錯誤類型
     */
    classifyError: (message) => {
      const msg = message.toLowerCase();
      
      // 網路相關錯誤
      if (msg.includes('網路') || msg.includes('連線') || msg.includes('timeout') || 
          msg.includes('offline') || msg.includes('請求失敗') || msg.includes('無法連接')) {
        return 'NETWORK';
      }
      
      // 驗證相關錯誤
      if (msg.includes('驗證') || msg.includes('必填') || msg.includes('格式') || 
          msg.includes('請輸入') || msg.includes('請選擇') || msg.includes('invalid')) {
        return 'VALIDATION';
      }
      
      // 權限相關錯誤
      if (msg.includes('權限') || msg.includes('授權') || msg.includes('阻擋') || 
          msg.includes('彈出視窗') || msg.includes('不在授權名單') || msg.includes('blocked')) {
        return 'PERMISSION';
      }
      
      // 系統相關錯誤
      if (msg.includes('系統') || msg.includes('伺服器') || msg.includes('server') || 
          msg.includes('500') || msg.includes('error') || msg.includes('exception')) {
        return 'SYSTEM';
      }
      
      return 'GENERAL';
    },

    /**
     * 增強版錯誤顯示功能
     * 支援錯誤分類、圖示、建議和重試機制
     * @param {string} msg - 錯誤訊息
     * @param {Object} options - 選項
     * @param {string} options.type - 錯誤類型（可選，會自動分類）
     * @param {Function} options.retryCallback - 重試回調函數（可選）
     * @param {boolean} options.showSuggestions - 是否顯示建議（預設true）
     */
    displayEnhancedError: (msg, options = {}) => {
      const errorType = options.type || UIUtils.classifyError(msg);
      const category = UIUtils.ERROR_CATEGORIES[errorType];
      const showSuggestions = options.showSuggestions !== false;
      
      // 創建錯誤容器
      const errorContainer = document.createElement('div');
      errorContainer.className = 'enhanced-error-container';
      errorContainer.style.cssText = `
        background-color: ${category.backgroundColor};
        border: 1px solid ${category.borderColor};
        border-radius: 8px;
        padding: 20px;
        margin: 10px 0;
        text-align: left;
        max-width: 500px;
        margin-left: auto;
        margin-right: auto;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      `;
      
      // 錯誤標題區域
      const titleSection = document.createElement('div');
      titleSection.style.cssText = `
        display: flex;
        align-items: center;
        margin-bottom: 15px;
        color: ${category.color};
        font-weight: bold;
        font-size: 16px;
      `;
      
      const icon = document.createElement('span');
      icon.className = 'material-icons';
      icon.textContent = category.icon;
      icon.style.cssText = `
        font-size: 24px;
        margin-right: 10px;
        color: ${category.color};
      `;
      
      const title = document.createElement('span');
      title.textContent = category.title;
      
      titleSection.appendChild(icon);
      titleSection.appendChild(title);
      
      // 錯誤訊息區域
      const messageSection = document.createElement('div');
      messageSection.style.cssText = `
        margin-bottom: 15px;
        color: #333;
        line-height: 1.5;
        font-size: 14px;
      `;
      DOMPurifyUtils.safeSetHTML(messageSection, msg);
      
      // 建議區域
      let suggestionsSection = null;
      if (showSuggestions && category.suggestions.length > 0) {
        suggestionsSection = document.createElement('div');
        suggestionsSection.style.cssText = `
          margin-bottom: 15px;
          padding: 10px;
          background-color: rgba(255,255,255,0.5);
          border-radius: 4px;
        `;
        
        const suggestionsTitle = document.createElement('div');
        suggestionsTitle.textContent = '建議解決方案：';
        suggestionsTitle.style.cssText = `
          font-weight: bold;
          margin-bottom: 8px;
          color: ${category.color};
          font-size: 13px;
        `;
        
        const suggestionsList = document.createElement('ul');
        suggestionsList.style.cssText = `
          margin: 0;
          padding-left: 20px;
          color: #555;
          font-size: 13px;
        `;
        
        category.suggestions.forEach(suggestion => {
          const listItem = document.createElement('li');
          listItem.textContent = suggestion;
          listItem.style.marginBottom = '4px';
          suggestionsList.appendChild(listItem);
        });
        
        suggestionsSection.appendChild(suggestionsTitle);
        suggestionsSection.appendChild(suggestionsList);
      }
      
      // 操作按鈕區域
      const actionsSection = document.createElement('div');
      actionsSection.style.cssText = `
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 10px;
      `;
      
      // 關閉按鈕
      const closeButton = document.createElement('button');
      closeButton.textContent = '關閉';
      closeButton.style.cssText = `
        background-color: #666;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 13px;
      `;
      closeButton.onclick = () => {
        elements.responseDiv.innerHTML = '';
        elements.responseDiv.classList.remove("show");
      };
      
      // 重試按鈕（如果提供重試回調）
      let retryButton = null;
      if (options.retryCallback && typeof options.retryCallback === 'function') {
        retryButton = document.createElement('button');
        retryButton.textContent = '重新嘗試';
        retryButton.style.cssText = `
          background-color: ${category.color};
          color: white;
          border: none;
          padding: 8px 16px;
          border-radius: 4px;
          cursor: pointer;
          font-size: 13px;
          margin-left: 10px;
        `;
        retryButton.onclick = () => {
          options.retryCallback();
        };
      }
      
      actionsSection.appendChild(closeButton);
      if (retryButton) {
        actionsSection.appendChild(retryButton);
      }
      
      // 組合所有元素
      errorContainer.appendChild(titleSection);
      errorContainer.appendChild(messageSection);
      if (suggestionsSection) {
        errorContainer.appendChild(suggestionsSection);
      }
      errorContainer.appendChild(actionsSection);
      
      // 顯示錯誤
      elements.responseDiv.innerHTML = '';
      elements.responseDiv.appendChild(errorContainer);
      elements.responseDiv.classList.add("show");
      
      console.log(`增強錯誤顯示: ${errorType} - ${msg}`);
    },

    /**
     * 顯示錯誤訊息（保持向下兼容性）
     * @param {string} msg - 錯誤訊息
     * @param {Object} options - 選項（可選）
     */
    displayError: (msg, options = {}) => {
      // 如果提供了 options 或訊息包含特定關鍵詞，使用增強版顯示
      if (Object.keys(options).length > 0 || 
          msg.includes('網路') || msg.includes('權限') || msg.includes('驗證') || msg.includes('系統')) {
        UIUtils.displayEnhancedError(msg, options);
      } else {
        // 否則使用原有的簡單顯示方式
        const errorDiv = document.createElement('div');
        errorDiv.style.color = 'red';
        errorDiv.style.textAlign = 'center';
        DOMPurifyUtils.safeSetHTML(errorDiv, msg);
        elements.responseDiv.innerHTML = '';
        elements.responseDiv.appendChild(errorDiv);
        elements.responseDiv.classList.add("show");
      }
    },

    /**
     * 安全地顯示成功訊息
     * @param {string} msg - 要顯示的訊息
     */
    displaySuccess: (msg) => {
          // V4.4 修正：使用 innerHTML 並搭配過濾器
    const successDiv = document.createElement('div');
    successDiv.style.color = 'green';
    successDiv.style.fontWeight = 'bold';
    successDiv.style.textAlign = 'center';
    DOMPurifyUtils.safeSetHTML(successDiv, msg); // 使用新的安全設定HTML
      elements.responseDiv.innerHTML = ''; // 先清空
      elements.responseDiv.appendChild(successDiv);
      elements.responseDiv.classList.add("show");
    },

    /**
     * 安全地顯示載入中訊息
     * @param {string} msg - 要顯示的訊息
     */
    displayLoading: (msg) => {
      const loadingSpan = document.createElement('span');
      loadingSpan.className = 'loading';
      loadingSpan.textContent = msg || "載入中｜請稍候..."; // 使用 .textContent

      elements.responseDiv.innerHTML = ''; // 先清空
      elements.responseDiv.appendChild(loadingSpan);
      elements.responseDiv.classList.add("show");
    },

    /**
     * 安全地顯示一般訊息
     * @param {string} msg - 要顯示的訊息
     */
    displayMessage: (msg) => {
      const messageDiv = document.createElement('div');
      messageDiv.style.color = '#76ff03';
      messageDiv.style.textAlign = 'center';
      messageDiv.style.fontSize = '16px';
      messageDiv.textContent = msg || "操作完成";

      elements.responseDiv.innerHTML = '';
      elements.responseDiv.appendChild(messageDiv);
      elements.responseDiv.classList.add("show");
    },

    clearResponse: () => {
      elements.responseDiv.innerHTML = "";
      elements.responseDiv.classList.remove("show");
    },

    disableAllButtons: () => {
      ButtonStateManager.setOperationStatus(true);
      Object.values(elements.allButtons).forEach(button => {
        button.disabled = true;
        if (!ButtonStateManager.isButtonCooling(button.id)) {
          button.style.background = 'linear-gradient(to right, var(--btn-disabled-from), var(--btn-disabled-to))';
        }
      });
    },

    enableAllButtons: () => {
      ButtonStateManager.setOperationStatus(false);
      Object.values(elements.allButtons).forEach(button => {
        if (!ButtonStateManager.isButtonCooling(button.id)) {
          button.disabled = false;
          button.style.background = originalButtonStyles[button.id];
        }
      });
    },

    enableAllButtonsExcept: (buttonToExclude) => {
      ButtonStateManager.setOperationStatus(false);
      Object.values(elements.allButtons).forEach(button => {
        if (button.id !== buttonToExclude.id && !ButtonStateManager.isButtonCooling(button.id)) {
          button.disabled = false;
          button.style.background = originalButtonStyles[button.id];
        }
      });
    },

    startIndividualCooldown: (button, duration, originalText) => {
      button.disabled = true;
      button.textContent = originalText;
      const btnId = button.id;
      
      ButtonStateManager.addToCooling(btnId);
      
      const colorFrom = buttonColors[btnId]?.from || 'var(--btn-disabled-from)';
      const colorTo = buttonColors[btnId]?.to || 'var(--btn-disabled-to)';
      button.style.setProperty('--color-from', colorFrom);
      button.style.setProperty('--color-to', colorTo);
      button.style.setProperty('--progress', '0%');
      button.classList.add('btn-cooldown');
      button.style.background = `linear-gradient(to right, ${colorFrom} 0%, ${colorTo} 0%, var(--btn-disabled-from) 0%, var(--btn-disabled-to) 100%)`;
      
      const startTime = performance.now();
      function updateProgress(timestamp) {
        const elapsed = timestamp - startTime;
        const progress = Math.min((elapsed / duration) * 100, 100);
        button.style.setProperty('--progress', `${progress}%`);
        
        if (progress < 100) {
          requestAnimationFrame(updateProgress);
        } else {
          button.classList.remove('btn-cooldown');
          button.style.background = originalButtonStyles[btnId];
          button.disabled = ButtonStateManager.operationInProgress;
          ButtonStateManager.removeFromCooling(btnId);
        }
      }
      
      requestAnimationFrame(updateProgress);
    },

    showCustomConfirm: (title, message, onConfirm, onCancel) => {
      // 創建自定義確認對話框
      const customOverlay = document.createElement('div');
      customOverlay.className = 'custom-confirm-overlay';
      customOverlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0, 0, 0, 0.8);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 10000;
        opacity: 0;
        visibility: hidden;
        transition: all 0.3s ease;
        backdrop-filter: blur(5px);
      `;

      const customDialog = document.createElement('div');
      customDialog.className = 'custom-confirm-dialog';
      customDialog.style.cssText = `
        background: linear-gradient(145deg, #1a1a1a, #0f0f0f);
        border: 1px solid #333333;
        border-radius: 12px;
        padding: 24px 28px;
        max-width: 420px;
        width: 90%;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4), inset 0 1px 0 rgba(255, 255, 255, 0.1);
        transform: scale(0.9) translateY(20px);
        transition: all 0.3s ease;
        text-align: center;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      `;

      const titleContainer = document.createElement('div');
      titleContainer.style.cssText = `
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 12px;
        margin-bottom: 16px;
      `;

      const iconDiv = document.createElement('div');
      iconDiv.style.cssText = `
        width: 36px;
        height: 36px;
        background: linear-gradient(145deg, #2a2a2a, #1a1a1a);
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.1);
      `;
      iconDiv.innerHTML = '<span class="material-icons" style="color: #ff6b6b; font-size: 20px;">warning</span>';

      const titleDiv = document.createElement('h3');
      titleDiv.style.cssText = `
        font-size: 20px;
        font-weight: 500;
        color: #ffffff;
        margin: 0;
      `;
      titleDiv.textContent = title || "確認操作";

      titleContainer.appendChild(iconDiv);
      titleContainer.appendChild(titleDiv);

      const messageDiv = document.createElement('p');
      messageDiv.style.cssText = `
        color: #cccccc;
        font-size: 15px;
        line-height: 1.5;
        margin-bottom: 28px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      `;
      messageDiv.textContent = message || "確定要執行此操作嗎？";

      const buttonsDiv = document.createElement('div');
      buttonsDiv.style.cssText = `
        display: flex;
        gap: 16px;
        justify-content: center;
      `;

      const cancelBtn = document.createElement('button');
      cancelBtn.style.cssText = `
        padding: 12px 24px;
        border-radius: 8px;
        border: 1px solid #444444;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
        min-width: 80px;
        background: linear-gradient(145deg, #2a2a2a, #1a1a1a);
        color: #aaaaaa;
        order: 1;
      `;
      cancelBtn.textContent = '取消';
      cancelBtn.onmouseover = () => {
        cancelBtn.style.background = 'linear-gradient(145deg, #3a3a3a, #2a2a2a)';
        cancelBtn.style.color = '#ffffff';
        cancelBtn.style.transform = 'translateY(-1px)';
      };
      cancelBtn.onmouseout = () => {
        cancelBtn.style.background = 'linear-gradient(145deg, #2a2a2a, #1a1a1a)';
        cancelBtn.style.color = '#aaaaaa';
        cancelBtn.style.transform = 'translateY(0)';
      };

      const confirmBtn = document.createElement('button');
      confirmBtn.style.cssText = `
        padding: 12px 24px;
        border-radius: 8px;
        border: 1px solid #3b82f6;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        transition: all 0.2s ease;
        min-width: 80px;
        background: linear-gradient(145deg, #3b82f6, #1d4ed8);
        color: #ffffff;
        box-shadow: 0 0 15px rgba(59, 130, 246, 0.3);
        order: 2;
      `;
      confirmBtn.textContent = '確定';
      confirmBtn.onmouseover = () => {
        confirmBtn.style.background = 'linear-gradient(145deg, #60a5fa, #3b82f6)';
        confirmBtn.style.boxShadow = '0 0 20px rgba(59, 130, 246, 0.4)';
        confirmBtn.style.transform = 'translateY(-1px)';
      };
      confirmBtn.onmouseout = () => {
        confirmBtn.style.background = 'linear-gradient(145deg, #3b82f6, #1d4ed8)';
        confirmBtn.style.boxShadow = '0 0 15px rgba(59, 130, 246, 0.3)';
        confirmBtn.style.transform = 'translateY(0)';
      };

      // 事件處理
      const closeDialog = () => {
        customOverlay.style.opacity = '0';
        customOverlay.style.visibility = 'hidden';
        setTimeout(() => {
          if (customOverlay.parentNode) {
            customOverlay.parentNode.removeChild(customOverlay);
          }
        }, 300);
      };

      cancelBtn.onclick = () => {
        closeDialog();
        if (onCancel) onCancel();
      };

      confirmBtn.onclick = () => {
        closeDialog();
        if (onConfirm) onConfirm();
      };

      // 點擊背景關閉
      customOverlay.onclick = (e) => {
        if (e.target === customOverlay) {
          closeDialog();
          if (onCancel) onCancel();
        }
      };

      // 組裝對話框
      buttonsDiv.appendChild(cancelBtn);
      buttonsDiv.appendChild(confirmBtn);
      customDialog.appendChild(titleContainer);
      customDialog.appendChild(messageDiv);
      customDialog.appendChild(buttonsDiv);
      customOverlay.appendChild(customDialog);

      // 添加到頁面並顯示
      document.body.appendChild(customOverlay);
      
      // 觸發顯示動畫
      setTimeout(() => {
        customOverlay.style.opacity = '1';
        customOverlay.style.visibility = 'visible';
        customDialog.style.transform = 'scale(1) translateY(0)';
      }, 10);
    }
  };

  /**
   * 裝置識別碼管理
   */
  const DeviceIdManager = {
    saveDeviceId: (deviceId) => {
      if (!deviceId) {
        console.error("嘗試儲存空的設備識別碼");
        return false;
      }
      let successCount = 0;
      const timestamp = Date.now().toString();
      const appVersion = App.config.APP_VERSION || 'unknown';
      const deviceIdKey = App.config.DEVICE_ID_KEY;

      try {
        localStorage.setItem(deviceIdKey, deviceId);
        localStorage.setItem(`${deviceIdKey}_timestamp`, timestamp);
        localStorage.setItem(`${deviceIdKey}_version`, appVersion);
        successCount++;
      } catch(e) {
        console.warn("無法儲存到 localStorage:", e);
      }
      try {
        sessionStorage.setItem(deviceIdKey, deviceId);
        sessionStorage.setItem(`${deviceIdKey}_timestamp`, timestamp);
        successCount++;
      } catch(e) {
        console.warn("無法儲存到 sessionStorage:", e);
      }
      if (StorageUtils.setCookie(deviceIdKey, deviceId, 365)) {
        successCount++;
      }
      return successCount > 0;
    },
    validateStoredIds: () => {
      const deviceIdKey = App.config.DEVICE_ID_KEY;
      const appVersion = App.config.APP_VERSION || 'unknown';

      const localId = localStorage.getItem(deviceIdKey);
      const sessionId = sessionStorage.getItem(deviceIdKey);
      const cookieId = StorageUtils.getCookie(deviceIdKey);
      let primaryId = null;
      let needsRepair = false;
      if (localId) {
        primaryId = localId;
        const storedVersion = localStorage.getItem(`${deviceIdKey}_version`);
        if (storedVersion !== appVersion) {
          localStorage.setItem(`${deviceIdKey}_version`, appVersion);
        }
      } else if (cookieId) {
        primaryId = cookieId;
        needsRepair = true;
      } else if (sessionId) {
        primaryId = sessionId;
        needsRepair = true;
      }
      if (primaryId && (primaryId!==localId || primaryId!==sessionId || primaryId!==cookieId)) {
        needsRepair = true;
      }
      if (primaryId && needsRepair) {
        console.log("檢測到識別碼不一致，正在修復...");
        DeviceIdManager.saveDeviceId(primaryId);
        return true;
      }
      return false;
    },
    getDeviceUuid: async () => {
      const deviceIdKey = App.config.DEVICE_ID_KEY;
      const storedId = localStorage.getItem(deviceIdKey) ||
                       sessionStorage.getItem(deviceIdKey) ||
                       StorageUtils.getCookie(deviceIdKey);
      if (storedId) {
        DeviceIdManager.validateStoredIds();
        return storedId;
      }
      try {
        const stableFeatures = [
          navigator.platform || '',
          `${screen.width}x${screen.height}`,
          navigator.language || '',
          screen.colorDepth || 0,
          navigator.userAgent.substring(0,100)
        ];
        if (!BrowserUtils.isSafari()) {
          stableFeatures.push(navigator.hardwareConcurrency || 0);
          stableFeatures.push(window.devicePixelRatio || 0);
          stableFeatures.push(BrowserUtils.getWebGLFingerprint());
        }
        const ip = await NetworkUtils.getIPAddress().catch(()=> 'unknown');
        stableFeatures.push(ip);
        const deviceFingerprint = stableFeatures.join('|');
        const newId = await CryptoUtils.hashString(deviceFingerprint)
          .catch(()=> CryptoUtils.generateFallbackID());
        DeviceIdManager.saveDeviceId(newId);
        return newId;
      } catch(err) {
        console.error("生成設備識別碼失敗:", err);
        const fallbackId = CryptoUtils.generateFallbackID();
        DeviceIdManager.saveDeviceId(fallbackId);
        return fallbackId;
      }
    },
    recoverDeviceId: () => {
      const deviceIdKey = App.config.DEVICE_ID_KEY;
      const validId = localStorage.getItem(deviceIdKey) ||
                      sessionStorage.getItem(deviceIdKey) ||
                      StorageUtils.getCookie(deviceIdKey);
      if (validId) {
        DeviceIdManager.saveDeviceId(validId);
        console.log('已恢復設備識別碼');
      }
    }
  };

  /**
   * 裝置資訊管理
   */
  const DeviceInfoManager = {
    collectDeviceInfo: () => {
      return {
        OS: navigator.platform || "未知",
        Browser: navigator.userAgent || "未知",
        ScreenResolution: `${window.screen.width}x${window.screen.height}`,
        TimeZone: Intl.DateTimeFormat().resolvedOptions().timeZone || "未知",
        Language: navigator.language || "未知",
        HardwareConcurrency: navigator.hardwareConcurrency || 0,
        DeviceMemory: navigator.deviceMemory || 0,
        MaxTouchPoints: navigator.maxTouchPoints || 0,
        ColorDepth: window.screen.colorDepth || 24,
        PixelRatio: window.devicePixelRatio || 1,
        IsSafari: BrowserUtils.isSafari(),
        TouchSupport: ('ontouchstart' in window) || (navigator.maxTouchPoints>0),
        AvailScreen: `${window.screen.availWidth}x${window.screen.availHeight}`,
        Fonts: BrowserUtils.detectCommonFonts(),
        AppVersion: App.config.APP_VERSION || 'unknown'
      };
    }
  };

  /**
   * 位置管理 (多點定位)
   */
  const LocationManager = {
    getCurrentLocation: async (sampleCount = 5, sampleInterval = 500) => {
      return new Promise(async (resolve, reject) => {
        if (!("geolocation" in navigator)) {
          reject(new Error("瀏覽器不支援定位"));
          return;
        }

        UIUtils.displayLoading(`定位中｜採集 ${sampleCount} 個位置樣本...`);

        let completedSamples = 0;
        const updateProgress = () => {
          completedSamples++;
          UIUtils.displayLoading(`定位中｜已完成 ${completedSamples}/${sampleCount} 個樣本...`);
        };

        const getSingleLocation = (index) => {
          return new Promise((resolveLoc, rejectLoc) => {
            setTimeout(() => {
              navigator.geolocation.getCurrentPosition(
                (position) => {
                  let ts = position.timestamp;
                  if (ts < 1e11) ts = ts * 1000;
                  console.log(`位置採樣 #${index + 1} 成功: 精度=${position.coords.accuracy}m`);
                  updateProgress();
                  resolveLoc({
                    lat: position.coords.latitude,
                    lng: position.coords.longitude,
                    accuracy: position.coords.accuracy,
                    timestamp: ts,
                    sampleIndex: index,
                  });
                },
                (error) => {
                  console.warn(`位置採樣 #${index + 1} 失敗: ${error.code}-${error.message}`);
                  updateProgress();
                  rejectLoc(error);
                },
                {
                  enableHighAccuracy: true,
                  timeout: 10000, // 採用參考檔案的 10 秒超時設定
                  maximumAge: 0,
                }
              );
            }, index * sampleInterval);
          });
        };

        try {
          const locationPromises = Array.from({ length: sampleCount }, (_, i) =>
            getSingleLocation(i)
          );
          // 使用 allSettled 來確保即使部分採樣失敗，成功的仍然會被處理
          const results = await Promise.allSettled(locationPromises);
          const successfulResults = results
            .filter((r) => r.status === "fulfilled" && r.value !== null)
            .map((r) => r.value);

          console.log(`位置採樣結果: ${successfulResults.length}/${sampleCount} 成功`);

          if (successfulResults.length > 0) {
            const sortedCoords = successfulResults.sort(
              (a, b) => (a.accuracy || 9999) - (b.accuracy || 9999)
            );
            resolve(sortedCoords);
          } else {
            // 核心修正：智能分析所有失敗的錯誤代碼
            const rejectedResults = results.filter(r => r.status === 'rejected');
            
            // 檢查是否有任何一個錯誤是權限被拒 (code === 1)
            const hasPermissionDenied = rejectedResults.some(r => 
              r.reason && r.reason.code === 1
            );
            
            if (hasPermissionDenied) {
              // 拋出包含 "denied" 關鍵字的錯誤，確保上層能正確識別
              reject(new Error("User denied geolocation access"));
            } else {
              // 如果不是權限問題，則拋出其他錯誤或通用錯誤
              const firstError = rejectedResults[0]?.reason;
              const errorMessage = firstError?.message || "無法獲取任何有效的位置數據";
              reject(new Error(errorMessage));
            }
          }
        } catch (err) {
          reject(new Error(`定位系統錯誤: ${err.message || err}`));
        }
      });
    },
  };

  /**
   * 輔助函數：透過模擬點擊<a>元素（指定在新分頁開啟）來嘗試觸發 App Link / Universal Link
   */
  function triggerAppLink(url) {
    const a = document.createElement('a');
    a.href = url;
    a.target = '_blank'; // 確保 target="_blank"
    a.style.display = 'none'; // 使其不可見
    // 可選：a.rel = 'noopener noreferrer'; // 針對 target="_blank" 的安全建議，可先測試不加
    document.body.appendChild(a); // 添加到 DOM
    a.click(); // 模擬點擊
    document.body.removeChild(a); // 從 DOM 中移除
  }

  /**
   * GPS增強定位數據處理
   */
  function enhancedLocationData(coordsList, userCode) {
    if (!coordsList || coordsList.length === 0) {
      return {
        processedCoords: [],
        accuracy: 0,
        confidence: 0,
        userCode: userCode || 'unknown'
      };
    }
    
    const avgLat = coordsList.reduce((sum, coord) => sum + coord.lat, 0) / coordsList.length;
    const avgLng = coordsList.reduce((sum, coord) => sum + coord.lng, 0) / coordsList.length;
    const avgAccuracy = coordsList.reduce((sum, coord) => sum + coord.accuracy, 0) / coordsList.length;
    
    return {
      processedCoords: [{
        lat: avgLat,
        lng: avgLng,
        accuracy: avgAccuracy,
        timestamp: Date.now()
      }],
      accuracy: avgAccuracy,
      confidence: coordsList.length >= 3 ? 0.8 : 0.6,
      userCode: userCode || 'unknown',
      sampleCount: coordsList.length
    };
  }
  
  /**
   * 主要DOM元素
   */
  const elements = {
    toggleBtn: document.getElementById("toggleBtn"),
    idInput: document.getElementById("idCode"),
    eyeIcon: document.getElementById("eyeIcon"),
    responseDiv: document.getElementById("response"),
    overlay: document.getElementById("overlay"),
    modalTitle: document.getElementById("modal-title"),
    modalMessage: document.getElementById("modal-message"),
    confirmBtn: document.getElementById("modal-confirm"),
    cancelBtn: document.getElementById("modal-cancel"),
    incognitoPrompt: document.getElementById("incognitoPrompt"),
    confirmIncognito: document.getElementById("confirmIncognito"),
    cancelIncognito: document.getElementById("cancelIncognito"),
    videoTutorialModal: document.getElementById("videoTutorialModal"),
    closeVideoModal: document.getElementById("closeVideoModal"),
    tutorialVideo: document.getElementById("tutorialVideo"),
    videoFallback: document.getElementById("videoFallback"),
    networkStatus: document.getElementById("networkStatus"),
    errorSpan: document.querySelector(".error-message"),
    safariNotice: document.getElementById("safariNotice"),
    allButtons: {
      btnOn: document.getElementById("btnOn"),
      btnOff: document.getElementById("btnOff"),
      btnQuery: document.getElementById("btnQuery"),
      btnTest: document.getElementById("btnTest"),
      btnForm: document.getElementById("btnForm"),
      btnSchedule: document.getElementById("btnSchedule"),
      btnNotice: document.getElementById("btnNotice"),
      btnManager: document.getElementById("btnManager"),
      btnPersonal: document.getElementById("btnPersonal")
    }
  };

  /**
   * 原始按鈕樣式
   */
  const originalButtonStyles = {
    btnOn: 'linear-gradient(to right, #28a745 0%, #1e7e34 100%)',
    btnOff: 'linear-gradient(to right, #e74c3c 0%, #c0392b 100%)',
    btnQuery: 'linear-gradient(to right, #3498db 0%, #2980b9 100%)',
    btnTest: 'linear-gradient(to right, #f39c12 0%, #e67e22 100%)',
    btnForm: 'linear-gradient(to right, var(--btn-form-from) 0%, var(--btn-form-to) 100%)',
    btnSchedule: 'linear-gradient(to right, var(--btn-schedule-from) 0%, var(--btn-schedule-to) 100%)',
    btnNotice: 'linear-gradient(to right, var(--btn-notice-from) 0%, var(--btn-notice-to) 100%)',
    btnManager: 'linear-gradient(to right, var(--btn-manager-from) 0%, var(--btn-manager-to) 100%)',
    btnPersonal: 'linear-gradient(to right, var(--btn-personal-from) 0%, var(--btn-personal-to) 100%)'
  };

  /**
   * 按鈕顏色配置
   */
  const buttonColors = {
    btnOn: { from: '#28a745', to: '#1e7e34' },
    btnOff: { from: '#e74c3c', to: '#c0392b' },
    btnQuery: { from: '#3498db', to: '#2980b9' },
    btnTest: { from: '#f39c12', to: '#e67e22' },
    btnForm: { from: '#9c27b0', to: '#673ab7' },
    btnSchedule: { from: '#20c997', to: '#0b845e' },
    btnNotice: { from: '#607d8b', to: '#455a64' }
  };

  /**
   * 選擇器配置中心
   */
  // UI配置現已統一管理於 App.UI_CONFIG，避免重複定義

  /**
     * 應用選擇器樣式函數
     */
    function initPickerStyles() {
      // 應用日期選擇器樣式（使用直接配置定義確保穩定性）
      $(document).on('click', '.datepicker', function() {
        setTimeout(function() {
          // 使用內嵌配置定義，避免跨模組依賴
          const config = {
            width: 320,
            minWidth: 280,
            maxWidth: 360,
            background: '#222',
            borderRadius: '6px',
            boxShadow: '0 4px 12px rgba(0,0,0,0.5)'
          };
          $('#ui-datepicker-div').css({
            'width': `${config.width}px`,
            'min-width': `${config.minWidth}px`,
            'max-width': `${config.maxWidth}px`,
            'background-color': config.background,
            'border': 'none',
            'box-shadow': config.boxShadow,
            'margin-top': '0',
            'padding': '0'
          });
        }, 10);
      });
      
      // 應用時間選擇器樣式
      /* 註解掉全局時間選擇器事件，避免樣式覆蓋衝突
      $(document).on('click', '.timepicker', function() {
        setTimeout(function() {
          // 使用內嵌配置定義，確保系統載入穩定性
          const datepickerConfig = {
            width: 320,
            minWidth: 280,
            maxWidth: 360,
            background: '#222',
            borderRadius: '6px',
            boxShadow: '0 4px 12px rgba(0,0,0,0.5)'
          };
          const timepickerConfig = {
            width: 300,
            minWidth: 300,
            maxWidth: 340,
            background: '#222',
            padding: '15px',
            borderTop: '1px solid #444'
          };
          
          // 首先確保日期選擇器容器正確設置
          $('#ui-datepicker-div').css({
            'width': `${timepickerConfig.width}px`,
            'min-width': `${timepickerConfig.minWidth}px`,
            'max-width': `${timepickerConfig.maxWidth}px`,
            'background-color': datepickerConfig.background,
            'border': 'none',
            'box-shadow': datepickerConfig.boxShadow,
            'margin-top': '0',
            'padding': '0'
          });
          
                    // 然後設置時間選擇器樣式
          $('.ui-timepicker-div').css({
            'width': '100%',
            'min-width': `${timepickerConfig.minWidth}px`,
            'max-width': `${timepickerConfig.maxWidth}px`,
            'background-color': timepickerConfig.background,
            'border': 'none',
            'border-top': timepickerConfig.borderTop,
            'border-radius': '0',
            'padding': timepickerConfig.padding,
            'box-shadow': 'none'
          });
          
          // 加入修復樣式，解決時間選擇器 minuteGrid: 30 的顯示問題
          $('.ui-timepicker-div dl').css({
            'display': 'flex',
            'flex-direction': 'row',
            'justify-content': 'center',  // 改為 center 對齊
            'align-items': 'flex-start',
            'margin': '0',
            'padding': '0'
          });
          
          $('.ui-timepicker-div dd').css({
            'margin': '0 5px',  // 增加水平邊距
            'padding': '0',
            'width': 'auto',    // 自適應寬度
            'text-align': 'center'
          });
        }, 10);
      });
      */
    }

    /* 主要功能函式 (打卡、定位、表單) */



  /**
   * 開啟並追蹤上傳視窗
   * @param {string} url - 上傳頁面 URL
   * @param {string} requestId - 對應的請求 ID
   * @returns {boolean} - 是否成功開啟視窗
   */
  function openAndTrackUploadWindow(url, requestId) {
    const reqIdStr = String(requestId); 
    if (!reqIdStr) {
      console.error("追蹤錯誤：無效的 requestId");
      return false;
    }
    
    if (window.openedUploadWindows[reqIdStr] && !window.openedUploadWindows[reqIdStr].closed) {
      console.warn(`正在為已存在的 requestId ${reqIdStr} 開啟新視窗，可能導致舊視窗無法關閉。`);
    }
    
    try {
      const uploadWindow = window.open(url, '_blank');
      if (uploadWindow) {
        window.openedUploadWindows[reqIdStr] = uploadWindow;
        console.log(`視窗已開啟並儲存追蹤。RequestId: ${reqIdStr}`);
        return true;
      } else {
        console.error(`無法開啟上傳視窗，RequestId: ${reqIdStr}`);
        return false;
      }
    } catch (e) {
      console.error(`開啟視窗異常: ${e.message}`);
      return false;
    }
  }

  /**
   * 顯示表單選擇模態框
   */
  function showFormApplicationModal() {
    const formApplicationModal = document.getElementById("formApplicationModal");
    formApplicationModal.classList.add("show");
    formApplicationModal.classList.add("form-overlay");
  }

  /**
   * 隱藏表單選擇模態框
   */
  function hideFormApplicationModal() {
    const formApplicationModal = document.getElementById("formApplicationModal");
    formApplicationModal.classList.remove("show");
    formApplicationModal.classList.remove("form-overlay");
  }

  /**
   * 顯示影片教學
   */
  function showVideoTutorialModal() {
    elements.videoTutorialModal.style.display = "flex";
    elements.tutorialVideo.play().catch(() => (elements.videoFallback.style.display = "block"));
  }

  /**
   * 隱藏影片教學
   */
  function hideVideoTutorialModal() {
    elements.videoTutorialModal.style.display = "none";
    elements.tutorialVideo.pause();
    elements.videoFallback.style.display = "none";
  }

  /**
   * 打開班表 (最佳化版本)
   * @param {HTMLButtonElement} button - 班表按鈕元素
   */
  function openSchedule(button) {
    const cooldownMs = App.config.COOLDOWN_MS;
    // 系統狀態初始設定
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    
    // 網路連線驗證
    if (!navigator.onLine) {
      UIUtils.displayError("離線無法開啟班表");
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false);
      return;
    }
    
    UIUtils.displayLoading("開啟中｜請稍候...");
    
    try {
      // 直接觸發班表連結開啟
      triggerAppLink(App.config.SCHEDULE_URL);
      UIUtils.displaySuccess("班表已在新視窗中打開");
      
      // 保留記錄追蹤功能
      const idCode = elements.idInput.value || "0000";
      NetworkUtils.callGASWithRetry("logButtonClick", [idCode, 'schedule'])
        .then(logResult => console.log("記錄點擊結果:", logResult?.status || '未知狀態'))
        .catch(logError => console.error("記錄點擊失敗:", logError));
    } catch (error) {
      UIUtils.displayError("開啟班表失敗，請確認瀏覽器設定");
    } finally {
      // 按鈕冷卻機制啟動
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false);
    }
  }

  /**
   * 開啟內部公告 (最佳化版本)
   * @param {HTMLButtonElement} button - 內部公告按鈕元素
   */
  function openNotice(button) {
    const cooldownMs = App.config.COOLDOWN_MS;
    // 系統狀態初始設定
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    
    // 網路連線驗證
    if (!navigator.onLine) {
      UIUtils.displayError("離線無法開啟公告");
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false);
      return;
    }
    
    UIUtils.displayLoading("開啟中｜請稍候...");
    
    try {
      // 直接觸發公告連結開啟
      const noticeUrl = App.config?.NOTICE_URL;
      
      if (!noticeUrl) {
        throw new Error("公告連結未設定，請聯繫系統管理員");
      }
      triggerAppLink(noticeUrl);
      UIUtils.displaySuccess("內部公告已在新視窗中打開");
      
      // 保留記錄追蹤功能
      const idCode = elements.idInput.value || "0000";
      NetworkUtils.callGASWithRetry("logButtonClick", [idCode, 'announcement'])
        .then(logResult => console.log("記錄點擊結果:", logResult?.status || '未知狀態'))
        .catch(logError => console.error("記錄點擊失敗:", logError));
    } catch (error) {
      UIUtils.displayError("開啟公告失敗，請確認瀏覽器設定");
    } finally {
      // 按鈕冷卻機制啟動
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false);
    }
  }

  /**
   * 測試定位功能
   */
  async function testLocation(button) {
    const cooldownMs = App.config.COOLDOWN_MS;
    // 設置操作狀態為進行中
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    
    const idCode = elements.idInput.value;

    if (!validateIdCode(idCode)) {
      UIUtils.displayError("請鍵入完整的身分證後四碼");
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false); // 修正：應重置狀態
      return;
    }
    if (!navigator.onLine) {
      UIUtils.displayError("離線無法測試定位");
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false); // 修正：應重置狀態
      return;
    }
    UIUtils.displayLoading("定位測試中｜請稍候...");

    try {
      // 新增：增加授權檢查，與參考版本和 queryTodayRecords 函式對齊
      const authCheck = await NetworkUtils.callGASWithRetry("checkDeviceBinding", [idCode, "測試定位"]);
              if (authCheck.status === "error") {
        if (["AUTH.NOT_IN_AUTHLIST", "AUTH.NOT_BOUND", "AUTH.DEVICE_MISMATCH", "AUTH.INVALID_USER"].includes(authCheck.error)) {
          UIUtils.displayError("不在授權名單中｜打卡失敗");
        } else {
          UIUtils.displayError(authCheck.message || "授權檢查失敗");
        }
        UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
        UIUtils.enableAllButtonsExcept(button);
        ButtonStateManager.setOperationStatus(false);
        return;
      }
      
      // 修正：直接呼叫標準的 LocationManager，並移除前端篩選邏輯
      UIUtils.displayLoading("開始定位...");
      const coordsList = await LocationManager.getCurrentLocation(5, 300); // 使用標準的 LocationManager 採樣
      
      if (!coordsList || coordsList.length === 0) {
        throw new Error("無法獲取任何有效的位置數據");
      }
      
      // 修復P005：調用GPS增強邏輯進行前端數據預處理
      const enhancedData = enhancedLocationData(coordsList, idCode);
      console.log("GPS增強處理完成:", enhancedData);
      
      // 將增強處理後的座標數據傳送到後端進行分析
      const result = await NetworkUtils.callGASWithRetry("testLocationWithAuth", [idCode, enhancedData.processedCoords]);

      if (result.status === "warning" || result.status === "success") {
        // V4.4 修正：使用 DOMPurifyUtils 來安全地設置 HTML
        const content = result.html ? `<div class="table-responsive">${result.html}</div>` : "無資料";
        DOMPurifyUtils.safeSetHTML(elements.responseDiv, content);
        
        // 移除整個 GPS 信號不穩定警告區塊
        
        elements.responseDiv.classList.add("show");
        UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
        UIUtils.enableAllButtonsExcept(button);
      } else {
        UIUtils.displayError(result.message);
        UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
        UIUtils.enableAllButtonsExcept(button);
      }
    } catch(err) {
      // 修正：使用與 handleAction 相同的錯誤處理邏輯
      let errorMessage = "";
      if (err.message.includes("denied")) { // 這是瀏覽器回報權限問題的關鍵字
        errorMessage = "請允許定位功能開啟｜請至裝置內部開啟";
      } else if (err.message.includes("GPS") || err.message.includes("定位")) {
        errorMessage = `GPS定位失敗｜${err.message}`;
      } else if (err.message.includes("網路") || err.message.includes("連線")) {
        errorMessage = `網路連線問題｜${err.message}`;
      } else {
        errorMessage = `測試失敗｜${err.message}`;
      }
      
      UIUtils.displayError(errorMessage);
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
    } finally {
      ButtonStateManager.setOperationStatus(false);
    }
  }

  /**
   * 查詢考勤歷史記錄 - 優化流程：先顯示選項，後驗證
   */
  async function showQueryMenu(button) {
    const idCode = elements.idInput.value;
    if (!validateIdCode(idCode)) {
      UIUtils.displayError("請先輸入有效的4位數字ID");
      return;
    }

    // 直接顯示查詢類型選擇，不進行按鈕禁用和冷卻
    UIUtils.clearResponse();
    showQueryTypeModal();
  }

  /**
   * 顯示查詢類型選擇模態框
   * Material Icons + form-modal架構
   */
  function showQueryTypeModal() {
    const userCode = elements.idInput.value;
    
    // Material Icons模態框結構
    const modalHTML = `
      <div id="queryTypeModal" class="overlay show form-overlay" style="display: flex;">
        <div class="form-modal" style="max-width: 400px;">
          <div class="form-modal-header">
            <h3 style="color: #3498db; text-align: center; width: 100%;">選擇查詢類型</h3>
          </div>
          
          <div class="form-modal-content">
            <div class="form-item-container">
              <div class="form-item maintenance" id="btnAttendanceQuery" style="cursor: not-allowed;">
                <div class="form-icon">
                  <span class="material-icons">event_note</span>
                </div>
                <div class="form-info">
                  <h4>出勤紀錄</h4>
                  <p>查看完整月度出勤記錄</p>
                </div>
                <div class="form-arrow">
                  <span class="material-icons">chevron_right</span>
                </div>
              </div>
              
              <div class="form-item" id="btnFormQuery" style="cursor: pointer;">
                <div class="form-icon">
                  <span class="material-icons">description</span>
                </div>
                <div class="form-info">
                  <h4>表單紀錄</h4>
                  <p>查看申請表單審核狀態</p>
                </div>
                <div class="form-arrow">
                  <span class="material-icons">chevron_right</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    // 插入模態框到頁面
    document.body.insertAdjacentHTML('beforeend', modalHTML);
    
    // 綁定事件監聽器
    document.getElementById('btnAttendanceQuery').onclick = (event) => {
      // 維修中
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      return false;
    };
    
    document.getElementById('btnFormQuery').onclick = () => {
      closeQueryTypeModal();
      showFormTypeSelector(userCode);
    };
    
    // 點擊背景顯示確認對話框
    const queryModal = document.getElementById('queryTypeModal');
    if (queryModal) {
      queryModal.onclick = (e) => {
        if (e.target === queryModal) {
          UIUtils.showCustomConfirm(
            "離開確認",
            "即將關閉查詢選擇｜返回主畫面",
            () => {
              closeQueryTypeModal();
            },
            () => {
              // 取消：什麼都不做
            }
          );
        }
      };
    }
  }

  /**
   * 關閉查詢類型模態框
   */
  function closeQueryTypeModal() {
    const modal = document.getElementById('queryTypeModal');
    if (modal) modal.remove();
  }

  /**
   * 執行出勤記錄查詢 - AbortController 整合版
   */
  async function showAttendanceQuery(userCode) {
    const controller = new AbortController();
    const signal = controller.signal;
    let removeCancellationHandlers;
    const cooldownMs = 5000; // 查詢完成後的冷卻時間
    
    // 禁用所有按鈕，開始查詢
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    UIUtils.displayLoading("處理中｜可點擊取消...");

    try {
      // 綁定取消事件
      removeCancellationHandlers = setupCancellationHandler(controller, '.overlay.show');
      
      // 驗證授權
      const authResult = await NetworkUtils.callGASWithRetry('checkDeviceBinding', [userCode, "查詢驗證"], undefined, signal);
      if (authResult.status !== "success") {
        if (["AUTH.NOT_IN_AUTHLIST", "AUTH.NOT_BOUND", "AUTH.DEVICE_MISMATCH", "AUTH.INVALID_USER"].includes(authResult.error)) {
          throw new Error("不在授權名單中｜查詢失敗");
        } else {
          throw new Error(authResult.message || "驗證授權失敗");
        }
      }

      // 查詢出勤資料
      const result = await NetworkUtils.callGASWithRetry('queryMonthlyAttendance', [userCode], undefined, signal);

      if (NetworkUtils.checkJSONPError(result) || !result.html) {
        throw new Error(result.message || '出勤記錄載入失敗');
      }
      
      // 動態DOM生成模式
      const responseDiv = elements.responseDiv;
      responseDiv.innerHTML = '';
      
      if (result.status !== 'success' || !result.data) {
        throw new Error('資料格式錯誤');
      }
      
      // 標題區
      const headerDiv = document.createElement('div');
      headerDiv.style.cssText = 'margin-bottom:10px;text-align:center;color:#2ecc71;font-weight:bold;';
      headerDiv.textContent = `${result.userName} 本月出勤記錄 (共 ${result.recordCount} 筆)`;
      responseDiv.appendChild(headerDiv);
      
      // 表格DOM生成
      const table = document.createElement('table');
      table.style.cssText = 'width:100%;border-collapse:collapse;text-align:center;border:1px solid #555;font-size:14px;color:#fff;';
      
      // 表頭
      const thead = document.createElement('thead');
      thead.style.cssText = 'background-color:#444;color:#FFF;font-weight:bold;';
      thead.innerHTML = `
        <tr>
          <th style='padding:6px;border:1px solid #666;'>日期</th>
          <th style='padding:6px;border:1px solid #666;'>上班時間</th>
          <th style='padding:6px;border:1px solid #666;'>下班時間</th>
          <th style='padding:6px;border:1px solid #666;'>狀態</th>
        </tr>
      `;
      table.appendChild(thead);
      
      // 表身
      const tbody = document.createElement('tbody');
      result.data.forEach(record => {
        const tr = document.createElement('tr');
        
        let statusColor = '#e74c3c';
        if (record.status === '正常') statusColor = '#2ecc71';
        else if (record.status === '未下班') statusColor = '#f39c12';
        
        tr.innerHTML = `
          <td style='padding:6px;border:1px solid #666;color:#76ff03;'>${record.dateDisplay}</td>
          <td style='padding:6px;border:1px solid #666;color:#2ecc71;'>${record.clockInTime}</td>
          <td style='padding:6px;border:1px solid #666;color:#e74c3c;'>${record.clockOutTime}</td>
          <td style='padding:6px;border:1px solid #666;color:${statusColor};'>${record.status}</td>
        `;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      
      responseDiv.appendChild(table);
      responseDiv.classList.add("show");

    } catch (err) {
      console.error("出勤記錄查詢失敗:", err);
      if (err.message === 'AbortError') {
        UIUtils.displayMessage('查詢已取消');
      } else if (!navigator.onLine) {
        UIUtils.displayError('網路連線中斷');
      } else {
        UIUtils.displayError(`出勤記錄載入失敗: ${err.message}`);
      }
    } finally {
      UIUtils.hideLoading();
      
      // 清理取消事件處理器
      if (removeCancellationHandlers) {
        removeCancellationHandlers();
      }
      
      // 重新啟用按鈕並開始冷卻
      const queryButton = elements.allButtons.btnQuery;
      if (queryButton) {
        UIUtils.startIndividualCooldown(queryButton, cooldownMs, queryButton.dataset.originalText);
        UIUtils.enableAllButtonsExcept(queryButton);
      } else {
        UIUtils.enableAllButtons();
      }
      
      ButtonStateManager.setOperationStatus(false);
    }
    
    // 模態框取消機制
    const setupCancelHandler = () => {
      document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') controller.abort();
      });
      
      const modal = document.querySelector('.form-overlay');
      if (modal) {
        modal.addEventListener('click', (e) => {
          if (e.target === modal) controller.abort();
        });
      }
    };
    setupCancelHandler();
  }

  /**
   * 顯示表單類型選擇器
   * Material Icons四分類設計
   */
  function showFormTypeSelector(userCode) {
    const formModalHTML = `
      <div id="formTypeModal" class="overlay show form-overlay" style="display: flex;">
        <div class="form-modal" style="max-width: 450px;">
          <div class="form-modal-header">
            <button id="btnCloseFormModal" class="back-button-shadow">
              <span class="material-icons">arrow_back</span>
              <span class="back-text">返回</span>
            </button>
            <h3 style="color: #3498db; text-align: center; width: 100%;">選擇表單類型</h3>
          </div>
          
          <div class="form-modal-content">
            <div class="form-item-container">
              <div class="form-item form-type-btn" data-type="card" style="cursor: pointer;" id="cardCorrectionFormQuery">
                <div class="form-icon">
                  <span class="material-icons">credit_card</span>
                </div>
                <div class="form-info">
                  <h4>補卡申請</h4>
                  <p>忘記打卡或打卡錯誤</p>
                </div>
                <div class="form-arrow">
                  <span class="material-icons">chevron_right</span>
                </div>
              </div>
              
              <div class="form-item form-type-btn" data-type="overtime" style="cursor: pointer;" id="overtimeFormQuery">
                <div class="form-icon">
                  <span class="material-icons">schedule</span>
                </div>
                <div class="form-info">
                  <h4>加班申請</h4>
                  <p>提交加班時數及原因</p>
                </div>
                <div class="form-arrow">
                  <span class="material-icons">chevron_right</span>
                </div>
              </div>
              
              <div class="form-item form-type-btn" data-type="leave" style="cursor: pointer;" id="leaveRequestFormQuery">
                <div class="form-icon">
                  <span class="material-icons">event_busy</span>
                </div>
                <div class="form-info">
                  <h4>請假申請</h4>
                  <p>請假日期及假別選擇</p>
                </div>
                <div class="form-arrow">
                  <span class="material-icons">chevron_right</span>
                </div>
              </div>
              
              <div class="form-item form-type-btn" data-type="shift" style="cursor: pointer;" id="shiftChangeFormQuery">
                <div class="form-icon">
                  <span class="material-icons">compare_arrows</span>
                </div>
                <div class="form-info">
                  <h4>異動班別</h4>
                  <p>變更排班班別</p>
                </div>
                <div class="form-arrow">
                  <span class="material-icons">chevron_right</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    document.body.insertAdjacentHTML('beforeend', formModalHTML);
    
    // 綁定表單類型按鈕事件
    document.querySelectorAll('.form-type-btn').forEach(btn => {
      btn.onclick = () => {
        const formType = btn.dataset.type;
        closeFormTypeModal();
        executeFormQuery(userCode, formType);
      };
    });
    
    document.getElementById('btnCloseFormModal').onclick = () => {
      closeFormTypeModal();
      showQueryTypeModal(userCode); // 返回主選單
    };
    
    document.getElementById('formTypeModal').onclick = (e) => {
      if (e.target.id === 'formTypeModal') {
        UIUtils.showCustomConfirm(
          "離開確認",
          "即將關閉視窗｜返回主畫面",
          () => {
            closeFormTypeModal();
          },
          () => {
            // 取消：什麼都不做
          }
        );
      }
    };
  }

  /**
   * 關閉表單類型模態框
   */
  function closeFormTypeModal() {
    const modal = document.getElementById('formTypeModal');
    if (modal) modal.remove();
  }

  /**
   * 執行表單記錄查詢 - 加入驗證步驟並修正API呼叫
   */
  async function executeFormQuery(userCode, formType) {
    const formTypeNames = {
      card: '補卡申請',
      overtime: '加班申請',
      leave: '請假申請',
      shift: '班別異動'
    };
    
    const cooldownMs = 5000; // 查詢完成後的冷卻時間

    // 禁用所有按鈕，開始查詢
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    UIUtils.displayLoading(`驗證並載入${formTypeNames[formType]}記錄中...`);

    try {
      // 在查詢資料前執行驗證
      const authResult = await NetworkUtils.callGASWithRetry('checkDeviceBinding', [userCode, "查詢驗證"]);
      if (authResult.status !== "success") {
        if (["AUTH.NOT_IN_AUTHLIST", "AUTH.NOT_BOUND", "AUTH.DEVICE_MISMATCH", "AUTH.INVALID_USER"].includes(authResult.error)) {
          throw new Error("不在授權名單中｜查詢失敗");
        } else {
          throw new Error(authResult.message || "驗證授權失敗");
        }
      }
      
      // 調用後端表單查詢API
      const result = await NetworkUtils.callGASWithRetry('queryFormsByType', [userCode, formType]);

      if (NetworkUtils.checkJSONPError(result) || !result.html) {
        throw new Error(result.message || '表單記錄載入失敗');
      }
      
      const responseDiv = elements.responseDiv;
      
      // 臨時相容模式 - 保持原有HTML注入至後端API完全重構
      if (result.html) {
        DOMPurifyUtils.safeSetHTML(responseDiv, result.html);
      } else if (result.data) {
        // 未來JSON模式預留介面
        responseDiv.innerHTML = '<div style="color:#f39c12;">表單記錄JSON模式待實施</div>';
      }
      
      responseDiv.classList.add("show");

    } catch (err) {
      console.error("表單記錄查詢失敗:", err);
      UIUtils.displayError(`${formTypeNames[formType]}記錄載入失敗: ${err.message}`);
    } finally {
      // 重新啟用按鈕並開始冷卻
      const queryButton = elements.allButtons.btnQuery;
      if (queryButton) {
        UIUtils.startIndividualCooldown(queryButton, cooldownMs, queryButton.dataset.originalText);
        UIUtils.enableAllButtonsExcept(queryButton);
      } else {
        UIUtils.enableAllButtons();
      }
      
      ButtonStateManager.setOperationStatus(false);
    }
  }
  async function startClockAction(action, force, button) {
    const cooldownMs = App.config.COOLDOWN_MS;
    // 設置操作狀態為進行中
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    
    const idCode = elements.idInput.value;

    if (!validateIdCode(idCode)) {
      UIUtils.displayError("請鍵入完整的身分證後四碼");
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      return;
    }
    if (!navigator.onLine) {
      UIUtils.displayError("請確認網路連線");
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      return;
    }
    UIUtils.displayLoading("檢查中｜請稍候...");

    try {
      const result = await NetworkUtils.callGASWithRetry("checkDeviceBinding", [idCode, `${action}打卡`]);

      if (result.status === "success") {
        startCountdown(action, force, false, button);
      } else if (result.status === "error") {
        if (result.error === "AUTH.NOT_IN_AUTHLIST" || result.error === "AUTH.DEVICE_MISMATCH" || result.error === "AUTH.INVALID_USER") {
          UIUtils.displayError("不在授權名單中｜打卡失敗");
          UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
          UIUtils.enableAllButtonsExcept(button);
          ButtonStateManager.setOperationStatus(false);
        } else if (result.error === "AUTH.NOT_BOUND") {
          const browser = BrowserUtils.parseBrowser(navigator.userAgent);
          // 使用與打卡確認相同的 overlay 和 modal
          elements.overlay.classList.add("show");
          elements.modalTitle.textContent = "裝置綁定確認";
          elements.modalMessage.textContent = `是否同意綁定 ${browser}`;
          
          // 確定按鈕事件
          elements.confirmBtn.onclick = () => {
            elements.overlay.classList.remove("show");
            startCountdown(action, force, true, button);
          };
          
          // 取消按鈕事件
          elements.cancelBtn.onclick = () => {
            elements.overlay.classList.remove("show");
            UIUtils.displayError("已取消打卡");
            UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
            UIUtils.enableAllButtonsExcept(button);
            ButtonStateManager.setOperationStatus(false);
          };
        } else {
          UIUtils.displayError(result.message);
          UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
          UIUtils.enableAllButtonsExcept(button);
          ButtonStateManager.setOperationStatus(false);
        }
      }
    } catch(err) {
      UIUtils.displayError(`檢查失敗: ${err.message}`);
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false);
    }
  }

  /**
   * 5 秒倒數
   */
  function startCountdown(action, force, bindDevice, button) {
    const cooldownMs = App.config.COOLDOWN_MS;
    // 保持操作狀態為進行中
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    
    elements.overlay.classList.add("show");
    elements.modalTitle.textContent = "打卡確認";

    let countdown = 5;
    elements.modalMessage.textContent = `將於 ${countdown} 秒後建立 ${action} 打卡`;

    const timer = setInterval(()=>{
      countdown--;
      elements.modalMessage.textContent = `將於 ${countdown} 秒後建立 ${action} 打卡`;

      if (countdown<=0) {
        clearInterval(timer);
        elements.overlay.classList.remove("show");
        handleAction(action, force, bindDevice, button);
      }
    }, 1000);

    elements.confirmBtn.onclick = () => {
      clearInterval(timer);
      elements.overlay.classList.remove("show");
      handleAction(action, force, bindDevice, button);
    };
    elements.cancelBtn.onclick = () => {
      clearInterval(timer);
      elements.overlay.classList.remove("show");
      UIUtils.displayError("已取消打卡");
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
      ButtonStateManager.setOperationStatus(false);
    };
  }

  /**
   * 實際打卡動作 (多點定位 + JSONP)
   */
  async function handleAction(action, force, bindDevice, button) {
    const cooldownMs = App.config.COOLDOWN_MS;
    // 保持操作狀態為進行中
    ButtonStateManager.setOperationStatus(true);
    UIUtils.disableAllButtons();
    
    const idCode = elements.idInput.value;
    UIUtils.displayLoading("審核中｜請稍候...");

    try {
      // 使用新的GPS定位邏輯
      const coordsList = await LocationManager.getCurrentLocation(5, 300);
      
      // 修復P005：調用GPS增強邏輯進行前端數據預處理
      const enhancedData = enhancedLocationData(coordsList, idCode);
      console.log("打卡GPS增強處理完成:", enhancedData);
      
      const deviceInfo = JSON.stringify(DeviceInfoManager.collectDeviceInfo());
      const currentDeviceInfo = JSON.stringify(DeviceInfoManager.collectDeviceInfo());
      const uuid = await DeviceIdManager.getDeviceUuid();
      const ip = await NetworkUtils.getIPAddress();
      console.log(`開始${action}打卡請求...`);

      const result = await NetworkUtils.callGASWithRetry("gpsCheckin", [
        idCode, deviceInfo, action, enhancedData.processedCoords, force, bindDevice, uuid, ip, currentDeviceInfo
      ]);

      if (result.status === "success") {
        let finalMsg = result.message;
        if (result.detail?.time) {
          finalMsg += `<br><span style="color:#fff;">${result.detail.time}</span>`;
        }
        
        // 移除安全邊界顯示
        
                // 增加對信號穩定度的顯示
        if (result.detail?.stability !== undefined) {
          const stabilityPercent = Math.round(result.detail.stability * 100);
          const stabilityColor = stabilityPercent >= 70 ? '#4CAF50' : stabilityPercent >= 50 ? '#FFC107' : '#FF5722';
          finalMsg += `<br><span style="color:${stabilityColor};">GPS 信號穩定度: ${stabilityPercent}%</span>`;
        }
        
        if (result.detail?.gpsAccuracy) {
          finalMsg += `<br><span style="color:#76ff03;">${result.detail.gpsAccuracy}</span>`;
        }
        
        UIUtils.displaySuccess(finalMsg);
        UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
        UIUtils.enableAllButtonsExcept(button);
      } else {
        let finalMsg = result.message || "系統錯誤｜請稍後再試";
        if (result.error === "SHIFT.INSUFFICIENT_REST") {
          finalMsg = "休息時間不足｜請稍後再試";
        } else if (result.error === "SYS.RATE_LIMIT_EXCEEDED") {
          finalMsg = "操作過於頻繁｜請稍後再試";
        } else if (result.error === "GPS.OUT_OF_RANGE" || result.error === "GPS.RANGE_ERROR") {
          finalMsg = "超出可打卡範圍｜請重新開關網路";
        } else if (result.error === "BROWSER.UNSUPPORTED") {
          finalMsg = "請使用支援的瀏覽器｜Chrome 或 Safari";
        } else if (result.error === "AUTH.DEVICE_MISMATCH") {
          finalMsg += "<br><span style='color:yellow;'>裝置異常｜請點擊 AI Bot</span>";
        } else if (result.error === "AUTH.DEVICE_TYPE_MISMATCH") {
          finalMsg += "<br><span style='color:yellow;'>綁定裝置不符｜請點擊 AI Bot</span>";
        }
        if (result.detail) {
          finalMsg += `<br><span style="font-size:14px;">詳細資訊: ${JSON.stringify(result.detail)}</span>`;
        }
        UIUtils.displayError(finalMsg);
        UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
        UIUtils.enableAllButtonsExcept(button);
      }
    } catch(err) {
      console.error("打卡流程錯誤:", err);
      
      // 根據錯誤類型顯示不同的錯誤訊息
      let errorMessage = "";
      if (err.message.includes("denied")) {
        errorMessage = "請允許定位功能開啟｜請至裝置內部開啟";
      } else if (err.message.includes("GPS") || err.message.includes("定位")) {
        errorMessage = `GPS定位失敗｜${err.message}`;
      } else if (err.message.includes("網路") || err.message.includes("連線")) {
        errorMessage = `網路連線問題｜${err.message}`;
      } else {
        errorMessage = `打卡失敗｜${err.message}`;
      }
      
      UIUtils.displayError(errorMessage);
      UIUtils.startIndividualCooldown(button, cooldownMs, button.dataset.originalText);
      UIUtils.enableAllButtonsExcept(button);
    } finally {
      ButtonStateManager.setOperationStatus(false);
    }
  }

  // 移除：createTextInputGroup 已遷移至 T-form-handler.ui.js

  // 初始化表單提交函式（使用依賴注入）
  let formSubmitters = null;

  /**
   * 初始化表單模組整合
   */
  function initializeFormIntegration() {
    if (App.FormHandler && App.FormHandler.createFormSubmitters) {
      // 注入相依性到表單模組
      formSubmitters = App.FormHandler.createFormSubmitters({
        elements,
        UIUtils,
        NetworkUtils,
        showUploadPrompt
      });
    } else {
      console.error('表單處理模組未載入，請確保 form-handler.ui.js 已正確載入');
    }
  }

  /**
   * 顯示補卡表單 - 移除前置驗證，提交時才驗證
   */
  function showCardCorrectionForm() {
    hideFormApplicationModal();
    
    try {
      if (App.FormHandler && App.FormHandler.createCardCorrectionFormContent && formSubmitters) {
        createFormModal("補卡申請", App.FormHandler.createCardCorrectionFormContent(), formSubmitters.submitCardCorrection);
      } else {
        UIUtils.displayError('表單模組未正確載入');
      }
    } catch (err) {
      UIUtils.displayError('表單顯示失敗: ' + err.message);
    }
  }

  /**
   * 顯示加班表單 - 移除前置驗證，提交時才驗證
   */
  function showOvertimeForm() {
    hideFormApplicationModal();
    
    try {
      if (App.FormHandler && App.FormHandler.createOvertimeFormContent && formSubmitters) {
        createFormModal("加班申請", App.FormHandler.createOvertimeFormContent(), formSubmitters.submitOvertimeRequest);
      } else {
        UIUtils.displayError('表單模組未正確載入');
      }
    } catch (err) {
      UIUtils.displayError('表單顯示失敗: ' + err.message);
    }
  }

  /**
   * 顯示請假表單 - 移除前置驗證，提交時才驗證
   */
  function showLeaveRequestForm() {
    hideFormApplicationModal();
    
    try {
      if (App.FormHandler && App.FormHandler.createLeaveRequestFormContent && formSubmitters) {
        createFormModal("請假申請", App.FormHandler.createLeaveRequestFormContent(), formSubmitters.submitLeaveRequest);
      } else {
        UIUtils.displayError('表單模組未正確載入');
      }
    } catch (err) {
      UIUtils.displayError('表單顯示失敗: ' + err.message);
    }
  }

  /**
   * 顯示班別異動表單 - 移除前置驗證，提交時才驗證
   */
  function showShiftChangeForm() {
    hideFormApplicationModal();
    
    try {
      if (App.FormHandler && App.FormHandler.createShiftChangeFormContent && formSubmitters) {
        createFormModal("異動班別", App.FormHandler.createShiftChangeFormContent(), formSubmitters.submitShiftChange);
      } else {
        UIUtils.displayError('表單模組未正確載入');
      }
    } catch (err) {
      UIUtils.displayError('表單顯示失敗: ' + err.message);
    }
  }
  
  /**
   * 創建動態按鈕布局 - 統一使用兩按鈕布局
   * @param {HTMLElement} container - 按鈕容器
   * @param {HTMLElement} cancelButton - 取消按鈕
   * @param {HTMLElement} submitButton - 提交按鈕
   * @param {string} formTitle - 表單標題
   */
  function createDynamicButtonLayout(container, cancelButton, submitButton, formTitle) {
    // 清空容器
    container.innerHTML = '';
    
    // 所有表單統一使用兩按鈕布局，上傳按鈕只在提交後動態添加
    container.className = "form-buttons-container two-buttons";
    
    // 統一的按鈕樣式，移除圖標保持簡潔風格（與補卡/加班表單一致）
    cancelButton.textContent = '取消';
    submitButton.textContent = '提交';
    
    // 添加按鈕到容器
    container.appendChild(cancelButton);
    container.appendChild(submitButton);
  }

  /**
   * 顯示手動上傳提示框 - 在現有表單 Modal 中動態新增上傳按鈕
   * @param {string} requestId - 申請ID
   * @param {string} uploadUrl - 上傳頁面的URL
   */
  function showUploadPrompt(requestId, uploadUrl) {
    // 獲取現有的表單 Modal
    const existingModal = document.getElementById('dynamicFormModal');
    if (!existingModal) {
      console.error('找不到現有的表單 Modal');
      return;
    }
    
    // 找到按鈕容器
    const formButtonsContainer = existingModal.querySelector('.form-buttons-container');
    if (!formButtonsContainer) {
      console.error('找不到按鈕容器');
      return;
    }
    
    // 構建上傳 URL
    const baseUrl = uploadUrl || App.config.WEB_APP_URL;
    const fullUploadUrl = `${baseUrl}?page=fileUpload&requestId=${requestId}&formType=leave&lastFour=${elements.idInput.value}`;
    
    // 【版本三：即時回饋版】專業的平滑過渡
    // 先讓現有按鈕優雅淡出
    const existingButtons = formButtonsContainer.querySelectorAll('button');
    existingButtons.forEach(btn => {
      btn.style.transition = 'all 0.5s ease';
      btn.style.opacity = '0';
      btn.style.transform = 'scale(0.9)';
    });
    
    // 延遲後清空並建立新按鈕
    setTimeout(() => {
      formButtonsContainer.innerHTML = '';
      
      // 確保容器只有一個按鈕時也能居中
      formButtonsContainer.className = 'form-buttons-container one-button';
      
      // 建立唯一的上傳按鈕（帶專業滑入動畫）
      const uploadBtn = document.createElement('button');
      uploadBtn.id = 'openUploadBtn';
      uploadBtn.className = 'btn btn-upload'; // 使用專屬的上傳樣式
      uploadBtn.innerHTML = '<span class="material-icons">attach_file</span> 上傳佐證檔案';
      uploadBtn.style.cssText = `
        background: linear-gradient(45deg, #17a2b8, #138496);
        animation: slideInFade 0.5s ease-out forwards;
        opacity: 0;
        transform: translateY(20px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        font-weight: 500;
      `;
      
      // 加入這唯一的新按鈕
      formButtonsContainer.appendChild(uploadBtn);
      
      // 觸發優雅的滑入動畫
      setTimeout(() => {
        uploadBtn.style.transition = 'all 0.5s ease';
        uploadBtn.style.opacity = '1';
        uploadBtn.style.transform = 'translateY(0)';
      }, 50);
      
      // 更新按鈕功能 - 移入 setTimeout 內確保按鈕存在後才綁定
      uploadBtn.onclick = function() {
        openUploadWindow(fullUploadUrl, requestId, existingModal);
      };
    }, 500);
    
    console.log('已顯示單一上傳按鈕');
    
    // 儲存 requestId 以便後續使用
    existingModal.dataset.uploadRequestId = requestId;
  }

  /**
   * 構建完整的上傳 URL
   * @param {string} baseUrl - 基礎 URL
   * @param {string} requestId - 申請 ID
   * @returns {string} 完整的上傳 URL
   */
  function constructUploadUrl(baseUrl, requestId) {
    // 取得身分證後四碼
    const lastFour = elements.idInput.value;
    
    // 如果 URL 已包含參數，使用 & 連接，否則使用 ?
    const separator = baseUrl.includes('?') ? '&' : '?';
    
    // 添加所有必要參數：requestId, formType, lastFour
    return `${baseUrl}${separator}requestId=${requestId}&formType=leave&lastFour=${lastFour}&timestamp=${Date.now()}`;
  }
  
  /**
   * 開啟上傳視窗
   * @param {string} uploadUrl - 上傳 URL
   * @param {string} requestId - 申請 ID
   * @param {HTMLElement} modalElement - Modal 元素
   */
  function openUploadWindow(uploadUrl, requestId, modalElement) {
    const uploadBtn = document.getElementById('openUploadBtn');
    
    // 更新按鈕狀態
    uploadBtn.disabled = true;
    uploadBtn.innerHTML = `
      <span class="material-icons" style="animation: spin 1s linear infinite; font-size: 20px; margin-right: 5px;">sync</span>
      開啟中...
    `;
    
    // 嘗試開啟新視窗
    const uploadWindow = window.open(uploadUrl, '_blank');
    
    if (uploadWindow) {
      // 成功開啟視窗
      console.log('上傳視窗已成功開啟，準備關閉 Modal');
      
      // 儲存視窗參考
      if (!window.openedUploadWindows) {
        window.openedUploadWindows = {};
      }
      window.openedUploadWindows[requestId] = uploadWindow;
      
      // 立即關閉 Modal
      if (modalElement && modalElement.parentNode) {
        modalElement.parentNode.removeChild(modalElement);
        console.log('Modal 已關閉');
        
        // 顯示成功訊息
        UIUtils.displaySuccess('請假表單已提交｜請完成上傳檔案');
      }
      
      // 開始監聽上傳完成事件
      waitForUploadComplete(requestId, null);
      
    } else {
      // 開啟失敗 - 保持 Modal 開啟
      console.log('無法開啟上傳視窗，可能被阻擋');
      
      uploadBtn.disabled = false;
      uploadBtn.innerHTML = `
        <span class="material-icons" style="font-size: 20px; margin-right: 5px;">upload_file</span>
        重試上傳
      `;
      
      // 在按鈕旁顯示錯誤提示
      const errorMsg = document.createElement('div');
      errorMsg.style.cssText = `
        color: #ff4444;
        font-size: 14px;
        margin-top: 10px;
        text-align: center;
      `;
      errorMsg.innerHTML = `
        <span class="material-icons" style="vertical-align: middle; font-size: 18px;">error</span>
        無法開啟視窗，請允許彈出視窗後重試
      `;
      
      // 如果已有錯誤訊息，先移除
      const existingError = uploadBtn.parentNode.querySelector('div[style*="ff4444"]');
      if (existingError) {
        existingError.remove();
      }
      
      uploadBtn.parentNode.appendChild(errorMsg);
    }
  }
  
  /**
   * 等待上傳完成
   * @param {string} requestId - 申請 ID
   * @param {HTMLElement} modalElement - Modal 元素
   */
  async function waitForUploadComplete(requestId, modalElement) {
    // 改為依賴 postMessage 機制，不使用輪詢
    console.log(`等待檔案上傳完成通知... (requestId: ${requestId})`);
    
    // 只保留顯示等待訊息的邏輯
    const uploadStatus = document.getElementById('uploadStatus');
    if (uploadStatus) {
      uploadStatus.innerHTML = `
        <p style="color: #2196F3; margin: 0;">
          <span class="material-icons" style="vertical-align: middle;">cloud_upload</span>
          等待檔案上傳完成...
        </p>
      `;
    }
    
    // 設定超時提醒（5分鐘後）
    setTimeout(() => {
      const currentStatus = document.getElementById('uploadStatus');
      if (currentStatus && modalElement && modalElement.parentNode) {
        currentStatus.innerHTML = `
          <p style="color: #ff9800; margin: 0;">
            <span class="material-icons" style="vertical-align: middle;">warning</span>
            等待逾時，請確認檔案是否已上傳
          </p>
        `;
      }
    }, 300000); // 5分鐘後顯示超時提醒
  }


  /**
   * 創建表單模態框 - 優化版（添加返回按鈕）
   * @param {string} title - 模態框標題
   * @param {HTMLElement} content - 模態框內容
   * @param {Function} submitCallback - 提交回調函數
   */
  function createFormModal(title, content, submitCallback) {
    // 創建模態框元素
    const formOverlay = document.createElement("div");
    formOverlay.className = "overlay show form-overlay";
    formOverlay.id = "dynamicFormModal";

    // 創建模態框內容
    const formModal = document.createElement("div");
    formModal.className = "form-modal";
    
    // 創建模態框標題區域（修改部分）
    const formModalHeader = document.createElement("div");
    formModalHeader.className = "form-modal-header";
    
    // 新增：添加陰影立體風格的返回按鈕
    const backButton = document.createElement("button");
    backButton.className = "back-button-shadow";
    backButton.innerHTML = `
      <span class="material-icons">arrow_back</span>
      <span class="back-text">返回</span>
    `;
    backButton.title = "返回表單選擇";
    
    // 添加標題文字
    const titleText = document.createElement("h3");
    titleText.style.width = "100%";
    titleText.style.textAlign = "center";
    titleText.textContent = title;
    
    // 組合標題區域
    formModalHeader.appendChild(backButton);
    formModalHeader.appendChild(titleText);
    
    // 創建模態框內容區域
    const formModalContent = document.createElement("div");
    formModalContent.className = "form-modal-content";
    formModalContent.appendChild(content);
    
    // 創建滾動到頂部按鈕
    const scrollTopButton = document.createElement("button");
    scrollTopButton.type = "button";
    scrollTopButton.className = "scroll-top-button";
    scrollTopButton.innerHTML = '<span class="material-icons">keyboard_arrow_up</span>';
    scrollTopButton.style.display = "none";
    
    // 創建按鈕區域
    const formButtonsContainer = document.createElement("div");
    formButtonsContainer.className = "form-buttons-container";
    
    // 創建取消按鈕
    const cancelButton = document.createElement("button");
    cancelButton.className = "btn btn-cancel";
    cancelButton.textContent = "取消";
    cancelButton.style.flex = "1";
    cancelButton.style.marginRight = "10px";
    
        // 創建提交按鈕 - 根據表單類型設定不同文字
        const submitButton = document.createElement("button");
        submitButton.className = "btn btn-confirm";
        submitButton.id = "submitFormBtn";
        submitButton.style.flex = "1";
        submitButton.style.marginLeft = "10px";
        
        // 優化：所有表單類型都有專屬按鈕文字
        if (title === "請假申請") {
            submitButton.textContent = "提交申請";
        } else if (title === "加班申請") {
            submitButton.textContent = "提交加班申請";
        } else if (title === "異動班別") {
            submitButton.textContent = "提交異動申請";
        } else if (title === "補卡申請") {
            submitButton.textContent = "提交補卡申請";
        } else {
            submitButton.textContent = "提交表單";
        }
        
        // 動態創建按鈕布局
        createDynamicButtonLayout(formButtonsContainer, cancelButton, submitButton, title);
        
        // 組合模態框
        formModal.appendChild(formModalHeader);
        formModal.appendChild(formModalContent);
        formModal.appendChild(formButtonsContainer);
        formOverlay.appendChild(formModal);
        formOverlay.appendChild(scrollTopButton);
        
        // 添加到 body
        document.body.appendChild(formOverlay);
        
        // 內容滾動事件處理
        formModalContent.addEventListener("scroll", () => {
          if (formModalContent.scrollTop > 100) {
            scrollTopButton.style.display = "flex";
          } else {
            scrollTopButton.style.display = "none";
          }
        });
        
        // 滾動到頂部按鈕點擊事件
        scrollTopButton.addEventListener("click", () => {
          formModalContent.scrollTo({
            top: 0,
            behavior: "smooth"
          });
        });
        
        // 新增：返回按鈕點擊事件
        backButton.addEventListener("click", () => {
          // 顯示自定義確認對話框
          UIUtils.showCustomConfirm(
            "返回確認",
            "即將返回表單首頁｜所有欄位將清空",
            () => {
              // 確認：關閉當前表單模態框
              document.body.removeChild(formOverlay);
              // 重新顯示表單選擇模態框
              showFormApplicationModal();
            },
            () => {
              // 取消：什麼都不做
            }
          );
        });
        
        // 取消按鈕事件
        cancelButton.addEventListener("click", () => {
          // 顯示自定義確認對話框
          UIUtils.showCustomConfirm(
            "關閉確認",
            "即將關閉視窗｜所有欄位將清空",
            () => {
              // 確認：關閉表單模態框
              document.body.removeChild(formOverlay);
            },
            () => {
              // 取消：什麼都不做
            }
          );
        });
        
        // 點擊背景關閉
        formOverlay.addEventListener("click", (e) => {
          if (e.target === formOverlay) {
            // 顯示自定義確認對話框
            UIUtils.showCustomConfirm(
              "離開確認",
              "離開視窗後｜所有欄位將清空",
              () => {
                // 確認：關閉表單模態框
                document.body.removeChild(formOverlay);
              },
              () => {
                // 取消：什麼都不做
              }
            );
          }
        });
        
        // 提交按鈕事件 - 優化UX設計
        submitButton.addEventListener("click", async () => {
          const formContentElement = formModalContent;
          const errorContainer = formContentElement.querySelector("#form-error-container");
          
          // 清除舊錯誤 & 重置樣式
          if (errorContainer) { errorContainer.style.display = 'none'; errorContainer.innerHTML = ''; }
          resetFieldStyles(formContentElement);
          
          // 即時UX反饋 - 按鈕視覺效果
          submitButton.disabled = true;
          submitButton.style.background = "linear-gradient(45deg, #ff6b35, #f7931e)";
          submitButton.style.transform = "scale(0.98)";
          submitButton.style.transition = "all 0.2s ease";
          submitButton.innerHTML = `
            <span class="material-icons" style="animation: spin 1s linear infinite; margin-right: 5px;">refresh</span>
            驗證中...
          `;
          
          // 添加動畫效果
          const style = document.createElement('style');
          style.textContent = `
            @keyframes spin {
              0% { transform: rotate(0deg); }
              100% { transform: rotate(360deg); }
            }
          `;
          document.head.appendChild(style);
            
            const formData = collectFormData(formContentElement);
            let isValid = false;
            
            // 根據表單類型執行對應的驗證函數（使用表單模組）
            if (App.FormHandler) {
              if (title === "補卡申請") {
                isValid = App.FormHandler.validateCardCorrectionFields(formData);
              } else if (title === "加班申請") {
                isValid = App.FormHandler.validateOvertimeFields(formData);
              } else if (title === "請假申請") {
                isValid = App.FormHandler.validateLeaveFields(formData);
              } else if (title === "異動班別") {
                isValid = App.FormHandler.validateShiftChangeFields(formData);
              }
            } else {
              console.error('表單處理模組未載入');
              isValid = false;
            }
            
            if (!isValid) {
              console.log("表單驗證失敗。");
              displayValidationError(errorContainer, formContentElement);
              // 恢復按鈕原始狀態 - 優化所有表單類型
              submitButton.style.background = "";
              submitButton.style.transform = "";
              if (title === "請假申請") {
                submitButton.textContent = "提交申請";
              } else if (title === "加班申請") {
                submitButton.textContent = "提交加班申請";
              } else if (title === "異動班別") {
                submitButton.textContent = "提交異動申請";
              } else if (title === "補卡申請") {
                submitButton.textContent = "提交補卡申請";
              } else {
                submitButton.textContent = "提交表單";
              }
              submitButton.disabled = false;
              return;
            }
            
            // 【新版：移除驗證成功階段，直接進入提交】
            // 驗證通過後直接開始提交，節省 500ms
            submitButton.style.background = "linear-gradient(45deg, #4CAF50, #45a049)";
            submitButton.style.transform = "";
            submitButton.innerHTML = `
              <span class="material-icons" style="animation: spin 1s linear infinite; margin-right: 5px;">sync</span>
              提交中...
            `;
            
            // 驗證通過，執行提交 (調用 submitCallback)
            console.log("驗證通過，開始調用 submitCallback...");
            
            // v2.0 檔案上傳檢查 (請假表單)
            if (title === "請假申請") {
              const requiresFile = App.FormHandler.LEAVE_TYPE_CONFIG && 
                                  App.FormHandler.LEAVE_TYPE_CONFIG.REQUIRES_FILE && 
                                  App.FormHandler.LEAVE_TYPE_CONFIG.REQUIRES_FILE.includes(formData.leaveType);
              
              if (requiresFile) {
                // 加入動畫效果，讓使用者知道正在檢查權限
                submitButton.style.background = "linear-gradient(45deg, #17a2b8, #138496)";
                submitButton.innerHTML = `
                  <span class="material-icons" style="animation: spin 1s linear infinite; margin-right: 5px;">security</span>
                  檢查權限...
                `;
                const hasPermission = await App.UI.BrowserUtils.checkPopupPermission();
                
                // 權限檢查失敗處理
                if (!hasPermission) {
                  const browser = App.UI.BrowserUtils.parseBrowser(navigator.userAgent);
                  let instructions = `您的 ${browser} 瀏覽器可能阻擋了彈出視窗，導致無法上傳檔案。請依照以下建議操作：<br>1. 檢查網址列右側是否有彈出視窗攔截圖示並允許本次操作。<br>2. 暫時停用廣告攔截擴充功能後重試。<br>3. 在瀏覽器設定中，將本站加入彈出視窗的例外清單。`;
                  
                  App.UI.UIUtils.displayError(
                    instructions,
                    { 
                      type: 'PERMISSION',
                      // 重試回呼
                      retryCallback: () => {
                        // 清除錯誤訊息並重新觸發提交點擊
                        App.UI.UIUtils.clearResponse();
                        submitButton.click();
                      }
                    }
                  );
                  
                  submitButton.disabled = false;
                  submitButton.textContent = "提交申請";
                  return; // 中斷流程
                }
              }
            }
            
            // 由動畫處理按鈕文字
            
            // 回調處理
            console.log("submitCallback 類型:", typeof submitCallback);
            console.log("submitCallback 值:", submitCallback);
            if (typeof submitCallback !== 'function') {
              console.error("submitCallback 不是函數！");
              UIUtils.displayError("表單提交功能未正確載入");
              submitButton.disabled = false;
              submitButton.textContent = "提交表單";
              return;
            }
            submitCallback(formData, async (success, message, details) => { 
              console.log("submitCallback 回調:", { success, message, details }); 
              
              if (success) {
                // 【版本三：即時回饋版】
                // 檢查後端是否要求我們保持視窗開啟以進行下一步的檔案上傳
                if (details && details.keepModalOpen && details.needsUpload) {
                  // 需要上傳檔案：簡化為 4 步驟流程
                  const modal = document.getElementById('dynamicFormModal');
                  const formContent = modal.querySelector('.form-modal-content');
                  
                  // 直接插入成功訊息並顯示上傳按鈕
                  const successMsg = document.createElement('div');
                  successMsg.className = 'inline-success-message';
                  successMsg.innerHTML = `
                    <span class="material-icons" style="color: #28a745; vertical-align: middle;">check_circle</span>
                    <span style="color: #28a745; font-weight: 500;">請假申請已成功提交！請完成檔案上傳。</span>
                  `;
                  successMsg.style.cssText = `
                    background: #d4edda;
                    color: #155724;
                    padding: 12px 16px;
                    border-radius: 8px;
                    margin-bottom: 20px;
                    animation: slideIn 0.3s ease-out;
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                  `;
                  formContent.insertBefore(successMsg, formContent.firstChild);
                  
                  // 立即將按鈕轉換為上傳按鈕（UI 控制權交給 showUploadPrompt）
                  // showUploadPrompt 會清空按鈕容器並建立新按鈕
                  
                  return;
                }
                
                // 如果不是需要上傳檔案的特殊情況，才執行正常的關閉視窗、顯示成功訊息等操作
                const modal = document.getElementById('dynamicFormModal');
                if (modal) {
                  document.body.removeChild(modal);
                }
                UIUtils.displaySuccess(message || '操作成功');

              } else { // 處理失敗
                // 授權錯誤檢查
                if (details && details.stage === 'auth') {
                  // 先關閉表單視窗
                  const modal = document.getElementById('dynamicFormModal');
                  if (modal) document.body.removeChild(modal);
                  
                  // 顯示明確的錯誤訊息
                  if (message && message.includes("AUTH.NOT_IN_AUTHLIST")) {
                    UIUtils.displayError("不在授權名單內");
                  } else if (message && message.includes("AUTH.DEVICE_MISMATCH")) {
                    UIUtils.displayError("裝置不匹配");
                  } else if (message && message.includes("AUTH.NOT_BOUND")) {
                    UIUtils.displayError("裝置未綁定");
                  } else {
                    UIUtils.displayError("不在授權名單內");
                  }
                } else {
                  // 其他錯誤保持原有處理方式
                  UIUtils.displayError(message || "提交失敗｜請稍後再試");
                  // 恢復按鈕文字 - 所有表單類型
                  if (title === "請假申請") {
                    submitButton.textContent = "提交申請";
                  } else if (title === "加班申請") {
                    submitButton.textContent = "提交加班申請";
                  } else if (title === "異動班別") {
                    submitButton.textContent = "提交異動申請";
                  } else if (title === "補卡申請") {
                    submitButton.textContent = "提交補卡申請";
                  } else {
                    submitButton.textContent = "提交表單";
                  }
                  submitButton.disabled = false;
                }
              }
            });
        });
      }
    
      /**
       * 收集表單數據
       * @param {HTMLElement} formContent - 表單內容元素
       * @returns {Object} 收集的表單數據
       */
      function collectFormData(formContent) {
        const formData = {};
        const inputs = formContent.querySelectorAll("input, select, textarea");
        
        inputs.forEach(input => {
          if (input.name) {
            formData[input.name] = input.value;
          }
        });
        
        return formData;
      }
    
      /**
       * 顯示驗證錯誤
       * @param {HTMLElement} errorContainer - 錯誤容器
       * @param {HTMLElement} formContainer - 表單容器
       */
      function displayValidationError(errorContainer, formContainer) {
        if (errorContainer) {
          errorContainer.textContent = "請填寫所有必填欄位";
          errorContainer.style.display = "block";
        }
        
        // 滾動到頂部顯示錯誤
        if (formContainer) {
          formContainer.scrollTo({
            top: 0,
            behavior: "smooth"
          });
        }
      }
    
      /**
       * 重置所有欄位樣式
       * @param {HTMLElement} formContainer - 表單容器
       */
      function resetFieldStyles(formContainer) {
        if (!formContainer) return;
        
        // 重設所有輸入欄位的樣式
        formContainer.querySelectorAll(".form-input, .form-textarea, select").forEach(el => {
          if (el.tagName) {
            el.style.borderColor = "#555";
          }
        });
        
        // 隱藏所有必填標記
        formContainer.querySelectorAll(".required-mark").forEach(mark => {
          mark.style.display = "none";
        });
      }
    
      /**
       * 處理API錯誤
       * @param {Error} error - 錯誤對象
       * @param {HTMLElement} errorContainer - 錯誤容器
       * @param {Function} callback - 回調函數
       */
      function handleProcessError(error, errorContainer, callback) {
        console.error("處理錯誤:", error);
        
        if (errorContainer) {
          errorContainer.textContent = error.message || "處理失敗，請稍後再試";
          errorContainer.style.display = "block";
        }
        callback(false, error.message || "處理失敗，請稍後再試");
      }
    
      // 表單相關函式已遷移至 form-handler.ui.js

  // 匯出模組

  // 靜態配置
  function fetchAndApplyConfig() {
    console.log("✓ 使用靜態配置模式");
    return Promise.resolve(true);
  }
  
  // 匯出到 App.UI
  App.UI = {
      // 配置
      config: {},
      fetchAndApplyConfig,
      getConfig: function(key) {
        return App.config ? App.config[key] : undefined;
      },

      // 模組
      ButtonStateManager,
      BrowserUtils,
      NetworkUtils,
      UIUtils,
      DeviceIdManager,
      LocationManager,  // ← 導出多點定位取樣功能
      TimeZoneUtils,
      DOMPurifyUtils,
      
      // 函式
      debounce,
      validateIdCode,
      initPickerStyles,
      showFormApplicationModal,
      hideFormApplicationModal,
      showVideoTutorialModal,
      hideVideoTutorialModal,
      openSchedule,
      openNotice,
      testLocation,
      showQueryMenu,
      showAttendanceQuery,
      showFormTypeSelector,
      executeFormQuery,
      closeQueryTypeModal,
      closeFormTypeModal,
      startClockAction,
      enhancedLocationData,  // ← 恢復打卡功能的關鍵函數導出
      showCardCorrectionForm,
      showOvertimeForm,
      showLeaveRequestForm,
      showShiftChangeForm,
      initializeFormIntegration,

      // DOM elements
      elements: elements
  };

  // 採用靜態配置策略，移除複雜的配置橋接機制
  // 確保系統載入穩定性，借鑑原始版本成功模式

  // 在模組載入完成後初始化表單整合
  document.addEventListener('DOMContentLoaded', () => {
    initializeFormIntegration();
  });

})(window.App);
