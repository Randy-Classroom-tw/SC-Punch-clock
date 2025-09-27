// 確保 App 總部已建立
window.App = window.App || {};

// 全局UI配置 (IIFE模組)
(function(App) {
  'use strict';

  // 預設配置
  const DEFAULT_CONFIG = {
    WEB_APP_URL: "https://script.google.com/macros/s/AKfycbzEbMJPxjth4jDYm1gjjbZdbF2Mqhgg2BmsBS1VSMCAgnVA6b_YN00514rYZ8jv4NWr/exec",
    COOLDOWN_MS: 5000,
    API_TIMEOUT: 30000,
    DEVICE_ID_KEY: 'OS_DEVICE_ID',
    APP_VERSION: '19.9.8',
    SCHEDULE_URL: "https://docs.google.com/spreadsheets/d/1IlbSMjDwpv5YnJDwP6ctV-1iJiv6Rzrp--HcTKZUcWo/edit?usp=drive_link",  // 排班表
    NOTICE_URL: "https://docs.google.com/document/d/1pASDU2UOVH3U9Wm4KfnZOMac5wI43XYHJ9uxtriwBCQ/edit?usp=sharing",    // 公告
    CONSISTENCY_CHECK_INTERVAL: 30000,
    
    // 統一日期時間選擇器配置
    PICKER_CONFIG: {
      // 日期選擇器配置
      datepicker: {
        dateFormat: 'yy-mm-dd',
        showButtonPanel: true,
        changeMonth: true,
        changeYear: true,
        yearRange: '2020:2030',
        minDate: new Date(2020, 0, 1),
        maxDate: new Date(2030, 11, 31),
        dayNamesMin: ['日', '一', '二', '三', '四', '五', '六'],
        monthNamesShort: ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'],
        currentText: '今天',
        closeText: '關閉',
        prevText: '上月',
        nextText: '下月',
        beforeShow: function(input, inst) {
          // 空函數，防止定位問題
        }
      },
      
      // 時間選擇器配置
      timepicker: {
        timeFormat: 'HH:mm',
        controlType: 'select',
        showSecond: false,
        hourGrid: 6,
        minuteGrid: 15,
        showButtonPanel: true,
        timeOnlyTitle: '選擇時間',
        timeText: '時間',
        hourText: '時',
        minuteText: '分',
        currentText: '現在時間',
        closeText: '關閉',
        
        // 自訂按鈕處理
        beforeShow: function(input, inst) {
          setTimeout(function() {
            const dpDiv = $('#ui-datepicker-div');
            const buttonPane = dpDiv.find('.ui-datepicker-buttonpane');
            
            buttonPane.empty();
            
            // 建立自訂按鈕容器
            const buttonContainer = $('<div>')
              .css({
                'display': 'flex',
                'gap': '10px',
                'justify-content': 'center',
                'padding': '10px'
              });
            
            // 建立取消按鈕
            const cancelBtn = $('<button>')
              .text('取消')
              .attr('type', 'button')
              .css({
                'background-color': '#e74c3c',
                'color': 'white',
                'border': 'none',
                'padding': '5px 15px',
                'border-radius': '4px',
                'cursor': 'pointer',
                'font-size': '14px'
              })
              .on('click', function(e) {
                e.preventDefault();
                e.stopPropagation();
                $(input).timepicker('hide');
              });
            
            // 建立確認按鈕
            const confirmBtn = $('<button>')
              .text('確認')
              .attr('type', 'button')
              .css({
                'background-color': '#28a745',
                'color': 'white',
                'border': 'none',
                'padding': '5px 15px',
                'border-radius': '4px',
                'cursor': 'pointer',
                'font-size': '14px'
              })
              .on('click', function(e) {
                e.preventDefault();
                e.stopPropagation();
                const hours = dpDiv.find('.ui_tpicker_hour_slider').find('select').val() || '00';
                const minutes = dpDiv.find('.ui_tpicker_minute_slider').find('select').val() || '00';
                const formattedTime = `${hours.padStart(2, '0')}:${minutes.padStart(2, '0')}`;
                $(input).val(formattedTime);
                $(input).timepicker('hide');
              });
            
            buttonContainer.append(cancelBtn).append(confirmBtn);
            buttonPane.append(buttonContainer);
          }, 10);
        }
      },
      
      // 視覺樣式配置
      styles: {
        datepicker: {
          width: 320,
          minWidth: 280,
          maxWidth: 360,
          background: '#222',
          borderRadius: '6px',
          boxShadow: '0 4px 12px rgba(0,0,0,0.5)'
        },
        timepicker: {
          width: 300,
          minWidth: 300,
          maxWidth: 340,
          background: '#222',
          padding: '15px',
          borderTop: '1px solid #444'
        }
      }
    }
  };

  // 等待資源載入
  window.addEventListener('load', async () => {

    const UI = App.UI;

    // 檢查UI模組
    if (!UI) {
      console.error("錯誤：UI 模組 (ui.js) 未能成功載入。");
      alert("系統初始化失敗｜請點擊 AI Bot");
      return;
    }

    // 主要初始化
    async function main() {
      console.log("=== 企業級出勤系統 v2.0 啟動 ===");
      
      App.config = { ...DEFAULT_CONFIG };
      console.log("✓ 已載入預設配置");
      
      UI.config = App.config;
      
      console.log("系統配置已載入:", UI.config);
      
      setupEventListeners();
      console.log("✓ 事件監聽器設定完成");
      
      UI.DeviceIdManager.recoverDeviceId();
      console.log("✓ 設備ID管理器初始化完成");
      
      UI.BrowserUtils.showSafariNotice();
      UI.NetworkUtils.updateNetworkStatus();
      setInterval(UI.DeviceIdManager.validateStoredIds, UI.getConfig('CONSISTENCY_CHECK_INTERVAL'));
      console.log("✓ 瀏覽器相關功能初始化完成");

      addCustomStyles();
      addFormNoScrollStyles();
      UI.initPickerStyles();
      initSvgOptions();
      addFormBackButtonStyles();
      console.log("✓ UI樣式初始化完成");
      
      console.log("=== 系統啟動完成 ===");
    }


  // 事件監聽與初始化

  // 按鈕冷卻時間
  const COOLDOWN_TIMES = {
    default: 5000,
    btnOn: 3000,
    btnOff: 3000,
    btnTest: 3000,
    btnSchedule: 3000,
    btnNotice: 3000,
    btnQuery: 5000,
    btnForm: 5000,
    btnManager: 3000,
    btnPersonal: 3000
  };

  // 設置事件監聽器
  function setupEventListeners() {
    // 切換密碼顯示
    UI.elements.toggleBtn.addEventListener("click", () => {
        UI.elements.idInput.type = (UI.elements.idInput.type === "password") ? "text" : "password";
        UI.elements.eyeIcon.textContent = (UI.elements.idInput.type === "password") ? "visibility" : "visibility_off";
        UI.elements.idInput.focus();
    });

    // 輸入驗證
    UI.elements.idInput.addEventListener("input", () => {
        if (!/^\d*$/.test(UI.elements.idInput.value)) {
          UI.elements.idInput.value = UI.elements.idInput.value.replace(/\D/g, "");
          UI.elements.errorSpan.textContent = "只允許輸入 4 位數字";
          UI.elements.errorSpan.classList.add("show");
      } else {
          UI.elements.errorSpan.textContent = "";
          UI.elements.errorSpan.classList.remove("show");
      }
        UI.elements.idInput.setCustomValidity(UI.elements.idInput.value.length !== 4 ? "必須輸入 4 位數" : "");
        if (UI.validateIdCode(UI.elements.idInput.value)) {
          UI.UIUtils.enableAllButtons();
      } else {
          UI.UIUtils.disableAllButtons();
      }
    });
      UI.elements.idInput.addEventListener("blur", () => {
        UI.elements.errorSpan.textContent = "";
        UI.elements.errorSpan.classList.remove("show");
        UI.elements.idInput.setCustomValidity(UI.elements.idInput.value.length !== 4 ? "必須輸入 4 位數" : "");
    });

    // 影片教學
    UI.elements.closeVideoModal.addEventListener("click", UI.hideVideoTutorialModal);
      UI.elements.tutorialVideo.addEventListener("error", () => {
        UI.elements.tutorialVideo.style.display = "none";
        UI.elements.videoFallback.style.display = "block";
    });

    // 功能按鈕
    UI.elements.allButtons.btnOn.addEventListener("click", () => {
        UI.startClockAction("上班", false, UI.elements.allButtons.btnOn);
    });
      UI.elements.allButtons.btnOff.addEventListener("click", () => {
        UI.startClockAction("下班", false, UI.elements.allButtons.btnOff);
    });

    // 防抖定位 & 查詢
    const debouncedTestLocation = UI.debounce(UI.testLocation, 300);
      UI.elements.allButtons.btnTest.addEventListener("click", () => {
        debouncedTestLocation(UI.elements.allButtons.btnTest);
    });

      const debouncedQueryRecords = UI.debounce(UI.showQueryMenu, 300);
      UI.elements.allButtons.btnQuery.addEventListener("click", () => {
        debouncedQueryRecords(UI.elements.allButtons.btnQuery);
    });

    // 班表按鈕
    UI.elements.allButtons.btnSchedule.addEventListener("click", () => {
        UI.openSchedule(UI.elements.allButtons.btnSchedule);
    });

    // 內部公告
    UI.elements.allButtons.btnNotice.addEventListener("click", () => {
        UI.openNotice(UI.elements.allButtons.btnNotice);
    });

    // 表單按鈕
    UI.elements.allButtons.btnForm.addEventListener("click", () => {
      UI.showFormApplicationModal();
    });

    // 主管限定 (禁用)
    UI.elements.allButtons.btnManager.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      return false;
    });

    // Duty日誌 (禁用)
    UI.elements.allButtons.btnPersonal.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      return false;
    });

    // 點擊背景關閉表單
    const formAppModal = document.getElementById("formApplicationModal");
    formAppModal.addEventListener("click", (e) => {
      if (e.target === formAppModal) {
        UI.UIUtils.showCustomConfirm(
          "離開確認",
          "即將關閉表單選擇｜返回主畫遢",
          () => {
            UI.hideFormApplicationModal();
          },
          () => {}
        );
      }
    });

    // 表單項目
    document.getElementById("cardCorrectionForm").addEventListener("click", () => {
        UI.hideFormApplicationModal();
        UI.showCardCorrectionForm();
    });

    document.getElementById("overtimeForm").addEventListener("click", () => {
        UI.hideFormApplicationModal();
        UI.showOvertimeForm();
    });

    document.getElementById("leaveRequestForm").addEventListener("click", (event) => {
        UI.showLeaveRequestForm();
    });

    document.getElementById("shiftChangeForm").addEventListener("click", (event) => {
        // 維修中
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        return false;
    });

      window.addEventListener("online", UI.NetworkUtils.updateNetworkStatus);
      window.addEventListener("offline", UI.NetworkUtils.updateNetworkStatus);
  }

  // 返回按鈕樣式
   function addFormBackButtonStyles() {
    const styleElement = document.createElement('style');
    styleElement.textContent = `
      /* 陰影立體風格按鈕樣式 */
      .back-button-shadow {
        position: absolute;
        left: 15px;
        background-color: #3a3a3a;
        border: none;
        color: white;
        padding: 6px 12px;
        display: flex;
        align-items: center;
        border-radius: 8px;
        cursor: pointer;
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.3), 
                    inset 0 1px 0 rgba(255, 255, 255, 0.15);
        transition: all 0.2s;
      }
      
      .back-button-shadow .material-icons {
        font-size: 20px;
        margin-right: 6px;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.4);
      }
      
      .back-button-shadow .back-text {
        font-size: 14px;
        font-weight: 500;
        text-shadow: 0 1px 2px rgba(0, 0, 0, 0.3);
      }
      
      .back-button-shadow:hover {
        background-color: #454545;
        transform: translateY(-1px);
        box-shadow: 0 3px 7px rgba(0, 0, 0, 0.4), 
                    inset 0 1px 0 rgba(255, 255, 255, 0.2);
      }
      
      .back-button-shadow:active {
        background-color: #333333;
        transform: translateY(1px);
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.3), 
                    inset 0 1px 2px rgba(0, 0, 0, 0.3);
      }
      
      /* 響應式設計 */
      @media (max-width: 480px) {
        .back-button-shadow {
          padding: 4px 10px;
        }
        
        .back-button-shadow .material-icons {
          font-size: 18px;
          margin-right: 4px;
        }
        
        .back-button-shadow .back-text {
          font-size: 12px;
        }
      }
    `;
    document.head.appendChild(styleElement);
  }

  // SVG樣式
  function addSvgStyles() {
    $("<style>")
      .prop("type", "text/css")
      .html(`
        /* SVG選項樣式 */
        .svg-option {
          overflow: hidden;
        }
        
        .svg-option svg {
          display: block;
        }
        
        .svg-option svg text {
          font-family: 'Roboto Mono', monospace;
        }
      `)
      .appendTo("head");
  }

  // 初始化SVG
  function initSvgOptions() {
    addSvgStyles();
  }

  // 全局CSS樣式
  function addCustomStyles() {
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

    $("<style>")
      .prop("type", "text/css")
      .html(`
        /* Datetimepicker 樣式 */
        .ui-datepicker {
            width: ${datepickerConfig.width || 320}px;
            min-width: ${datepickerConfig.minWidth || 280}px;
            max-width: ${datepickerConfig.maxWidth || 360}px;
          background: #333;
          border: 1px solid #555;
            border-radius: ${datepickerConfig.borderRadius || '6px'};
            box-shadow: ${datepickerConfig.boxShadow || '0 4px 12px rgba(0,0,0,0.5)'};
          padding: 10px;
          box-sizing: border-box;
        }
        
        #ui-datepicker-div {
          z-index: 10000;
            background-color: ${datepickerConfig.background || '#222'};
          padding: 0;
          border: none;
          margin-top: 0;
            box-shadow: ${datepickerConfig.boxShadow || '0 4px 12px rgba(0,0,0,0.5)'};
        }
        
        /* 時間選擇器 */
        .ui-timepicker-div {
            width: ${timepickerConfig.width || 300}px;
            min-width: ${timepickerConfig.minWidth || 300}px;
            max-width: ${timepickerConfig.maxWidth || 340}px;
            background-color: ${timepickerConfig.background || '#222'};
            padding: ${timepickerConfig.padding || '15px'};
          border-radius: 0;
          border: none;
            border-top: ${timepickerConfig.borderTop || '1px solid #444'};
          box-shadow: none;
        }

        /* 🔧 修復：完全隱藏 jQuery UI Timepicker 的 grid 顯示元素 */
        .ui-timepicker-grid,
        .ui-timepicker-div .ui-timepicker-hour-label,
        .ui-timepicker-div .ui-timepicker-minute-label,
        .ui-timepicker-div .ui-widget-content,
        .ui-timepicker-div table.ui-timepicker-table,
        .ui-timepicker-div [class*="grid"],
        .ui-timepicker-div [class*="table"] {
          display: none !important;
          visibility: hidden !important;
        }
        
        .ui-datepicker-buttonpane,
        .ui-timepicker-buttonpane {
          background: transparent;
          border: none;
        }
        
        .form-input,
        input.datepicker,
        input.timepicker {
          text-align: left;
          padding-left: 15px;
        }
        
        .form-input, .weekday-display, .date-input-container, .time-input-container, .radio-option {
          height: 42px;
          box-sizing: border-box;
        }
        
        .weekday-display {
          width: 60px;
          display: flex;
          justify-content: center;
          align-items: center;
        }
        
        .ui-timepicker-div .ui_tpicker_time,
        .ui-timepicker-div .ui_tpicker_time_label,
        .ui-timepicker-div .ui_tpicker_hour_label,
        .ui-timepicker-div .ui_tpicker_minute_label {
          display: none;
        }
        
        .ui-slider {
          display: none;
        }
        
        .ui-timepicker-div select {
          width: 95%;
          margin: 0 auto;
          text-align: center;
          text-align-last: center;
          -moz-text-align-last: center;
          -webkit-text-align-last: center;
        }
        
        .ui-timepicker-div select option {
          text-align: center;
        }
      `)
      .appendTo("head");
  }

  // 禁止表單滑動樣式
   function addFormNoScrollStyles() {
    $("<style>")
      .prop("type", "text/css")
      .html(`
        .form-content {
          overflow-x: hidden !important;
        }
        
        .form-modal {
          overflow-x: hidden !important;
          max-width: 450px !important;
          width: 90% !important;
        }
        
        .form-group,
        .form-input,
        .form-textarea,
        .radio-container,
        .form-error-container {
          max-width: 100% !important;
          box-sizing: border-box !important;
        }
        
        .date-input-container {
          position: relative !important;
          width: calc(100% - 90px) !important;
          max-width: calc(100% - 90px) !important;
          box-sizing: border-box !important;
          margin-right: 90px !important;
        }
        
        .weekday-display {
          position: absolute !important;
          right: -90px !important;
          top: 0 !important;
          height: 42px !important;
          width: 80px !important;
          display: flex !important;
          justify-content: center !important;
          align-items: center !important;
          visibility: visible !important;
          z-index: 10 !important;
        }
        
        .time-input-container {
          max-width: 100% !important;
          box-sizing: border-box !important;
        }
        
        .form-error-container {
          text-align: center !important;
        }
        
        @media (max-width: 480px) {
          .form-modal {
            width: 95% !important;
            max-width: 95% !important;
          }
          
          .weekday-display {
            width: 70px !important;
            right: -80px !important;
            font-size: 14px !important;
          }
        }
      `)
      .appendTo("head");
  }

  // postMessage 監聽器
  if (!window.hasAddedCloseWindowListener) { 
    window.addEventListener('message', function(event) {
      console.log('[主頁面] 收到 Message:', '來源:', event.origin, '資料:', event.data); 
      
      try {
        // 來源驗證
        const webAppUrl = App.config.WEB_APP_URL;  // 直接使用 App.config
        if (!webAppUrl) {
          console.warn('[主頁面] 無法取得有效的 WEB_APP_URL，嘗試使用降級方案');
          // 降級：接受來自 script.google.com 的訊息
          if (event.origin.includes('script.google.com')) {
            console.log('[主頁面] 使用降級方案：接受來自 script.google.com 的訊息');
          } else {
            return;
          }
        } else {
          const expectedOrigin = new URL(webAppUrl).origin;
          if (event.origin !== expectedOrigin) {
            console.log('[主頁面] 忽略的訊息：來源不符。預期:', expectedOrigin, '實際:', event.origin);
            return; 
          }
        }
      } catch (error) {
        console.error('[主頁面] postMessage 處理錯誤:', error);
        return;
      }
      
      if (event.data && event.data.action === 'closeUploadWindow' && event.data.requestId) {
        const requestIdToClose = String(event.data.requestId);
        console.log(`[主頁面] 收到有效關閉請求，RequestId: ${requestIdToClose}`);
         
        const targetWindow = window.openedUploadWindows[requestIdToClose];
        console.log(`[主頁面] 根據 RequestId 查找到的視窗參照:`, targetWindow); 
         
        if (targetWindow) {
          if (!targetWindow.closed) {
            console.log(`[主頁面] 找到視窗且未關閉，準備執行 close()`);
            try {
              targetWindow.close();
              console.log(`[主頁面] targetWindow.close() 已呼叫`); 
            } catch (e) {
              console.error(`[主頁面] 執行 targetWindow.close() 時出錯:`, e); 
            }
            
            setTimeout(() => {
                delete window.openedUploadWindows[requestIdToClose]; 
                console.log(`[主頁面] 已移除追蹤，RequestId: ${requestIdToClose}`);
            }, 500);
          } else {
            console.log(`[主頁面] 視窗先前已關閉`);
            delete window.openedUploadWindows[requestIdToClose]; 
          }
        } else {
           console.warn(`[主頁面] 找不到要關閉的視窗參照。`);
        }
        
        // 關閉表單 Modal
        const formModal = document.getElementById('dynamicFormModal');
        if (formModal) {
          console.log('[主頁面] 檔案上傳完成，關閉表單 Modal');
          document.body.removeChild(formModal);
          
          // 顯示成功訊息
          if (App.UI && App.UI.UIUtils) {
            App.UI.UIUtils.displaySuccess('檔案上傳成功｜請假申請已完成');
          }
        }
      }
    }, false);
    window.hasAddedCloseWindowListener = true; 
    console.log("已添加關閉視窗的 Message 事件監聽器。");
  }

    main();

  });

})(window.App);
