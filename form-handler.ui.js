// 表單處理模組
window.App = window.App || {};

// IIFE模組
(function(App) {
  'use strict';

  // 表單處理模組 - 四種表單類型：補卡、加班、請假、班別異動

  // jQuery檢查
  if (typeof $ === 'undefined' || typeof $.fn === 'undefined') {
    console.warn('jQuery 未載入');
  }

  // 表單配置

  // UI配置在 App.UI_CONFIG

  // 統一配置存取函數
  function getPickerStyles() {
    return App.config?.PICKER_CONFIG?.styles || {
      timepicker: {
        width: 300,
        minWidth: 300,
        maxWidth: 340,
        background: '#222',
        padding: '15px',
        borderTop: '1px solid #444'
      },
      datepicker: {
        width: 320,
        minWidth: 280,
        maxWidth: 360,
        background: '#222',
        borderRadius: '6px',
        boxShadow: '0 4px 12px rgba(0,0,0,0.5)'
      }
    };
  }

  // 表單工具函式

  // 加班時數計算
  function calculateEstimatedOvertimeHours(startTime, endTime, forcesCrossDay = false) {
    if (!startTime || !endTime || !/^\d{2}:\d{2}$/.test(startTime) || !/^\d{2}:\d{2}$/.test(endTime)) return null;
    try {
      const [startH, startM] = startTime.split(':').map(Number);
      const [endH, endM] = endTime.split(':').map(Number);
      if (isNaN(startH) || isNaN(startM) || isNaN(endH) || isNaN(endM)) {
        return null;
      }
      let startMinutes = startH * 60 + startM;
      let endMinutes = endH * 60 + endM;
      
      // 如果強制跨日或結束時間小於開始時間，加24小時
      if (forcesCrossDay || endMinutes < startMinutes) {
        endMinutes += 24 * 60;
      }
      
      const durationMinutes = endMinutes - startMinutes;
      if (durationMinutes <= 0) return 0;
      
      const hours = Math.floor(durationMinutes / 60);
      const remainingMinutes = durationMinutes % 60;
      
      let roundedHours = hours;
      if (remainingMinutes >= 45) {
        roundedHours += 1.0;
      } else if (remainingMinutes >= 15) {
        roundedHours += 0.5;
      }
      
      return roundedHours;
    } catch (e) {
      console.error("計算加班時數錯誤:", e);
      return null;
    }
  }

  // 計算請假天數
  function calculateEstimatedLeaveDays(startDate, endDate) {
    if (!startDate || !endDate || !/^\d{4}-\d{2}-\d{2}$/.test(startDate) || !/^\d{4}-\d{2}-\d{2}$/.test(endDate)) return null;
    try {
      const startParts = startDate.split('-').map(Number);
      const endParts = endDate.split('-').map(Number);
      const start = new Date(startParts[0], startParts[1] - 1, startParts[2]);
      const end = new Date(endParts[0], endParts[1] - 1, endParts[2]);
      if (isNaN(start.getTime()) || isNaN(end.getTime()) || start > end) {
        return null;
      }
      const oneDay = 24 * 60 * 60 * 1000;
      const startUTC = Date.UTC(start.getFullYear(), start.getMonth(), start.getDate());
      const endUTC = Date.UTC(end.getFullYear(), end.getMonth(), end.getDate());
      const diffDays = Math.round(Math.abs(endUTC - startUTC) / oneDay) + 1;
      return diffDays;
    } catch (e) {
      console.error("計算請假天數錯誤:", e);
      return null;
    }
  }

  // 表單驗證函式

  // 驗證欄位
  function validateField(fieldId) {
    const field = document.getElementById(fieldId);
    if (!field) return false;
    let isValid = true;
    if (field.type === "file") {
      isValid = field.files && field.files.length > 0;
    } else if (field.tagName === "SELECT") {
      isValid = !!field.value;
    } else if (field.tagName === "TEXTAREA" || field.type === "text" || field.type === "password" || field.classList.contains("datepicker") || field.classList.contains("timepicker")) {
      isValid = !!field.value.trim();
    } else if (field.type === "hidden") {
      isValid = !!field.value;
    }
    if (!isValid) {
      showFieldError(fieldId);
    }
    return isValid;
  }

  // 顯示錯誤提示
  function showFieldError(fieldId, errorMessage = null) {
    const field = document.getElementById(fieldId);
    if (!field) return;
    
    const formGroup = field.closest('.form-group');
    if (!formGroup) return;
    
    // 添加錯誤類別
    formGroup.classList.add('has-error');
    
    // 顯示必填標記
    const errorMark = formGroup.querySelector('.required-mark');
    if (errorMark) {
      errorMark.style.display = 'inline';
    }
    
    // 如果有錯誤訊息，顯示行內錯誤
    if (errorMessage) {
      let errorInline = formGroup.querySelector('.error-inline');
      if (!errorInline) {
        errorInline = document.createElement('div');
        errorInline.className = 'error-inline';
        errorInline.innerHTML = `
          <span class="material-icons" style="font-size: 14px;">error</span>
          <span class="error-text"></span>
        `;
        formGroup.appendChild(errorInline);
      }
      errorInline.querySelector('.error-text').textContent = errorMessage;
    }
    
    // 設定邊框顏色
    if (field.tagName === "SELECT" || field.tagName === "TEXTAREA" || field.tagName === "INPUT" || field.type === "text" || field.classList.contains("datepicker") || field.classList.contains("timepicker")) {
      field.style.borderColor = "#ff4444";
    }
  }

  // 清除錯誤提示
  function clearFieldError(fieldId) {
    const field = document.getElementById(fieldId);
    if (!field) return;
    
    const formGroup = field.closest('.form-group');
    if (!formGroup) return;
    
    // 移除錯誤類別
    formGroup.classList.remove('has-error');
    
    // 隱藏錯誤訊息
    const errorInline = formGroup.querySelector('.error-inline');
    if (errorInline) {
      errorInline.style.display = 'none';
    }
    
    // 隱藏必填標記
    const errorMark = formGroup.querySelector('.required-mark');
    if (errorMark) {
      errorMark.style.display = 'none';
    }
    
    // 重設邊框顏色
    if (field.tagName === "SELECT" || field.tagName === "TEXTAREA" || field.tagName === "INPUT" || field.type === "text" || field.classList.contains("datepicker") || field.classList.contains("timepicker")) {
      field.style.borderColor = "#555";
      field.style.backgroundColor = "";
    }
    
    const formErrorContainer = document.getElementById("form-error-container");
    if (formErrorContainer) {
      formErrorContainer.style.display = "none";
    }
  }

  // 添加驗證事件
  function addValidationEvents(element) {
    if (!element || !element.id) return;
    const fieldId = element.id;
    element.addEventListener("blur", () => {
      let isEmpty = false;
      if (element.type === "file") {
        isEmpty = !element.files || element.files.length === 0;
      } else if (element.tagName === "SELECT") {
        isEmpty = !element.value;
      } else {
        isEmpty = !element.value.trim();
      }
      if (isEmpty) {
        showFieldError(fieldId);
      } else {
        clearFieldError(fieldId);
      }
    });
    const inputHandler = () => {
      let hasValue = false;
      if (element.type === "file") {
        hasValue = element.files && element.files.length > 0;
      } else if (element.tagName === "SELECT") {
        hasValue = !!element.value;
      } else {
        hasValue = !!element.value.trim();
      }
      if (hasValue) {
        clearFieldError(fieldId);
      }
    };
    if (element.tagName === "SELECT" || element.type === "file") {
      element.addEventListener("change", inputHandler);
    } else {
      element.addEventListener("input", inputHandler);
    }
    if (element.classList && element.classList.contains && (element.classList.contains("datepicker") || element.classList.contains("timepicker"))) {
      if (typeof $ !== 'undefined') {
        $(element).on("change", function() {
          if (this.value) clearFieldError(fieldId);
        });
      }
    }
  }

  // 按類型驗證

  // 驗證補卡表單
  function validateCardCorrectionFields(formData) {
    let isValid = true;
    if (!validateField("cardDate")) isValid = false;
    if (!validateField("cardTime")) isValid = false;
    if (!validateField("cardAction")) isValid = false;
    if (!validateField("cardType")) isValid = false;
    if (!validateField("cardReason")) isValid = false;
    if (!validateField("cardWitness")) isValid = false;
    if (!validateField("cardDuty")) isValid = false;
    return isValid;
  }

  // 驗證加班表單
  function validateOvertimeFields(formData) {
    let isValid = true;
    if (!validateField("overtimeDate")) isValid = false;
    if (!validateField("overtimeType")) isValid = false;
    if (!validateField("overtimePeriod")) isValid = false;
    if (!validateField("overtimeStartTime")) isValid = false;
    if (!validateField("overtimeEndTime")) isValid = false;
    if (!validateField("overtimeReason")) isValid = false;
    if (!validateField("overtimeDuty")) isValid = false;
    const hours = calculateEstimatedOvertimeHours(formData.overtimeStartTime, formData.overtimeEndTime);
    if (hours === null || hours <= 0) {
      if (formData.overtimeStartTime && formData.overtimeEndTime) {
        isValid = false;
        const errorContainer = document.getElementById("form-error-container");
        if (errorContainer) {
          errorContainer.textContent = "結束時間必須晚於開始時間";
          errorContainer.style.display = "block";
        }
        showFieldError("overtimeStartTime");
        showFieldError("overtimeEndTime");
      }
    }
    return isValid;
  }

  // 驗證請假表單
  function validateLeaveFields(formData) {
    let isValid = true;
    if (!validateField("leaveStartDate")) isValid = false;
    if (!validateField("leaveEndDate")) isValid = false;
    if (!validateField("leaveStartTime")) isValid = false;
    if (!validateField("leaveEndTime")) isValid = false;
    if (!validateField("leaveType")) isValid = false;
    if (!validateField("leaveReason")) isValid = false;
    if (!validateField("leaveSubstituteStatus")) isValid = false;
    if (formData.leaveSubstituteStatus === "有找到") {
      if (!validateField("leaveSubstituteName")) isValid = false;
    }
    if (!validateField("leaveAuthDuty")) isValid = false;
    
    // 驗證日期時間邏輯
    if (formData.leaveStartDate && formData.leaveEndDate) {
      const startDate = new Date(formData.leaveStartDate);
      const endDate = new Date(formData.leaveEndDate);
      
      // 清除之前的錯誤
      clearFieldError("leaveEndDate");
      clearFieldError("leaveEndTime");
      
      // 檢查結束日期是否早於開始日期
      if (endDate < startDate) {
        isValid = false;
        showFieldError("leaveEndDate", "結束日期不能早於開始日期");
      } else if (startDate.getTime() === endDate.getTime()) {
        // 同一天的情況，檢查時間
        if (formData.leaveStartTime && formData.leaveEndTime) {
          const [startHour, startMin] = formData.leaveStartTime.split(':').map(Number);
          const [endHour, endMin] = formData.leaveEndTime.split(':').map(Number);
          const startMinutes = startHour * 60 + startMin;
          const endMinutes = endHour * 60 + endMin;
          
          if (endMinutes <= startMinutes) {
            isValid = false;
            showFieldError("leaveEndTime", "同一天內結束時間必須晚於開始時間");
          }
        }
      }
    }
    return isValid;
  }

  // 驗證班別異動表單
  function validateShiftChangeFields(formData) {
    let isValid = true;
    if (!validateField("shiftDate")) isValid = false;
    if (!validateField("teamType")) isValid = false;
    if (!validateField("originalShift")) isValid = false;
    if (!validateField("newShift")) isValid = false;
    if (!validateField("shiftReason")) isValid = false;
    if (formData.originalShift && formData.newShift && formData.originalShift === formData.newShift) {
      isValid = false;
      const errorContainer = document.getElementById("form-error-container");
      if (errorContainer) {
        errorContainer.textContent = "新班別不能與原班別相同";
        errorContainer.style.display = "block";
      }
      showFieldError("originalShift");
      showFieldError("newShift");
      const shiftWarning = document.getElementById("shiftWarning");
      if (shiftWarning) {
        shiftWarning.style.display = "block";
      }
    }
    return isValid;
  }

  // 表單輔助函式

  // 創建必填標籤
  function createLabelWithRequired(labelText) {
    const labelContainer = document.createElement("div");
    labelContainer.style.display = "flex";
    labelContainer.style.alignItems = "center";
    labelContainer.style.marginBottom = "5px";
    labelContainer.style.flexWrap = "nowrap";
    
    const labelElement = document.createElement("label");
    labelElement.textContent = labelText;
    labelElement.style.display = "inline-block";
    labelElement.style.margin = "0";
    labelElement.style.padding = "0";
    labelElement.style.lineHeight = "1.5";
    
    const requiredMark = document.createElement("span");
    requiredMark.textContent = "error";
    requiredMark.className = "required-mark material-icons";
    requiredMark.style.color = "#ff4444";
    requiredMark.style.fontSize = "16px";
    requiredMark.style.marginLeft = "5px";
    requiredMark.style.fontWeight = "normal";
    requiredMark.style.whiteSpace = "nowrap";
    requiredMark.style.verticalAlign = "middle";
    requiredMark.style.lineHeight = "1.5";
    requiredMark.style.display = "none";
    
    labelContainer.appendChild(labelElement);
    labelContainer.appendChild(requiredMark);
    
    return labelContainer;
  }

  // 創建文字輸入
  function createTextInputGroup(name, label, placeholder = "") {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";
    group.style.position = "relative";

    const labelContainer = createLabelWithRequired(label);

    const input = document.createElement("input");
    input.type = "text";
    input.name = name;
    input.id = name;
    input.className = "form-input";
    input.placeholder = placeholder;
    input.style.width = "100%";
    input.style.height = "42px";
    input.style.boxSizing = "border-box";
    input.style.padding = "10px 15px";
    input.style.border = "1px solid #555";
    input.style.borderRadius = "5px";
    input.style.backgroundColor = "#333";
    input.style.color = "#fff";

    group.appendChild(labelContainer);
    group.appendChild(input);
    addValidationEvents(input);
    return group;
  }

  // 創建日期輸入
  function createDateInputGroup(name, label) {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";
    
    const labelContainer = createLabelWithRequired(label);
    
    const inputContainer = document.createElement("div");
    inputContainer.className = "date-input-container";
    inputContainer.style.display = "flex";
    inputContainer.style.alignItems = "center"; 
    inputContainer.style.height = "42px";
    inputContainer.style.gap = "10px";
    inputContainer.style.width = "100%";
    
    const input = document.createElement("input");
    input.type = "text";
    input.name = name;
    input.id = name;
    input.required = true;
    input.className = "form-input datepicker";
    input.placeholder = "請選擇日期";
    input.autocomplete = "off";
    input.readOnly = true;
    input.style.height = "42px";
    input.style.boxSizing = "border-box";
    input.style.padding = "10px";
    input.style.border = "1px solid #555";
    input.style.borderRadius = "5px";
    input.style.backgroundColor = "#333";
    input.style.color = "#fff";
    input.style.textAlign = "left";
    input.style.paddingLeft = "15px";
    input.style.flex = "1";
    
    const weekdayElement = document.createElement("div");
    weekdayElement.id = `${name}-weekday`;
    weekdayElement.className = "weekday-display";
    weekdayElement.style.height = "42px";
    weekdayElement.style.minWidth = "70px";
    weekdayElement.style.boxSizing = "border-box";
    weekdayElement.style.padding = "0 15px";
    weekdayElement.style.backgroundColor = "#444";
    weekdayElement.style.borderRadius = "5px";
    weekdayElement.style.display = "flex";
    weekdayElement.style.justifyContent = "center";
    weekdayElement.style.alignItems = "center";
    weekdayElement.style.fontSize = "16px";
    weekdayElement.style.fontWeight = "bold";
    weekdayElement.style.color = "#3498db";
    weekdayElement.style.whiteSpace = "nowrap";
    
    inputContainer.appendChild(input);
    inputContainer.appendChild(weekdayElement);
    
    group.appendChild(labelContainer);
    group.appendChild(inputContainer);
    
    // 初始化 jQuery UI Datepicker
    setTimeout(() => {
      const weekdays = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"];
      
      if (typeof $ !== 'undefined' && $.fn.datepicker) {
        $(input).datepicker({
          dateFormat: 'yy-mm-dd',
          onSelect: function(dateText, inst) {
            const date = $(this).datepicker('getDate');
            if (date) {
              weekdayElement.textContent = weekdays[date.getDay()];
              clearFieldError(name);
              // 觸發 change 事件，讓綁定的事件處理器執行
              $(this).trigger('change');
            }
          },
          beforeShow: function(input, inst) {
            setTimeout(function() {
              // 使用直接配置定義，確保載入穩定性
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
              
              // 移除動態定位邏輯，讓 jQuery UI 自行處理定位
              // 這可避免日期選擇後的位置偏移問題
            }, 10);
          }
        });
      }
      
      addValidationEvents(input);
    }, 0);
    
    return group;
  }

  // 創建時間輸入
  function createTimeInputGroup(name, label, skipDefaultTime = false, isLeaveForm = false) {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";

    const labelContainer = createLabelWithRequired(label);

    const inputContainer = document.createElement("div");
    inputContainer.className = "time-input-container";
    inputContainer.style.display = "flex";
    inputContainer.style.alignItems = "center";
    inputContainer.style.height = "42px";
    inputContainer.style.position = "relative";
    inputContainer.style.width = "calc(100% - 0px)";

    const input = document.createElement("input");
    input.type = "text";
    input.name = name;
    input.id = name;
    input.required = true;
    input.className = "form-input timepicker";
    input.placeholder = "請選擇時間";
    input.autocomplete = "off";
    input.readOnly = true;
    input.style.height = "42px";
    input.style.boxSizing = "border-box";
    input.style.padding = "10px";
    input.style.border = "1px solid #555";
    input.style.borderRadius = "5px";
    input.style.backgroundColor = "#333";
    input.style.color = "#fff";
    input.style.textAlign = "left";
    input.style.paddingLeft = "15px";
    input.style.width = "100%";
    
    // 確保預設值為空
    input.value = "";

    const spacerElement = document.createElement("div");
    spacerElement.className = "spacer-element";
    spacerElement.style.position = "absolute";
    spacerElement.style.right = "-90px";
    spacerElement.style.top = "0";
    spacerElement.style.width = "80px";
    spacerElement.style.minWidth = "80px";
    spacerElement.style.height = "42px";
    spacerElement.style.visibility = "hidden";

    inputContainer.appendChild(input);
    inputContainer.appendChild(spacerElement);
    
    group.appendChild(labelContainer);
    group.appendChild(inputContainer);
    
    // 初始化 jQuery UI Timepicker
    setTimeout(() => {
      const timepickerOptions = {
        timeFormat: 'HH:mm',
        controlType: 'select',
        showSecond: false,
        showMillisec: false,
        showMicrosec: false,
        showTimezone: false,
        amNames: ['上午', 'AM', 'A'],
        pmNames: ['下午', 'PM', 'P'],
        hourMin: 0,
        hourMax: 23,
        hourGrid: 0,
        minuteGrid: isLeaveForm ? 30 : 0,
        minuteMin: 0,
        minuteMax: 59,
        stepMinute: isLeaveForm ? 30 : 1,
        showButtonPanel: false,
        timeOnlyTitle: '二十四小時制',
        onSelect: function(timeText, inst) {
          if (timeText) {
            clearFieldError(name);
            // 觸發 change 事件，讓綁定的事件處理器執行
            $(input).trigger('change');
          }
        },
        onClose: function(timeText, inst) {
          if (timeText) {
            clearFieldError(name);
          }
        },
        beforeShow: function(input, inst) {
          setTimeout(function() {
            // 使用統一配置
            const pickerStyles = getPickerStyles();
            const TIMEPICKER_CONFIG = pickerStyles.timepicker;
            const DATEPICKER_CONFIG = pickerStyles.datepicker;
            
            // 修復桌面版定位問題
            $('#ui-datepicker-div').css({
              'width': `${TIMEPICKER_CONFIG.width}px`,
              'min-width': `${TIMEPICKER_CONFIG.minWidth}px`,
              'max-width': `${TIMEPICKER_CONFIG.maxWidth}px`,
              'background-color': TIMEPICKER_CONFIG.background,
              'border': 'none',
              'box-shadow': DATEPICKER_CONFIG.boxShadow,
              'margin-top': '0',
              'padding': '0',
              'overflow': 'hidden',
              'z-index': '10000'
            });
            
            $('.ui-timepicker-div').css({
              'width': '100%',
              'min-width': `${TIMEPICKER_CONFIG.minWidth}px`,
              'max-width': `${TIMEPICKER_CONFIG.maxWidth}px`,
              'background-color': TIMEPICKER_CONFIG.background,
              'border': 'none',
              'border-top': TIMEPICKER_CONFIG.borderTop,
              'border-radius': '0',
              'padding': TIMEPICKER_CONFIG.padding,
              'box-shadow': 'none',
              'overflow': 'visible'
            });

            $('.ui-slider').hide();

            $('.ui-timepicker-div dl').css({
              'display': 'flex',
              'flex-direction': 'row',
              'justify-content': 'space-between',
              'align-items': 'flex-start',
              'margin': '0',
              'padding': '0'
            });
            
            $('.ui-timepicker-div dd').css({
              'margin': '0',
              'padding': '0',
              'width': '48%',
              'text-align': 'center'
            });
            
            $('.ui-timepicker-div select').css({
              'width': '95%',
              'margin': '0 auto',
              'text-align': 'center',
              'text-align-last': 'center',
              '-moz-text-align-last': 'center',
              '-webkit-text-align-last': 'center',
              'padding': '8px 4px',
              'font-size': '16px',
              'background-color': '#333',
              'color': '#fff',
              'border': '1px solid #555',
              'border-radius': '4px'
            });

            if ($('.ui-timepicker-div .timepicker-buttons').length === 0) {
              const buttonContainer = $('<div class="timepicker-buttons"></div>');
              buttonContainer.css({
                'display': 'flex',
                'justify-content': 'space-between',
                'margin-top': '15px',
                'padding-top': '15px',
                'border-top': '1px solid #444',
                'width': '100%'
              });
              
              const cancelButton = $('<button type="button" class="timepicker-button cancel">取消</button>');
              cancelButton.css({
                'background-color': '#666', 'color': '#fff', 'border': '1px solid #555',
                'border-radius': '5px', 'padding': '8px 15px', 'font-size': '14px',
                'cursor': 'pointer', 'min-width': '80px', 'text-align': 'center'
              });
              cancelButton.on('click', function() { $(input).timepicker('hide'); });
              
              const confirmButton = $('<button type="button" class="timepicker-button confirm">確認</button>');
              confirmButton.css({
                'background-color': '#2c7a3e', 'color': '#fff', 'border': '1px solid #1e5a2d',
                'border-radius': '5px', 'padding': '8px 15px', 'font-size': '14px',
                'cursor': 'pointer', 'min-width': '80px', 'text-align': 'center'
              });
              confirmButton.on('click', function() {
                // 直接從下拉選單讀取值
                const hourSelect = $('.ui-timepicker-div select').eq(0);
                const minuteSelect = $('.ui-timepicker-div select').eq(1);
                
                if (hourSelect.length && minuteSelect.length) {
                  const hours = hourSelect.val();
                  const minutes = minuteSelect.val();
                  
                  if (hours !== undefined && minutes !== undefined) {
                    const timeString = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
                    
                    // 對於跨晝夜加班結束時間，如果是00:00，保持原樣但確保正確設置
                    input.value = timeString;
                    
                    // 觸發change事件，用於加班時數計算等
                    $(input).trigger('change');
                  }
                }
                
                clearFieldError(name);
                $(input).timepicker('hide');
              });
              
              buttonContainer.append(cancelButton);
              buttonContainer.append(confirmButton);
              
              $('.ui-timepicker-div').append(buttonContainer);
            }
          }, 10);
        }
      };
      
      // 如果 skipDefaultTime 為 true，不設定預設時間
      if (!skipDefaultTime) {
        // 可以在這裡設定預設時間，但目前保持為空
      }
      
      $(input).timepicker(timepickerOptions);
      
      addValidationEvents(input);
    }, 0);
    
    return group;
  }

  // 創建下拉選單
  function createSelectInputGroup(name, label, options) {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";

    const labelContainer = createLabelWithRequired(label);

    const select = document.createElement("select");
    select.name = name;
    select.id = name;
    select.className = "form-select";
    select.required = true;
    select.style.width = "100%";
    select.style.height = "42px";
    select.style.boxSizing = "border-box";
    select.style.padding = "8px 15px";
    select.style.border = "1px solid #555";
    select.style.borderRadius = "5px";
    select.style.backgroundColor = "#333";
    select.style.color = "#fff";

    options.forEach(option => {
      const optionElement = document.createElement("option");
      optionElement.value = option.value;
      optionElement.textContent = option.label;
      // 處理分隔線選項
      if (option.disabled) {
        optionElement.disabled = true;
        optionElement.style.fontWeight = "bold";
        optionElement.style.color = "#888";
        optionElement.style.backgroundColor = "#2a2a2a";
      }
      select.appendChild(optionElement);
    });

    group.appendChild(labelContainer);
    group.appendChild(select);
    addValidationEvents(select);
    return group;
  }

  // 創建文字區域
  function createTextAreaInputGroup(name, label, placeholder = "") {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";

    const labelContainer = createLabelWithRequired(label);

    const textarea = document.createElement("textarea");
    textarea.name = name;
    textarea.id = name;
    textarea.className = "form-textarea";
    textarea.placeholder = placeholder;
    textarea.required = true;
    textarea.rows = 2;
    textarea.style.width = "100%";
    textarea.style.boxSizing = "border-box";
    textarea.style.padding = "10px 15px";
    textarea.style.border = "1px solid #555";
    textarea.style.borderRadius = "5px";
    textarea.style.backgroundColor = "#333";
    textarea.style.color = "#fff";
    textarea.style.resize = "vertical";
    textarea.style.minHeight = "50px";

    group.appendChild(labelContainer);
    group.appendChild(textarea);
    addValidationEvents(textarea);
    return group;
  }

  // 創建動作按鈕 (上班/下班)
  function createActionButtonGroup(name, label) {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";

    const labelContainer = createLabelWithRequired(label);

    const buttonContainer = document.createElement("div");
    buttonContainer.style.display = "flex";
    buttonContainer.style.gap = "8px";
    buttonContainer.style.width = "100%";
    buttonContainer.style.boxSizing = "border-box";
    buttonContainer.style.justifyContent = "space-between";

    const onButton = document.createElement("button");
    onButton.type = "button";
    onButton.className = "btn action-btn";
    onButton.textContent = "上班";
    onButton.dataset.value = "上班";
    onButton.style.flex = "1";
    onButton.style.padding = "12px";
    onButton.style.border = "2px solid #555";
    onButton.style.borderRadius = "5px";
    onButton.style.backgroundColor = "#333";
    onButton.style.color = "#fff";
    onButton.style.cursor = "pointer";

    const offButton = document.createElement("button");
    offButton.type = "button";
    offButton.className = "btn action-btn";
    offButton.textContent = "下班";
    offButton.dataset.value = "下班";
    offButton.style.flex = "1";
    offButton.style.padding = "12px";
    offButton.style.border = "2px solid #555";
    offButton.style.borderRadius = "5px";
    offButton.style.backgroundColor = "#333";
    offButton.style.color = "#fff";
    offButton.style.cursor = "pointer";

    // 隱藏的input來儲存選擇的值
    const hiddenInput = document.createElement("input");
    hiddenInput.type = "hidden";
    hiddenInput.name = name;
    hiddenInput.id = name;
    hiddenInput.required = true;

    // 按鈕點擊事件
    [onButton, offButton].forEach(button => {
      button.addEventListener("click", () => {
        // 重置所有按鈕樣式
        [onButton, offButton].forEach(btn => {
          btn.style.borderColor = "#555";
          btn.style.backgroundColor = "#333";
          btn.style.boxShadow = "none";
          btn.style.color = "#fff";
          btn.style.fontWeight = "normal";
        });
        
        // 設置選中樣式 - 更明顯的效果
        button.style.borderColor = "#3498db";
        button.style.backgroundColor = "#2c3e50";
        button.style.boxShadow = "0 0 10px rgba(52, 152, 219, 0.5)";
        button.style.color = "#3498db";
        button.style.fontWeight = "bold";
        
        // 設置值
        hiddenInput.value = button.dataset.value;
        clearFieldError(name);
      });
    });

    buttonContainer.appendChild(onButton);
    buttonContainer.appendChild(offButton);

    group.appendChild(labelContainer);
    group.appendChild(buttonContainer);
    group.appendChild(hiddenInput);
    
    addValidationEvents(hiddenInput);
    return group;
  }

  // 創建見證人輸入
  function createWitnessInputGroup(name, label) {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";

    const labelContainer = createLabelWithRequired(label);

    const input = document.createElement("input");
    input.type = "text";
    input.name = name;
    input.id = name;
    input.className = "form-input";
    input.required = true; // 見證人為必填欄位
    input.placeholder = "請輸入見證人姓名";
    input.style.width = "100%";
    input.style.height = "42px";
    input.style.boxSizing = "border-box";
    input.style.padding = "10px 15px";
    input.style.border = "1px solid #555";
    input.style.borderRadius = "5px";
    input.style.backgroundColor = "#333";
    input.style.color = "#fff";

    group.appendChild(labelContainer);
    group.appendChild(input);
    addValidationEvents(input);
    return group;
  }

  // 創建日期範圍
  function createDateRangeInputGroup(startName, endName, label) {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";

    const labelContainer = createLabelWithRequired(label);

    const rangeContainer = document.createElement("div");
    rangeContainer.style.display = "flex";
    rangeContainer.style.gap = "10px";
    rangeContainer.style.alignItems = "center";

    // 開始日期
    const startInputContainer = document.createElement("div");
    startInputContainer.style.flex = "1";
    startInputContainer.style.position = "relative";

    const startInput = document.createElement("input");
    startInput.type = "text";
    startInput.name = startName;
    startInput.id = startName;
    startInput.className = "form-input datepicker";
    startInput.placeholder = "開始日期";
    startInput.autocomplete = "off";
    startInput.readOnly = true;
    startInput.required = true;
    startInput.style.width = "100%";
    startInput.style.height = "42px";
    startInput.style.boxSizing = "border-box";
    startInput.style.padding = "10px";
    startInput.style.border = "1px solid #555";
    startInput.style.borderRadius = "5px";
    startInput.style.backgroundColor = "#333";
    startInput.style.color = "#fff";

    startInputContainer.appendChild(startInput);

    // 分隔符
    const separator = document.createElement("span");
    separator.textContent = "至";
    separator.style.color = "#fff";
    separator.style.fontSize = "14px";

    // 結束日期
    const endInputContainer = document.createElement("div");
    endInputContainer.style.flex = "1";
    endInputContainer.style.position = "relative";

    const endInput = document.createElement("input");
    endInput.type = "text";
    endInput.name = endName;
    endInput.id = endName;
    endInput.className = "form-input datepicker";
    endInput.placeholder = "結束日期";
    endInput.autocomplete = "off";
    endInput.readOnly = true;
    endInput.required = true;
    endInput.style.width = "100%";
    endInput.style.height = "42px";
    endInput.style.boxSizing = "border-box";
    endInput.style.padding = "10px";
    endInput.style.border = "1px solid #555";
    endInput.style.borderRadius = "5px";
    endInput.style.backgroundColor = "#333";
    endInput.style.color = "#fff";

    endInputContainer.appendChild(endInput);

    rangeContainer.appendChild(startInputContainer);
    rangeContainer.appendChild(separator);
    rangeContainer.appendChild(endInputContainer);

    group.appendChild(labelContainer);
    group.appendChild(rangeContainer);

    // 初始化日期選擇器
    setTimeout(() => {
      if (typeof $ !== 'undefined' && $.fn.datepicker) {
        $(startInput).datepicker({
          dateFormat: 'yy-mm-dd',
          onSelect: function(dateText) {
            clearFieldError(startName);
            // 設置結束日期的最小值
            $(endInput).datepicker("option", "minDate", dateText);
          }
        });

        $(endInput).datepicker({
          dateFormat: 'yy-mm-dd',
          onSelect: function(dateText) {
            clearFieldError(endName);
          }
        });
      }

      addValidationEvents(startInput);
      addValidationEvents(endInput);
    }, 0);

    return group;
  }

  // 創建日期時間範圍輸入組（版本1：照原圖布局 + 增加時間）
  function createDateTimeRangeInputGroup(startDateName, endDateName, startTimeName, endTimeName, label) {
    const group = document.createElement("div");
    group.className = "form-group";
    group.style.marginBottom = "15px";

    const labelContainer = createLabelWithRequired(label);
    group.appendChild(labelContainer);

    const dateRangeContainer = document.createElement("div");
    dateRangeContainer.style.cssText = `
      display: flex;
      flex-direction: column;
      gap: 15px;
    `;

    // === 開始日期行 ===
    const startDateRow = document.createElement("div");
    startDateRow.style.cssText = `
      display: flex;
      gap: 10px;
      align-items: center;
    `;

    const startDateLabel = document.createElement("span");
    startDateLabel.textContent = "開始日期：";
    startDateLabel.style.cssText = `
      color: #ccc;
      font-size: 14px;
      min-width: 80px;
    `;
    startDateRow.appendChild(startDateLabel);

    const startDateInput = document.createElement("input");
    startDateInput.type = "text";
    startDateInput.name = startDateName;
    startDateInput.id = startDateName;
    startDateInput.className = "form-input datepicker";
    startDateInput.placeholder = "選擇日期";
    startDateInput.autocomplete = "off";
    startDateInput.readOnly = true;
    startDateInput.required = true;
    startDateInput.style.cssText = `
      flex: 1;
      max-width: 200px;
      height: 42px;
      padding: 10px;
      border: 1px solid #555;
      border-radius: 5px;
      background: #333;
      color: #fff;
    `;
    startDateRow.appendChild(startDateInput);

    const startWeekday = document.createElement("button");
    startWeekday.type = "button";
    startWeekday.className = "weekday-display";
    startWeekday.style.cssText = `
      padding: 8px 16px;
      background: #444;
      border: 1px solid #555;
      border-radius: 5px;
      color: #fff;
      min-width: 60px;
      cursor: default;
    `;
    startWeekday.textContent = "週一";
    startDateRow.appendChild(startWeekday);

    dateRangeContainer.appendChild(startDateRow);

    // === 開始時間行 ===
    const startTimeRow = document.createElement("div");
    startTimeRow.style.cssText = `
      display: flex;
      gap: 10px;
      align-items: center;
      margin-top: 10px;
    `;

    const startTimeLabel = document.createElement("span");
    startTimeLabel.textContent = "開始時間：";
    startTimeLabel.style.cssText = `
      color: #ccc;
      font-size: 14px;
      min-width: 80px;
    `;
    startTimeRow.appendChild(startTimeLabel);

    const startTimeInput = document.createElement("input");
    startTimeInput.type = "text";
    startTimeInput.name = startTimeName;
    startTimeInput.id = startTimeName;
    startTimeInput.className = "form-input timepicker";
    startTimeInput.placeholder = "HH:MM";
    startTimeInput.value = "08:30";
    startTimeInput.required = true;
    startTimeInput.style.cssText = `
      width: 120px;
      height: 42px;
      padding: 10px;
      border: 1px solid #555;
      border-radius: 5px;
      background: #333;
      color: #fff;
      text-align: center;
    `;
    startTimeRow.appendChild(startTimeInput);

    dateRangeContainer.appendChild(startTimeRow);

    // === 結束日期行 ===
    const endDateRow = document.createElement("div");
    endDateRow.style.cssText = `
      display: flex;
      gap: 10px;
      align-items: center;
    `;

    const endDateLabel = document.createElement("span");
    endDateLabel.textContent = "結束日期：";
    endDateLabel.style.cssText = `
      color: #ccc;
      font-size: 14px;
      min-width: 80px;
    `;
    endDateRow.appendChild(endDateLabel);

    const endDateInput = document.createElement("input");
    endDateInput.type = "text";
    endDateInput.name = endDateName;
    endDateInput.id = endDateName;
    endDateInput.className = "form-input datepicker";
    endDateInput.placeholder = "選擇日期";
    endDateInput.autocomplete = "off";
    endDateInput.readOnly = true;
    endDateInput.required = true;
    endDateInput.style.cssText = `
      flex: 1;
      max-width: 200px;
      height: 42px;
      padding: 10px;
      border: 1px solid #555;
      border-radius: 5px;
      background: #333;
      color: #fff;
    `;
    endDateRow.appendChild(endDateInput);

    const endWeekday = document.createElement("button");
    endWeekday.type = "button";
    endWeekday.className = "weekday-display";
    endWeekday.style.cssText = `
      padding: 8px 16px;
      background: #444;
      border: 1px solid #555;
      border-radius: 5px;
      color: #fff;
      min-width: 60px;
      cursor: default;
    `;
    endWeekday.textContent = "週一";
    endDateRow.appendChild(endWeekday);

    dateRangeContainer.appendChild(endDateRow);

    // === 結束時間行 ===
    const endTimeRow = document.createElement("div");
    endTimeRow.style.cssText = `
      display: flex;
      gap: 10px;
      align-items: center;
      margin-top: 10px;
    `;

    const endTimeLabel = document.createElement("span");
    endTimeLabel.textContent = "結束時間：";
    endTimeLabel.style.cssText = `
      color: #ccc;
      font-size: 14px;
      min-width: 80px;
    `;
    endTimeRow.appendChild(endTimeLabel);

    const endTimeInput = document.createElement("input");
    endTimeInput.type = "text";
    endTimeInput.name = endTimeName;
    endTimeInput.id = endTimeName;
    endTimeInput.className = "form-input timepicker";
    endTimeInput.placeholder = "HH:MM";
    endTimeInput.value = "17:30";
    endTimeInput.required = true;
    endTimeInput.style.cssText = `
      width: 120px;
      height: 42px;
      padding: 10px;
      border: 1px solid #555;
      border-radius: 5px;
      background: #333;
      color: #fff;
      text-align: center;
    `;
    endTimeRow.appendChild(endTimeInput);

    dateRangeContainer.appendChild(endTimeRow);

    group.appendChild(dateRangeContainer);

    // === 初始化 ===
    setTimeout(() => {
      const weekdays = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];

      // 初始化日期選擇器
      if (typeof $ !== 'undefined' && $.fn.datepicker) {
        $(startDateInput).datepicker({
          dateFormat: 'yy-mm-dd',
          onSelect: function(dateText) {
            clearFieldError(startDateName);
            $(endDateInput).datepicker("option", "minDate", dateText);
            
            const date = new Date(dateText);
            startWeekday.textContent = weekdays[date.getDay()];
          },
          beforeShow: function(input, inst) {
            setTimeout(function() {
              // 修正定位相關的 CSS 屬性
              $('#ui-datepicker-div').css({
                'position': 'absolute',
                'left': 'auto',
                'right': 'auto',
                'top': 'auto',
                'transform': 'none',
                'margin-left': '0',
                'margin-right': '0',
                'z-index': '10000'
              });
            }, 50);
          }
        });

        $(endDateInput).datepicker({
          dateFormat: 'yy-mm-dd',
          onSelect: function(dateText) {
            clearFieldError(endDateName);
            
            const date = new Date(dateText);
            endWeekday.textContent = weekdays[date.getDay()];
          },
          beforeShow: function(input, inst) {
            setTimeout(function() {
              // 修正定位相關的 CSS 屬性
              $('#ui-datepicker-div').css({
                'position': 'absolute',
                'left': 'auto',
                'right': 'auto',
                'top': 'auto',
                'transform': 'none',
                'margin-left': '0',
                'margin-right': '0',
                'z-index': '10000'
              });
            }, 50);
          }
        });
      }

      // 初始化時間選擇器
      if (typeof $ !== 'undefined' && $.fn.timepicker) {
        $(startTimeInput).timepicker({
          timeFormat: 'HH:mm',
          controlType: 'select',
          showSecond: false,
          showMillisec: false,
          showMicrosec: false,
          showTimezone: false,
          amNames: ['上午', 'AM', 'A'],
          pmNames: ['下午', 'PM', 'P'],
          hourMin: 0,
          hourMax: 23,
          hourGrid: 0,
          minuteGrid: 30,
          stepMinute: 30,
          showButtonPanel: false,
          timeOnlyTitle: '二十四小時制',
          onSelect: function(timeText, inst) {
            if (timeText) {
              clearFieldError(startTimeName);
            }
          },
          beforeShow: function(input, inst) {
            setTimeout(function() {
              // 使用統一配置
              const pickerStyles = getPickerStyles();
              const TIMEPICKER_CONFIG = pickerStyles.timepicker;
              
              $('#ui-datepicker-div').css({
                'width': `${TIMEPICKER_CONFIG.width}px`,
                'min-width': `${TIMEPICKER_CONFIG.minWidth}px`,
                'max-width': `${TIMEPICKER_CONFIG.maxWidth}px`,
                'background-color': TIMEPICKER_CONFIG.background,
                'border': 'none',
                'box-shadow': '0 4px 12px rgba(0,0,0,0.5)',
                'margin-top': '0',
                'padding': '0',
                'overflow': 'hidden',
                'z-index': '10000'
              });
              
              $('.ui-timepicker-div').css({
                'width': '100%',
                'background-color': TIMEPICKER_CONFIG.background,
                'padding': TIMEPICKER_CONFIG.padding
              });

              $('.ui-slider').hide();

              $('.ui-timepicker-div dl').css({
                'display': 'flex',
                'flex-direction': 'row',
                'justify-content': 'center',
                'align-items': 'flex-start',
                'margin': '0',
                'padding': '0'
              });
              
              $('.ui-timepicker-div dd').css({
                'margin': '0 5px',
                'padding': '0',
                'width': 'auto',
                'text-align': 'center'
              });
              
              $('.ui-timepicker-div select').css({
                'width': '95%',
                'margin': '0 auto',
                'text-align': 'center',
                'padding': '8px 4px',
                'font-size': '16px',
                'background-color': '#333',
                'color': '#fff',
                'border': '1px solid #555',
                'border-radius': '4px',
                'appearance': 'none',
                '-webkit-appearance': 'none',
                '-moz-appearance': 'none'
              });
              
              // 修正選項對齊
              $('.ui-timepicker-div select option').css({
                'text-align': 'center',
                'padding': '4px'
              });

              // 添加自定義按鈕
              if ($('.ui-timepicker-div .timepicker-buttons').length === 0) {
                const buttonContainer = $('<div class="timepicker-buttons"></div>');
                buttonContainer.css({
                  'display': 'flex',
                  'justify-content': 'space-between',
                  'margin-top': '15px',
                  'padding-top': '15px',
                  'border-top': '1px solid #444',
                  'width': '100%'
                });
                
                const cancelButton = $('<button type="button" class="timepicker-button cancel">取消</button>');
                cancelButton.css({
                  'background-color': '#666', 'color': '#fff', 'border': '1px solid #555',
                  'border-radius': '5px', 'padding': '8px 15px', 'font-size': '14px',
                  'cursor': 'pointer', 'min-width': '80px', 'text-align': 'center'
                });
                cancelButton.on('click', function() { $(input).timepicker('hide'); });
                
                const confirmButton = $('<button type="button" class="timepicker-button confirm">確認</button>');
                confirmButton.css({
                  'background-color': '#2c7a3e', 'color': '#fff', 'border': '1px solid #1e5a2d',
                  'border-radius': '5px', 'padding': '8px 15px', 'font-size': '14px',
                  'cursor': 'pointer', 'min-width': '80px', 'text-align': 'center'
                });
                confirmButton.on('click', function() {
                  const hourSelect = $('.ui-timepicker-div select').eq(0);
                  const minuteSelect = $('.ui-timepicker-div select').eq(1);
                  
                  if (hourSelect.length && minuteSelect.length) {
                    const hours = hourSelect.val();
                    const minutes = minuteSelect.val();
                    
                    if (hours !== undefined && minutes !== undefined) {
                      const timeString = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
                      input.value = timeString;
                      $(input).trigger('change');
                    }
                  }
                  
                  clearFieldError(startTimeName);
                  $(input).timepicker('hide');
                });
                
                buttonContainer.append(cancelButton);
                buttonContainer.append(confirmButton);
                
                $('.ui-timepicker-div').append(buttonContainer);
              }
            }, 10);
          }
        });

        $(endTimeInput).timepicker({
          timeFormat: 'HH:mm',
          controlType: 'select',
          showSecond: false,
          showMillisec: false,
          showMicrosec: false,
          showTimezone: false,
          amNames: ['上午', 'AM', 'A'],
          pmNames: ['下午', 'PM', 'P'],
          hourMin: 0,
          hourMax: 23,
          hourGrid: 0,
          minuteGrid: 30,
          stepMinute: 30,
          showButtonPanel: false,
          timeOnlyTitle: '二十四小時制',
          onSelect: function(timeText, inst) {
            if (timeText) {
              clearFieldError(endTimeName);
            }
          },
          beforeShow: function(input, inst) {
            setTimeout(function() {
              // 使用統一配置
              const pickerStyles = getPickerStyles();
              const TIMEPICKER_CONFIG = pickerStyles.timepicker;
              
              $('#ui-datepicker-div').css({
                'width': `${TIMEPICKER_CONFIG.width}px`,
                'min-width': `${TIMEPICKER_CONFIG.minWidth}px`,
                'max-width': `${TIMEPICKER_CONFIG.maxWidth}px`,
                'background-color': TIMEPICKER_CONFIG.background,
                'border': 'none',
                'box-shadow': '0 4px 12px rgba(0,0,0,0.5)',
                'margin-top': '0',
                'padding': '0',
                'overflow': 'hidden',
                'z-index': '10000'
              });
              
              $('.ui-timepicker-div').css({
                'width': '100%',
                'background-color': TIMEPICKER_CONFIG.background,
                'padding': TIMEPICKER_CONFIG.padding
              });

              $('.ui-slider').hide();

              $('.ui-timepicker-div dl').css({
                'display': 'flex',
                'flex-direction': 'row',
                'justify-content': 'center',
                'align-items': 'flex-start',
                'margin': '0',
                'padding': '0'
              });
              
              $('.ui-timepicker-div dd').css({
                'margin': '0 5px',
                'padding': '0',
                'width': 'auto',
                'text-align': 'center'
              });
              
              $('.ui-timepicker-div select').css({
                'width': '95%',
                'margin': '0 auto',
                'text-align': 'center',
                'padding': '8px 4px',
                'font-size': '16px',
                'background-color': '#333',
                'color': '#fff',
                'border': '1px solid #555',
                'border-radius': '4px',
                'appearance': 'none',
                '-webkit-appearance': 'none',
                '-moz-appearance': 'none'
              });
              
              // 修正選項對齊
              $('.ui-timepicker-div select option').css({
                'text-align': 'center',
                'padding': '4px'
              });

              // 添加自定義按鈕
              if ($('.ui-timepicker-div .timepicker-buttons').length === 0) {
                const buttonContainer = $('<div class="timepicker-buttons"></div>');
                buttonContainer.css({
                  'display': 'flex',
                  'justify-content': 'space-between',
                  'margin-top': '15px',
                  'padding-top': '15px',
                  'border-top': '1px solid #444',
                  'width': '100%'
                });
                
                const cancelButton = $('<button type="button" class="timepicker-button cancel">取消</button>');
                cancelButton.css({
                  'background-color': '#666', 'color': '#fff', 'border': '1px solid #555',
                  'border-radius': '5px', 'padding': '8px 15px', 'font-size': '14px',
                  'cursor': 'pointer', 'min-width': '80px', 'text-align': 'center'
                });
                cancelButton.on('click', function() { $(input).timepicker('hide'); });
                
                const confirmButton = $('<button type="button" class="timepicker-button confirm">確認</button>');
                confirmButton.css({
                  'background-color': '#2c7a3e', 'color': '#fff', 'border': '1px solid #1e5a2d',
                  'border-radius': '5px', 'padding': '8px 15px', 'font-size': '14px',
                  'cursor': 'pointer', 'min-width': '80px', 'text-align': 'center'
                });
                confirmButton.on('click', function() {
                  const hourSelect = $('.ui-timepicker-div select').eq(0);
                  const minuteSelect = $('.ui-timepicker-div select').eq(1);
                  
                  if (hourSelect.length && minuteSelect.length) {
                    const hours = hourSelect.val();
                    const minutes = minuteSelect.val();
                    
                    if (hours !== undefined && minutes !== undefined) {
                      const timeString = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
                      input.value = timeString;
                      $(input).trigger('change');
                    }
                  }
                  
                  clearFieldError(endTimeName);
                  $(input).timepicker('hide');
                });
                
                buttonContainer.append(cancelButton);
                buttonContainer.append(confirmButton);
                
                $('.ui-timepicker-div').append(buttonContainer);
              }
            }, 10);
          }
        });
      }

      // 注意：createDateTimeRangeInputGroup 已經不再被使用
      // 請假時數計算功能已經移到 createLeaveRequestFormContent 函數中

      // 添加驗證事件
      addValidationEvents(startDateInput);
      addValidationEvents(endDateInput);
      addValidationEvents(startTimeInput);
      addValidationEvents(endTimeInput);
    }, 0);

    return group;
  }

  // 創建簡單的時間輸入元素
  function createTimeInputElement(name, defaultValue) {
    const input = document.createElement("input");
    input.type = "text";
    input.name = name;
    input.id = name;
    input.className = "form-input timepicker";
    input.placeholder = "HH:MM";
    input.value = defaultValue || "";
    input.pattern = "^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$";
    input.required = true;
    input.style.cssText = `
      width: 100%;
      height: 42px;
      padding: 10px;
      border: 1px solid #555;
      border-radius: 5px;
      background-color: #333;
      color: #fff;
      text-align: center;
    `;

    // 初始化時間選擇器
    setTimeout(() => {
      if (typeof $ !== 'undefined' && $.fn.timepicker) {
        $(input).timepicker({
          timeFormat: 'HH:mm',
          controlType: 'select',
          showSecond: false,
          defaultTime: defaultValue || '09:00',
          hourGrid: 6,
          minuteGrid: 15,
          showButtonPanel: true,
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
        });
      }
    }, 0);

    return input;
  }

  // 假別分類配置
  const LEAVE_TYPE_CONFIG = {
    // 需要檔案上傳的假別（除了生理假、特別休假外，其他都需要）
    REQUIRES_FILE: [
      '事假',
      '病假',
      '喪假', 
      '婚假',
      '公假',
      '家庭'
    ],
    // 不需要檔案上傳的假別
    NO_FILE_REQUIRED: [
      '特別休假',
      '生理假',
      '國定假日'
    ]
  };

  // 檢查假別是否需要檔案上傳
  function requiresFileUpload(leaveType) {
    return LEAVE_TYPE_CONFIG.REQUIRES_FILE.includes(leaveType);
  }

  // 創建動態檔案上傳區塊
  function createDynamicFileUploadSection() {
    const section = document.createElement("div");
    section.id = "fileUploadSection";
    section.className = "file-upload-section";
    section.style.display = "none"; // 初始隱藏
    section.style.marginBottom = "20px";
    section.style.padding = "20px";
    section.style.border = "2px solid #4CAF50";
    section.style.borderRadius = "8px";
    section.style.backgroundColor = "rgba(76, 175, 80, 0.1)";

    // 檔案上傳狀態容器
    const statusContainer = document.createElement("div");
    statusContainer.id = "uploadStatusContainer";
    statusContainer.innerHTML = `
      <div style="
        background: rgba(255, 152, 0, 0.1); 
        border: 1px solid #ff9800; 
        border-radius: 8px; 
        padding: 16px; 
        margin-bottom: 15px;">
        <p style="margin: 0; color: #ff9800; font-size: 16px; text-align: center;">
          需要上傳證明文件
        </p>
      </div>
    `;

    // 檔案上傳按鈕 - 透明輪廓設計
    const uploadButton = document.createElement("button");
    uploadButton.type = "button";
    uploadButton.id = "fileUploadButton";
    uploadButton.className = "btn";
    uploadButton.style.cssText = `
      background-color: transparent !important;
      color: #f5a623 !important;
      padding: 16px 32px !important;
      font-size: 18px !important;
      border: 2px solid #f5a623 !important;
      border-radius: 12px !important;
      cursor: pointer !important;
      transition: all 0.3s ease !important;
      display: block !important;
      margin: 0 auto !important;
      font-weight: 500 !important;
    `;
    uploadButton.textContent = "上傳檔案";

    // 為按鈕添加 hover 效果
    uploadButton.onmouseover = function() {
      this.style.backgroundColor = '#f5a623 !important';
      this.style.color = '#ffffff !important';
      this.style.borderColor = '#f5a623 !important';
    };
    uploadButton.onmouseout = function() {
      this.style.backgroundColor = 'transparent !important';
      this.style.color = '#f5a623 !important';
      this.style.borderColor = '#f5a623 !important';
    };

    // 檔案上傳狀態指示器
    const statusIndicator = document.createElement("div");
    statusIndicator.id = "uploadStatusIndicator";
    statusIndicator.style.marginTop = "15px";
    statusIndicator.style.padding = "10px";
    statusIndicator.style.borderRadius = "5px";
    statusIndicator.style.textAlign = "center";
    statusIndicator.style.display = "none";

    // 組合所有元素
    section.appendChild(statusContainer);
    section.appendChild(uploadButton);
    section.appendChild(statusIndicator);

    return section;
  }

  // 更新檔案上傳狀態
  function updateFileUploadStatus(status, message = '') {
    const indicator = document.getElementById('uploadStatusIndicator');
    const button = document.getElementById('fileUploadButton');
    
    if (!indicator || !button) return;

    indicator.style.display = 'block';
    
    switch(status) {
      case 'waiting':
        indicator.style.backgroundColor = 'rgba(255, 193, 7, 0.1)';
        indicator.style.color = '#ffc107';
        indicator.style.border = '1px solid #ffc107';
        indicator.innerHTML = `
          <span class="material-icons" style="vertical-align: middle; margin-right: 5px;">schedule</span>
          等待檔案上傳中...
        `;
        button.disabled = false;
        button.style.opacity = '1';
        break;
        
      case 'uploading':
        indicator.style.backgroundColor = 'rgba(33, 150, 243, 0.1)';
        indicator.style.color = '#2196F3';
        indicator.style.border = '1px solid #2196F3';
        indicator.innerHTML = `
          <span class="material-icons" style="vertical-align: middle; margin-right: 5px;">cloud_upload</span>
          檔案上傳進行中...
        `;
        button.disabled = true;
        button.style.opacity = '0.6';
        break;
        
      case 'completed':
        indicator.style.backgroundColor = 'rgba(76, 175, 80, 0.1)';
        indicator.style.color = '#4CAF50';
        indicator.style.border = '1px solid #4CAF50';
        indicator.innerHTML = `
          <span class="material-icons" style="vertical-align: middle; margin-right: 5px;">check_circle</span>
          檔案上傳完成！現在可以提交申請表單
        `;
        button.innerHTML = `
          <span class="material-icons" style="margin-right: 8px;">check</span>
          檔案已上傳完成
        `;
        button.disabled = true;
        button.style.opacity = '0.8';
        break;
        
      case 'failed':
        indicator.style.backgroundColor = 'rgba(244, 67, 54, 0.1)';
        indicator.style.color = '#f44336';
        indicator.style.border = '1px solid #f44336';
        indicator.innerHTML = `
          <span class="material-icons" style="vertical-align: middle; margin-right: 5px;">error</span>
          檔案上傳失敗，請重新嘗試
        `;
        button.disabled = false;
        button.style.opacity = '1';
        break;
    }
  }

  // 創建檔案上傳提示
  function createFileUploadNote() {
    const note = document.createElement("div");
    note.className = "file-upload-note";
    note.style.backgroundColor = "rgba(255, 152, 0, 0.1)";
    note.style.border = "1px solid #ff9800";
    note.style.borderRadius = "5px";
    note.style.padding = "15px";
    note.style.marginBottom = "15px";
    note.style.color = "#ff9800";
    note.style.fontSize = "14px";

    note.innerHTML = `
      <div style="display: flex; align-items: center; margin-bottom: 8px;">
        <span class="material-icons" style="margin-right: 8px; font-size: 18px;">info</span>
        <strong>重要提醒</strong>
      </div>
      <p style="margin: 0; line-height: 1.4;">
        此假別需要上傳佐證檔案，請在表單下方完成檔案上傳後再提交申請
      </p>
    `;

    return note;
  }

  // 移除：處理檔案上傳（改為提交後處理）
  /*
  function handleFileUpload() {
    // 更新狀態為準備開啟視窗
    updateFileUploadStatus('uploading', '正在開啟檔案上傳頁面...');
    
    try {
      // 使用正確的後端 URL
      const baseUrl = "https://script.google.com/macros/s/AKfycbyrIw_NFN3chyn3XAP5V7Ab3Z67zQJwoM56d5i7tJBB0B3rEVGsPJQBWAf0WEN74sz-/exec";
      const idCode = App.UI.elements.idInput.value;
      
      // 構建檔案上傳 URL（預先上傳模式，生成臨時 requestId）
      const tempRequestId = "TEMP-" + Date.now() + "_" + idCode;
      const uploadUrl = `${baseUrl}?page=fileUpload&formType=leave&lastFour=${idCode}&requestId=${tempRequestId}`;
      
      // 開啟檔案上傳視窗
      const uploadWindow = window.open(uploadUrl, '_blank', 'width=800,height=600,scrollbars=yes,resizable=yes');
      
      if (uploadWindow) {
        console.log('檔案上傳視窗已開啟');
        
        // 監聽來自上傳視窗的訊息
        window.addEventListener('message', function(event) {
          // 驗證訊息來源（應該與上傳頁面的域名匹配）
          if (event.origin !== 'https://script.google.com') {
            return;
          }
          
          if (event.data && event.data.type === 'file_upload_complete') {
            // 檔案上傳完成
            updateFileUploadStatus('completed');
            console.log('收到檔案上傳完成訊息:', event.data);
            
            // 儲存上傳完成狀態（用於表單提交驗證）
            window.fileUploadCompleted = true;
            window.uploadedFileInfo = event.data.fileInfo || {};
          } else if (event.data && event.data.type === 'file_upload_failed') {
            // 檔案上傳失敗
            updateFileUploadStatus('failed');
            console.error('檔案上傳失敗:', event.data.error);
            
            window.fileUploadCompleted = false;
          }
        });
        
        // 移除輪詢機制，完全依賴 postMessage 通訊
        
      } else {
        // 彈出視窗被阻擋
        updateFileUploadStatus('failed');
        
        // 顯示彈出視窗被阻擋的提示
        const indicator = document.getElementById('uploadStatusIndicator');
        if (indicator) {
          indicator.innerHTML = `
            <span class="material-icons" style="vertical-align: middle; margin-right: 5px;">block</span>
            彈出視窗被阻擋，請允許彈出視窗後重試
          `;
        }
        
        console.error('檔案上傳視窗被瀏覽器阻擋');
      }
      
    } catch (error) {
      updateFileUploadStatus('failed');
      console.error('開啟檔案上傳視窗時發生錯誤:', error);
    }
  }
  */

  // 表單內容創建

  // 創建補卡表單
  function createCardCorrectionFormContent() {
    const content = document.createElement("div");
    content.className = "form-content";
    
    const dateGroup = createDateInputGroup("cardDate", "日期");
    content.appendChild(dateGroup);
    
    const timeGroup = createTimeInputGroup("cardTime", "時間");
    content.appendChild(timeGroup);
    
    const actionGroup = createActionButtonGroup("cardAction", "時段");
    content.appendChild(actionGroup);
    
    const typeOptions = [
      { value: "", label: "請選擇類型" },
      { value: "重複打卡", label: "重複打卡" },
      { value: "空缺紀錄", label: "空缺紀錄" },
      { value: "其他", label: "其他" }
    ];
    const typeGroup = createSelectInputGroup("cardType", "類型", typeOptions);
    content.appendChild(typeGroup);
    
    const reasonGroup = createTextAreaInputGroup("cardReason", "原因");
    content.appendChild(reasonGroup);
    
    const witnessGroup = createWitnessInputGroup("cardWitness", "見證人");
    content.appendChild(witnessGroup);
    
    const dutyGroup = createTextInputGroup("cardDuty", "當班 Duty", "請填入當天Duty本名");
    content.appendChild(dutyGroup);
    
    const errorContainer = document.createElement("div");
    errorContainer.id = "form-error-container";
    errorContainer.className = "form-error-container";
    errorContainer.style.display = "none";
    content.appendChild(errorContainer);
    
    return content;
  }

  // 創建加班表單
  function createOvertimeFormContent() {
    const content = document.createElement("div");
    content.className = "form-content";
    
    const dateGroup = createDateInputGroup("overtimeDate", "日期");
    content.appendChild(dateGroup);
    
    const overtimeTypeOptions = [
      { value: "", label: "請選擇類型" },
      { value: "平常日", label: "平常日" },
      { value: "休息日", label: "休息日" }
    ];
    const typeGroup = createSelectInputGroup("overtimeType", "類型", overtimeTypeOptions);
    content.appendChild(typeGroup);
    
    const overtimePeriodOptions = [
      { value: "", label: "請選擇時段" },
      { value: "提前上班", label: "提前上班" },
      { value: "延後下班", label: "延後下班" },
      { value: "其他", label: "其他" }
    ];
    const periodGroup = createSelectInputGroup("overtimePeriod", "時段", overtimePeriodOptions);
    content.appendChild(periodGroup);
    
    // 添加跨晝夜加班勾選框
    const crossDayGroup = document.createElement("div");
    crossDayGroup.className = "form-group";
    crossDayGroup.style.marginBottom = "15px";

    const crossDayContainer = document.createElement("div");
    crossDayContainer.className = "checkbox-group";
    crossDayContainer.style.cssText = `
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 12px 15px;
      background: rgba(52, 152, 219, 0.1);
      border: 1px solid rgba(52, 152, 219, 0.3);
      border-radius: 6px;
    `;

    crossDayContainer.innerHTML = `
      <input type="checkbox" id="overtimeCrossDay" name="overtimeCrossDay" class="checkbox-input" style="width: 28px; height: 28px; cursor: pointer; margin: 0;">
      <label for="overtimeCrossDay" style="cursor: pointer; color: #3498db; font-weight: 500; display: flex; align-items: center; margin: 0;">
        <span class="material-icons" style="margin-right: 8px; font-size: 28px;">nights_stay</span>
        跨晝夜加班
      </label>
    `;

    crossDayGroup.appendChild(crossDayContainer);
    content.appendChild(crossDayGroup);
    
    const startTimeGroup = createTimeInputGroup("overtimeStartTime", "開始時間");
    content.appendChild(startTimeGroup);
    
    const endTimeGroup = createTimeInputGroup("overtimeEndTime", "結束時間", true);
    // 在結束時間標籤後添加隔天提示
    const endTimeLabel = endTimeGroup.querySelector('.form-label');
    if (endTimeLabel) {
      const nextDayLabel = document.createElement("span");
      nextDayLabel.id = "nextDayLabel";
      nextDayLabel.style.cssText = "display: none; color: #2ecc71; font-size: 12px; margin-left: 8px;";
      nextDayLabel.textContent = "（隔天）";
      endTimeLabel.appendChild(nextDayLabel);
    }
    content.appendChild(endTimeGroup);
    
    const hoursDisplay = document.createElement("div");
    hoursDisplay.id = "overtimeHoursDisplay";
    hoursDisplay.style.marginTop = "-10px";
    hoursDisplay.style.marginBottom = "15px";
    hoursDisplay.style.padding = "8px 12px";
    hoursDisplay.style.backgroundColor = "rgba(52, 152, 219, 0.1)";
    hoursDisplay.style.borderRadius = "4px";
    hoursDisplay.style.color = "#3498db";
    hoursDisplay.style.fontSize = "13px";
    hoursDisplay.style.display = "none";
    hoursDisplay.style.textAlign = "center"; 
    hoursDisplay.innerHTML = `
      <span class="material-icons" style="font-size:16px; vertical-align:middle; margin-right:5px;">schedule</span>
      預計加班時數：<strong id="calculatedHours" style="font-weight:bold; color:#fff;">--</strong> 小時
      <br><span style="font-size:11px; color:#aaa;">(15分鐘↑進位0.5小時，45分鐘↑進位1小時)</span>
      <div id="crossDayInfo" style="display: none; margin-top: 5px; font-size: 12px; color: #ffc107;">
        <span class="material-icons" style="font-size: 14px; vertical-align: middle;">info</span>
        跨晝夜計算：結束時間為隔天
      </div>
    `;
    content.appendChild(hoursDisplay);
    
    const reasonGroup = createTextAreaInputGroup(
      "overtimeReason", 
      "原因", 
      "e.g., 21:00~21:30 協助現場收班"
    );
    content.appendChild(reasonGroup);
    
    // 使用與補卡表單相同的簡單文字輸入框
    const dutyGroup = createTextInputGroup("overtimeDuty", "當班 Duty", "請填入當天Duty本名");
    content.appendChild(dutyGroup);
    
    const errorContainer = document.createElement("div");
    errorContainer.id = "form-error-container";
    errorContainer.className = "form-error-container";
    errorContainer.style.display = "none";
    content.appendChild(errorContainer);
    
    const startTimeInput = content.querySelector('#overtimeStartTime');
    const endTimeInput = content.querySelector('#overtimeEndTime');
    const calculatedHoursSpan = content.querySelector('#calculatedHours');
    
    const updateHours = () => {
      const startVal = startTimeInput.value;
      const endVal = endTimeInput.value;
      const crossDayCheckbox = content.querySelector('#overtimeCrossDay');
      const isCrossDay = crossDayCheckbox ? crossDayCheckbox.checked : false;
      
      // 計算時數，傳入跨晝夜參數
      const hours = calculateEstimatedOvertimeHours(startVal, endVal, isCrossDay);
      
      if (hours !== null && hours >= 0) {
        calculatedHoursSpan.textContent = hours;
        hoursDisplay.style.display = "block";
      } else {
        calculatedHoursSpan.textContent = "--";
        hoursDisplay.style.display = "none";
      }
    };
    
    // 跨晝夜勾選框事件處理
    const crossDayCheckboxInput = content.querySelector('#overtimeCrossDay');
    if (crossDayCheckboxInput) {
      crossDayCheckboxInput.addEventListener('change', function() {
        const nextDayLabel = content.querySelector('#nextDayLabel');
        const crossDayInfo = content.querySelector('#crossDayInfo');
        
        if (this.checked) {
          if (nextDayLabel) nextDayLabel.style.display = 'inline';
          if (crossDayInfo) crossDayInfo.style.display = 'block';
          
          // 當勾選跨晝夜時，如果結束時間是空的或是 00:00，觸發一次更新
          if (endTimeInput && (!endTimeInput.value || endTimeInput.value === '00:00')) {
            // 強制觸發時數計算
            updateHours();
          }
        } else {
          if (nextDayLabel) nextDayLabel.style.display = 'none';
          if (crossDayInfo) crossDayInfo.style.display = 'none';
        }
        
        // 重新計算時數
        updateHours();
      });
    }
    
    // 使用原生JavaScript事件監聽器，避免jQuery依賴
    if (startTimeInput) {
      startTimeInput.addEventListener('change', updateHours);
      // 如果jQuery可用，也綁定jQuery事件（為了與jQuery UI timepicker相容）
      if (typeof $ !== 'undefined') {
        $(startTimeInput).on('change', updateHours);
      }
    }
    if (endTimeInput) {
      endTimeInput.addEventListener('change', updateHours);
      // 如果jQuery可用，也綁定jQuery事件（為了與jQuery UI timepicker相容）
      if (typeof $ !== 'undefined') {
        $(endTimeInput).on('change', updateHours);
        
        // 特殊處理：當用戶點擊時間選擇器的確認按鈕時，強制觸發更新
        $(endTimeInput).on('blur', function() {
          // 延遲執行以確保值已經更新
          setTimeout(() => {
            if (this.value) {
              updateHours();
            }
          }, 100);
        });
      }
    }
    
    return content;
  }

  // 創建請假表單
  function createLeaveRequestFormContent() {
    const content = document.createElement("div");
    content.className = "form-content";
    
    // ===== 第一區塊：假別、原因、代班狀況、通報Duty =====
    
    // 假別類型選擇
    const leaveTypeGroup = createSelectInputGroup("leaveType", "假別類型", [
      { value: "", label: "請選擇假別" },
      { value: "__separator1__", label: "─── 需要佐證文件 ───", disabled: true },
      { value: "事假", label: "事假" },
      { value: "病假", label: "病假" },
      { value: "喪假", label: "喪假" },
      { value: "婚假", label: "婚假" },
      { value: "公假", label: "公假" },
      { value: "家庭", label: "家庭" },
      { value: "__separator2__", label: "─── 不需要佐證 ───", disabled: true },
      { value: "特別休假", label: "特別休假" },
      { value: "生理假", label: "生理假" },
      { value: "國定假日", label: "國定假日" }
    ]);
    content.appendChild(leaveTypeGroup);
    
    // 創建佐證提醒區域
    const fileRequirementHint = document.createElement('div');
    fileRequirementHint.id = 'fileRequirementHint';
    fileRequirementHint.className = 'file-requirement-hint';
    fileRequirementHint.style.cssText = `
      display: none;
      background: rgba(255, 152, 0, 0.15) !important;
      border: 1px solid #ff9800 !important;
      border-radius: 8px !important;
      padding: 12px 16px !important;
      margin-top: 15px !important;
      margin-bottom: 15px !important;
      animation: fadeIn 0.3s ease-in-out !important;
      box-shadow: 0 2px 8px rgba(255, 152, 0, 0.2) !important;
    `;
    fileRequirementHint.innerHTML = `
      <p style="color: #ff9800; font-weight: 500; margin: 0; font-size: 16px; text-align: center;">
        此假別必須上傳對應佐證文件
      </p>
    `;
    content.appendChild(fileRequirementHint);
    
    // 加入 CSS 動畫
    if (!document.getElementById('fileHintStyles')) {
      const style = document.createElement('style');
      style.id = 'fileHintStyles';
      style.textContent = `
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(-10px); }
          to { opacity: 1; transform: translateY(0); }
        }
      `;
      document.head.appendChild(style);
    }
    
    // 原因輸入（加入placeholder）
    const reasonGroup = createTextAreaInputGroup("leaveReason", "原因", "請填寫事由");
    content.appendChild(reasonGroup);
    
    // 代班狀況
    const findReplacementGroup = createSelectInputGroup(
      "leaveSubstituteStatus",
      "代班狀況",
      [
        { value: "", label: "請選擇..." },
        { value: "有找到", label: "有找到" },
        { value: "找不到夥伴", label: "找不到夥伴" }
      ]
    );
    content.appendChild(findReplacementGroup);

    // 代班夥伴姓名（條件顯示）
    const replacementNameGroup = createTextInputGroup(
      "leaveSubstituteName",
      "代班夥伴本名",
      "請鍵入代班夥伴本名"
    );
    replacementNameGroup.id = "substituteNameContainer";
    replacementNameGroup.style.display = "none";
    content.appendChild(replacementNameGroup);

    // 授權 Duty
    const authorizingDutyGroup = createTextInputGroup(
      "leaveAuthDuty",
      "當班Duty",
      "請鍵入授權Duty本名"
    );
    const authDutyInput = authorizingDutyGroup.querySelector('input');
    if(authDutyInput) authDutyInput.required = true;
    content.appendChild(authorizingDutyGroup);
    
    // 第一個分隔線
    const firstSeparator = document.createElement('div');
    firstSeparator.style.cssText = `
      height: 1px;
      background-color: #444;
      margin: 20px 0;
      opacity: 0.3;
    `;
    content.appendChild(firstSeparator);
    
    // ===== 第二區塊：開始請假資訊 =====
    
    // 開始日期和時間
    const startDateGroup = createDateInputGroup("leaveStartDate", "開始日期");
    content.appendChild(startDateGroup);
    
    const startTimeGroup = createTimeInputGroup("leaveStartTime", "開始時間", false, true);
    content.appendChild(startTimeGroup);
    
    // 第二個分隔線
    const secondSeparator = document.createElement('div');
    secondSeparator.style.cssText = `
      height: 1px;
      background-color: #444;
      margin: 20px 0;
      opacity: 0.3;
    `;
    content.appendChild(secondSeparator);
    
    // ===== 第三區塊：結束請假資訊 =====
    
    // 結束日期和時間
    const endDateGroup = createDateInputGroup("leaveEndDate", "結束日期");
    content.appendChild(endDateGroup);
    
    const endTimeGroup = createTimeInputGroup("leaveEndTime", "結束時間", false, true);
    content.appendChild(endTimeGroup);
    
    // 時數顯示區域（放在結束時間後面）- 採用加班表單的緊湊風格
    const hoursDisplay = document.createElement("div");
    hoursDisplay.id = "leaveHoursDisplay";
    hoursDisplay.style.marginTop = "-10px";
    hoursDisplay.style.marginBottom = "15px";
    hoursDisplay.style.padding = "8px 12px";
    hoursDisplay.style.backgroundColor = "rgba(33, 150, 243, 0.1)";
    hoursDisplay.style.borderRadius = "4px";
    hoursDisplay.style.color = "#2196F3";
    hoursDisplay.style.fontSize = "13px";
    hoursDisplay.style.display = "none";
    hoursDisplay.style.textAlign = "center";
    hoursDisplay.innerHTML = `
      <span class="material-icons" style="font-size:16px; vertical-align:middle; margin-right:5px;">event_available</span>
      預計請假時數：<strong id="calculatedLeaveHours" style="font-weight:bold; color:#1976D2;">--</strong>
      <br><span style="font-size:11px; color:#aaa;">實際時數以審核為準</span>
    `;
    content.appendChild(hoursDisplay);
    
    // 加入假別選擇事件監聽器
    const leaveTypeSelect = leaveTypeGroup.querySelector('#leaveType');
    leaveTypeSelect.addEventListener('change', function() {
      const selectedType = this.value;
      const hint = document.getElementById('fileRequirementHint');
      
      if (LEAVE_TYPE_CONFIG.REQUIRES_FILE.includes(selectedType)) {
        hint.style.display = 'block';
      } else {
        hint.style.display = 'none';
      }
      
      // 驗證欄位
      validateField('leaveType');
    });
    
    // 移除：動態檔案上傳區塊（改為表單提交後處理）
    // const fileUploadSection = createDynamicFileUploadSection();
    // content.appendChild(fileUploadSection);
    
    const fileNote = createFileUploadNote();
    fileNote.style.display = "none"; // 隱藏舊的提示，使用新的動態區塊
    content.appendChild(fileNote);
    
    const errorContainer = document.createElement("div");
    errorContainer.id = "form-error-container";
    errorContainer.className = "form-error-container";
    errorContainer.style.display = "none";
    content.appendChild(errorContainer);
    
    const findReplacementSelect = content.querySelector('#leaveSubstituteStatus');
    const replacementNameDiv = content.querySelector('#substituteNameContainer');
    const replacementNameInput = content.querySelector('#leaveSubstituteName');

    if (findReplacementSelect && replacementNameDiv && replacementNameInput) {
      findReplacementSelect.addEventListener('change', function() {
        if (this.value === '有找到') {
          replacementNameDiv.style.display = 'block';
          replacementNameInput.required = true;
        } else {
          replacementNameDiv.style.display = 'none';
          replacementNameInput.required = false;
          replacementNameInput.value = '';
          clearFieldError('leaveSubstituteName');
        }
        validateField('leaveSubstituteStatus');
        if(this.value !== '有找到') validateField('leaveSubstituteName');
      });
    }

    // 添加時數計算功能
    const startDateInput = content.querySelector('#leaveStartDate');
    const endDateInput = content.querySelector('#leaveEndDate');
    const startTimeInput = content.querySelector('#leaveStartTime');
    const endTimeInput = content.querySelector('#leaveEndTime');
    const hoursDisplayElement = content.querySelector('#leaveHoursDisplay');
    
    // 即時驗證日期時間邏輯
    function validateDateTime() {
      const startDate = startDateInput.value;
      const endDate = endDateInput.value;
      const startTime = startTimeInput.value;
      const endTime = endTimeInput.value;
      
      // 清除之前的錯誤
      clearFieldError('leaveEndDate');
      clearFieldError('leaveEndTime');
      
      if (startDate && endDate) {
        // 如果有時間就用完整的日期時間比較，沒有就只比較日期
        const startDatetime = new Date(`${startDate} ${startTime || '00:00'}`);
        const endDatetime = new Date(`${endDate} ${endTime || '23:59'}`);
        
        if (endDatetime <= startDatetime) {
          // 判斷是日期問題還是時間問題
          const startDateOnly = new Date(startDate);
          const endDateOnly = new Date(endDate);
          
          if (endDateOnly < startDateOnly) {
            showFieldError('leaveEndDate', '結束日期不能早於開始日期');
          } else {
            // 同一天或結束日期晚於開始日期，但時間有問題
            showFieldError('leaveEndTime', '結束時間必須晚於開始時間');
          }
          return false;
        }
      }
      return true;
    }
    
    // 計算請假時數函數
    function calculateLeaveHours() {
      const startDate = startDateInput.value;
      const endDate = endDateInput.value;
      const startTime = startTimeInput.value || '09:00';  // 預設上班時間
      const endTime = endTimeInput.value || '18:00';      // 預設下班時間
      
      // 只要有日期就可以計算
      if (startDate && endDate) {
        try {
          const start = new Date(`${startDate} ${startTime}`);
          const end = new Date(`${endDate} ${endTime}`);
          
          if (end > start) {
            // 計算實際時間差異
            const diffMs = end - start;
            const diffHours = Math.round(diffMs / (1000 * 60 * 60) * 10) / 10; // 保留一位小數
            
            // 更新顯示內容 - 根據時數決定顯示格式
            const hoursValueSpan = content.querySelector('#calculatedLeaveHours');
            
            if (hoursValueSpan) {
              if (diffHours < 9) {
                // 小於9小時，顯示小時數
                hoursValueSpan.textContent = `${diffHours} 小時`;
              } else {
                // 9小時或以上，顯示天數和剩餘小時
                const days = Math.floor(diffHours / 9);
                const remainingHours = diffHours % 9;
                
                if (remainingHours === 0) {
                  // 整數天
                  hoursValueSpan.textContent = `${days} 天`;
                } else {
                  // 有剩餘小時
                  hoursValueSpan.textContent = `${days} 天 ${remainingHours} 小時`;
                }
              }
              
              hoursDisplayElement.style.display = 'block';
            }
          } else {
            hoursDisplayElement.style.display = 'none';
          }
        } catch (error) {
          console.warn('日期時間格式錯誤:', error);
          hoursDisplayElement.style.display = 'none';
        }
      } else {
        hoursDisplayElement.style.display = 'none';
      }
    }
    
    // 綁定事件監聽器
    if (startDateInput && endDateInput && startTimeInput && endTimeInput && hoursDisplayElement) {
      // 結合驗證和計算的函數（加入短暫延遲確保jQuery UI已更新值）
      function validateAndCalculate() {
        // 使用 requestAnimationFrame 而非 setTimeout，效能更好
        requestAnimationFrame(() => {
          const isValid = validateDateTime();
          if (isValid) {
            calculateLeaveHours();
          } else {
            hoursDisplayElement.style.display = 'none';
          }
        });
      }
      
      startDateInput.addEventListener('change', validateAndCalculate);
      endDateInput.addEventListener('change', validateAndCalculate);
      startTimeInput.addEventListener('change', validateAndCalculate);
      endTimeInput.addEventListener('change', validateAndCalculate);
      
      // 如果jQuery可用，也綁定jQuery事件（為了與jQuery UI timepicker相容）
      if (typeof $ !== 'undefined') {
        $(startDateInput).on('change', validateAndCalculate);
        $(endDateInput).on('change', validateAndCalculate);
        $(startTimeInput).on('change', validateAndCalculate);
        $(endTimeInput).on('change', validateAndCalculate);
        
        // 特殊處理：blur事件
        $(startTimeInput).on('blur', function() {
          setTimeout(() => {
            if (this.value || startDateInput.value) {
              validateAndCalculate();
            }
          }, 100);
        });
        
        $(endTimeInput).on('blur', function() {
          setTimeout(() => {
            if (this.value || endDateInput.value) {
              validateAndCalculate();
            }
          }, 100);
        });
      }
      
      // 添加驗證事件
      addValidationEvents(startDateInput);
      addValidationEvents(endDateInput);
      addValidationEvents(startTimeInput);
      addValidationEvents(endTimeInput);
    }
    
    return content;
  }

  // 創建班別異動表單
  function createShiftChangeFormContent() {
    const content = document.createElement("div");
    content.className = "form-content";
    
    const dateGroup = createDateInputGroup("shiftDate", "日期");
    content.appendChild(dateGroup);
    
    const teamGroup = createSelectInputGroup("teamType", "團隊", [
      { value: "", label: "請選擇團隊" },
      { value: "SV", label: "SV-外場" },
      { value: "KT", label: "KT-內廚" }
    ]);
    content.appendChild(teamGroup);
    
    const originalShiftGroup = createSelectInputGroup("originalShift", "原班別", [
      { value: "", label: "請先選擇團隊類別" }
    ]);
    content.appendChild(originalShiftGroup);
    
    const newShiftGroup = createSelectInputGroup("newShift", "新班別", [
      { value: "", label: "請先選擇團隊類別" }
    ]);
    content.appendChild(newShiftGroup);
    
    const shiftWarning = document.createElement("div");
    shiftWarning.id = "shiftWarning";
    shiftWarning.style.display = "none";
    shiftWarning.style.color = "#ff4444";
    shiftWarning.style.fontSize = "14px";
    shiftWarning.style.marginTop = "5px";
    shiftWarning.style.marginBottom = "10px";
    shiftWarning.style.padding = "10px";
    shiftWarning.style.borderRadius = "5px";
    shiftWarning.style.backgroundColor = "rgba(255, 68, 68, 0.1)";
    shiftWarning.style.textAlign = "center";
    shiftWarning.innerHTML = "<strong>系統提示：</strong>原班別與新班別相同｜請選擇不同班別";
    content.appendChild(shiftWarning);
    
    const reasonGroup = createTextAreaInputGroup("shiftReason", "原因");
    content.appendChild(reasonGroup);
    
    const errorContainer = document.createElement("div");
    errorContainer.id = "form-error-container";
    errorContainer.className = "form-error-container";
    errorContainer.style.display = "none";
    content.appendChild(errorContainer);
    
    const teamSelect = content.querySelector('#teamType');
    const originalSelect = content.querySelector('#originalShift');
    const newSelect = content.querySelector('#newShift');
    
    if (teamSelect && originalSelect && newSelect) {
      const shiftOptions = {
        'SV': [
          { value: "", label: "請選擇班別" }, { value: "A1", label: "A1" }, { value: "A2", label: "A2" }, { value: "A3", label: "A3" },
          { value: "B1", label: "B1" }, { value: "B2", label: "B2" }, { value: "B3", label: "B3" },
          { value: "X1", label: "X1" }, { value: "X2", label: "X2" }, { value: "X3", label: "X3" }
        ],
        'KT': [
          { value: "", label: "請選擇班別" }, { value: "KA", label: "KA" }, { value: "KB", label: "KB" }
        ]
      };
      
      teamSelect.addEventListener('change', function() {
        const selectedTeam = this.value;
        while (originalSelect.options.length > 0) { originalSelect.remove(0); }
        while (newSelect.options.length > 0) { newSelect.remove(0); }
        shiftWarning.style.display = "none";
        if (selectedTeam && shiftOptions[selectedTeam]) {
          const options = shiftOptions[selectedTeam];
          options.forEach(opt => {
            originalSelect.add(new Option(opt.label, opt.value));
            newSelect.add(new Option(opt.label, opt.value));
          });
          clearFieldError('originalShift');
          clearFieldError('newShift');
        } else {
          originalSelect.add(new Option("請先選擇團隊類別", ""));
          newSelect.add(new Option("請先選擇團隊類別", ""));
        }
      });
      
      function checkDuplicateShift() {
        const originalValue = originalSelect.value;
        const newValue = newSelect.value;
        if (originalValue && newValue && originalValue === newValue) {
          shiftWarning.style.display = "block";
          showFieldError('newShift');
        } else {
          shiftWarning.style.display = "none";
          clearFieldError('newShift');
        }
      }
      
      originalSelect.addEventListener('change', checkDuplicateShift);
      newSelect.addEventListener('change', checkDuplicateShift);
    }
    
    return content;
  }

  // 表單提交函式

  // 創建表單提交函式工廠
  function createFormSubmitters(dependencies) {
    const { elements, UIUtils, NetworkUtils, showUploadPrompt } = dependencies;

    // 補卡申請提交
    async function submitCardCorrection(formData, callback) {
      const idCode = elements.idInput.value;
      try {
        // 1. 先進行驗證
        UIUtils.displayLoading("驗證中｜請稍候...");
        const authResult = await NetworkUtils.callGASWithRetry('checkDeviceBinding', [idCode, "補卡申請提交"]);
        if (authResult.status !== "success") {
          callback(false, authResult.message || "驗證失敗", { stage: 'auth' });
          return;
        }
        
        // 2. 驗證通過後提交表單
        UIUtils.displayLoading("處理中｜請稍候...");
        const apiParams = [ idCode, formData.cardDate, formData.cardTime, formData.cardAction, formData.cardType, formData.cardReason, formData.cardWitness, formData.cardDuty ];
        const result = await NetworkUtils.callGASWithRetry("submitCardCorrection", apiParams);
        if (result.status === "success") {
          callback(true, result.message || `補卡申請已成功提交`, { formType: 'card', needFile: false });
        } else {
          callback(false, result.message || "提交失敗", { stage: 'api' });
        }
      } catch (error) {
        callback(false, error.message || "處理失敗", { stage: 'api' });
      }
    }

    // 加班申請提交
    async function submitOvertimeRequest(formData, callback) {
      const idCode = elements.idInput.value;
      try {
        // 1. 先進行驗證
        UIUtils.displayLoading("驗證中｜請稍候...");
        const authResult = await NetworkUtils.callGASWithRetry('checkDeviceBinding', [idCode, "加班申請提交"]);
        if (authResult.status !== "success") {
          callback(false, authResult.message || "驗證失敗", { stage: 'auth' });
          return;
        }
        
        // 2. 驗證通過後提交表單
        UIUtils.displayLoading("處理中｜請稍候...");
        const apiParams = [ idCode, formData.overtimeDate, formData.overtimeType, formData.overtimePeriod, formData.overtimeStartTime, formData.overtimeEndTime, formData.overtimeReason, formData.overtimeDuty ];
        const result = await NetworkUtils.callGASWithRetry("submitOvertimeRequest", apiParams);
        if (result.status === "success") {
          callback(true, result.message || `加班申請已成功提交`, { formType: 'overtime', needFile: false });
        } else {
          callback(false, result.message || "提交失敗", { stage: 'api' });
        }
      } catch (error) {
        callback(false, error.message || "處理失敗", { stage: 'api' });
      }
    }

    // 請假申請提交
    async function submitLeaveRequest(formData, callback) {
      const idCode = elements.idInput.value;
      try {
        // 1. 先進行驗證
        UIUtils.displayLoading("驗證中｜請稍候...");
        const authResult = await NetworkUtils.callGASWithRetry('checkDeviceBinding', [idCode, "請假申請提交"]);
        if (authResult.status !== "success") {
          callback(false, authResult.message || "驗證失敗", { stage: 'auth' });
          return;
        }

        // 2. 驗證成功後進行表單提交
        UIUtils.displayLoading("處理中｜請稍候...");
        // 加入時間參數
        const apiParams = [ 
          idCode, 
          formData.leaveStartDate, 
          formData.leaveEndDate, 
          formData.leaveStartTime,
          formData.leaveEndTime,
          formData.leaveType, 
          formData.leaveReason, 
          formData.leaveSubstituteStatus, 
          formData.leaveSubstituteName || "", 
          formData.leaveAuthDuty 
        ];
        const result = await NetworkUtils.callGASWithRetry("submitLeaveRequest", apiParams);
        if (result.status === "success") {
          // 檢查是否需要上傳檔案
          const needsUpload = requiresFileUpload(formData.leaveType);
          
          if (needsUpload && result.uploadUrl && result.requestId) {
            // 需要上傳檔案的假別，顯示上傳提示（不關閉 Modal）
            showUploadPrompt(result.requestId, result.uploadUrl);
            callback(true, "請假申請已提交｜請完成檔案上傳", { 
              stage: 'text_submitted', 
              requestId: result.requestId,
              needsUpload: true,
              keepModalOpen: true  // 告訴回調不要關閉 Modal
            });
          } else {
            // 不需要上傳檔案的假別，正常回調（會關閉 Modal）
            callback(true, result.message || "請假申請已成功提交", { 
              stage: 'completed',
              needsUpload: false,
              keepModalOpen: false
            });
          }
        } else {
          callback(false, result.message || "提交失敗", { stage: 'api' });
        }
      } catch (error) {
        callback(false, error.message || "處理失敗", { stage: 'api' });
      }
    }

    // 班別異動提交
    async function submitShiftChange(formData, callback) {
      const idCode = elements.idInput.value;
      try {
        // 1. 先進行驗證
        UIUtils.displayLoading("驗證中｜請稍候...");
        const authResult = await NetworkUtils.callGASWithRetry('checkDeviceBinding', [idCode, "班別異動提交"]);
        if (authResult.status !== "success") {
          callback(false, authResult.message || "驗證失敗", { stage: 'auth' });
          return;
        }

        // 2. 驗證成功後進行表單提交
        UIUtils.displayLoading("處理中｜請稍候...");
        const apiParams = [ idCode, formData.shiftDate, formData.teamType, formData.originalShift, formData.newShift, formData.shiftReason ];
        const result = await NetworkUtils.callGASWithRetry("submitShiftChange", apiParams);
        if (result.status === "success") {
          callback(true, result.message || `班別異動申請已成功提交`, { formType: 'shift', needFile: false });
        } else {
          callback(false, result.message || "提交失敗", { stage: 'api' });
        }
      } catch (error) {
        callback(false, error.message || "處理失敗", { stage: 'api' });
      }
    }

    return {
      submitCardCorrection,
      submitOvertimeRequest,
      submitLeaveRequest,
      submitShiftChange
    };
  }

  // 導出表單處理模組
  App.FormHandler = {
    // 假別分類配置
    LEAVE_TYPE_CONFIG,
    
    // 驗證函式
    validateCardCorrectionFields,
    validateOvertimeFields,
    validateLeaveFields,
    validateShiftChangeFields,
    validateField,
    showFieldError,
    clearFieldError,
    addValidationEvents,
    
    // 工具函式
    calculateEstimatedOvertimeHours,
    calculateEstimatedLeaveDays,
    
    // 表單元素創建函式
    createLabelWithRequired,
    createTextInputGroup,
    createDateInputGroup,
    createTimeInputGroup,
    createSelectInputGroup,
    createTextAreaInputGroup,
    createActionButtonGroup,
    createWitnessInputGroup,
    createDateRangeInputGroup,
    createDateTimeRangeInputGroup,
    createTimeInputElement,
    createFileUploadNote,
    
    // 表單內容創建函式
    createCardCorrectionFormContent,
    createOvertimeFormContent,
    createLeaveRequestFormContent,
    createShiftChangeFormContent,
    
    // 表單提交函式工廠（需要依賴注入）
    createFormSubmitters
    
    // 移除配置導出，採用自包含配置策略確保載入穩定性
  };

})(window.App); 
